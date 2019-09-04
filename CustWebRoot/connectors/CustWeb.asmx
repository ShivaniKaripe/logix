<%@ WebService Language="VB" Class="Service" %>
Imports System
Imports System.Data
Imports System.IO
Imports System.Web
Imports System.Web.Services
Imports System.Xml
Imports System.Xml.Schema
Imports System.Xml.Serialization
Imports System.Xml.Xsl
Imports System.Xml.XPath
Imports System.Data.SqlClient
Imports System.Data.Odbc
Imports System.String
Imports System.Net.NetworkInformation
Imports Copient.CommonInc
Imports Copient.LogixInc
Imports Copient.AlternateID
Imports Copient.CustomerLookup
Imports Copient.ConnectorInc
Imports Copient.ExternalRewards
Imports Copient.commonShared
Imports Copient.Customer
Imports Copient.Card
Imports Copient.CryptLib

<WebService(Namespace:="http://www.copienttech.com/CustomerFacingWebsite/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class Service
  Inherits System.Web.Services.WebService
  ' 
    Private Const bDebugLogOn As Boolean = False
 
    Private MyCommon As New Copient.CommonInc
    Private logixInc As New Copient.LogixInc
    Private MyAltID As New Copient.AlternateID
    Private MyCryptLib As New Copient.CryptLib()
    
    Private DebugStartTimes As New ArrayList()
    Private Const sLogFileName As String = "CustWeb"
    Private Const sAppName As String = "CustWeb"
    Private sInstallationName As String = ""
    Private sLogLines As String = ""
    Private Const scDashes As String = "-----"
    Private sInputForLog As String = ""
    Dim bEnableFilterForResponse As Boolean = IIf(MyCommon.Fetch_SystemOption(317) = "1", True, False)
	Dim bDisableGraphicFileSearch As Boolean = IIf(MyCommon.Fetch_SystemOption(332) = "1", True, False)
    Dim AddHHToGroup As Boolean = False

    Private Enum DebugState
        BeginTime = -1
        CurrentTime = 0
        EndTime = 1
    End Enum
  
    Private Enum MessageType
        Info = 0
        Warning = 1
        AppError = 2
        SysError = 3
        Debug = 4
    End Enum

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
        FAILID_NO_EXT_PROGRAM = 12
        INVALID_PRODUCTID = 13
        INVALID_PRODUCTTYPE = 14
        NOTFOUND_PRODUCTGROUP = 15
        NOTFOUND_RECORDS = 16
        INVALID_CUSTSUPPLEMENTATTRIBUTE = 17
        INVALID_SUPPATTRIBUTEVALUE = 18
        NOTFOUND_CUSTSUPPLEMENTATTRIBUTE = 19
	INVALID_CUSTOMERGROUPNAME = 20
        INVALID_EXTLOCATIONCODE = 20
        INVALID_STARTDATE = 21
        INVALID_ENDDATE = 22
        INVALID_LAST4_CARDID = 23
        INVALID_TRANSNUM = 24
        NOTFOUND_TRANSACTIONS = 25
        NOTFOUND_TRANSACTIONITEMS = 26
        FAILED_TOOMANYRECORDS = 27
        APPLICATION_EXCEPTION = 9999
    End Enum

    Public Class CustWebStatus
        Public StatusCode As String = "0"
        Public Description As String = "Success"
    End Class

    <WebMethod()> _
    Public Function CustomerDetails(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As String) As System.Data.DataSet
        Dim iCustomerTypeID As Integer = -1
        Try
            InitApp()
            iCustomerTypeID = CInt(CustomerTypeID)
        Catch ex As Exception
            iCustomerTypeID = -1
        End Try
    If (bEnableFilterForResponse) Then
        Return _CustomerDetailsFilteredResponse(GUID, CustomerID, iCustomerTypeID)
    Else
        Return _CustomerDetails(GUID, CustomerID, iCustomerTypeID)
    End If
    End Function

    Public Class CustWebCustomer
        Public CustomerPK As Long = 0
        Public FirstName As String = ""
        Public LastName As String = ""
        Public Employee As String = "false"
        Public CustomerStatusID As String = ""
        Public CurrentYearSTD As Decimal = 0.0
        Public LastYearSTD As Decimal = 0.0
        Public CustomerTypeID As String = "0"
        Public HouseholdPK As Long = 0
        Public Address As String = ""
        Public City As String = ""
        Public State As String = ""
        Public Zip As String = ""
        Public Country As String = ""
        Public Phone As String = ""
        Public CardID As String = ""
        Public MiddleName As String = ""
        Public Mobile As String = ""
        Public Email As String = ""
    
    End Class
  
    Public Class CustomerPointsProgram
        Public ID As String = ""
        Public Balance As Long = 0
    End Class

    Public Class CustomerDetailsClass
        Public Status As CustWebStatus
        Public Customer As CustWebCustomer
        Public PointsPrograms() As CustomerPointsProgram
    End Class
    
    
    'This WebMethod returns the customerdetails given GUID, CardId, CustomerTypeID 
    <WebMethod()> _
    Public Function CustomerDetails_Class(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As Integer) As CustomerDetailsClass
        Dim ds As System.Data.DataSet
        Dim oCustomerDetails As New CustomerDetailsClass
        Dim oStatus As New CustWebStatus
        Dim dt As DataTable
        Dim dr As DataRow
        Dim i As Integer
      
        InitApp()
    
        ds = _CustomerDetails(GUID, CustomerID, CustomerTypeID)

        dt = ds.Tables("Status")
        oStatus.StatusCode = dt.Rows(0).Item("StatusCode")
        oStatus.Description = dt.Rows(0).Item("Description")
        oCustomerDetails.Status = oStatus
    
        If oStatus.StatusCode = "0" Then
            dt = ds.Tables("Customer")
            If dt.Rows.Count > 0 Then
                Dim oCustomer As New CustWebCustomer
                oCustomer.CustomerPK = dt.Rows(0).Item("CustomerPK")
                oCustomer.FirstName = MyCommon.NZ(dt.Rows(0).Item("FirstName"), "")
                oCustomer.LastName = MyCommon.NZ(dt.Rows(0).Item("LastName"), "")
                oCustomer.Employee = MyCommon.NZ(dt.Rows(0).Item("Employee").ToString.ToLower, "")
                oCustomer.CustomerStatusID = MyCommon.NZ(dt.Rows(0).Item("CustomerStatusID"), 0)
                oCustomer.CurrentYearSTD = MyCommon.NZ(dt.Rows(0).Item("CurrYearSTD"), 0)
                oCustomer.LastYearSTD = MyCommon.NZ(dt.Rows(0).Item("LastYearSTD"), 0)
                oCustomer.CustomerTypeID = MyCommon.NZ(dt.Rows(0).Item("CustomerTypeID"), 0)
                oCustomer.HouseholdPK = MyCommon.NZ(dt.Rows(0).Item("HHPK"), 0)
                oCustomer.Address = MyCommon.NZ(dt.Rows(0).Item("Address"), "")
                oCustomer.City = MyCommon.NZ(dt.Rows(0).Item("City"), "")
                oCustomer.State = MyCommon.NZ(dt.Rows(0).Item("State"), "")
                oCustomer.Zip = MyCommon.NZ(dt.Rows(0).Item("Zip"), "")
                oCustomer.Country = MyCommon.NZ(dt.Rows(0).Item("Country"), "")
                oCustomer.Phone = MyCommon.NZ(dt.Rows(0).Item("Phone"), "")
                oCustomer.CardID = MyCommon.NZ(dt.Rows(0).Item("InitialCardID"), "")
                oCustomer.MiddleName = MyCommon.NZ(dt.Rows(0).Item("MiddleName"), "")
                oCustomer.Mobile = MyCommon.NZ(dt.Rows(0).Item("Mobile"), "")
                oCustomer.Email = MyCommon.NZ(dt.Rows(0).Item("Email"), "")
                oCustomerDetails.Customer = oCustomer
            End If
            dt = ds.Tables("PointsProgram")
            If dt.Rows.Count > 0 Then
                ReDim oCustomerDetails.PointsPrograms(dt.Rows.Count - 1)
                Dim pp As CustomerPointsProgram
                i = 0
                For Each dr In dt.Rows
                    pp = New CustomerPointsProgram
                    pp.ID = dr.Item("ProgramID")
                    pp.Balance = dr.Item("Balance")
                    oCustomerDetails.PointsPrograms(i) = pp
                    i += 1
                Next
            End If
        End If

        Return oCustomerDetails
    End Function
  
    
    <WebMethod()> _
    Public Function CustomerDetails_Class_ByCardID(ByVal GUID As String, ByVal CustomerID As String, ByVal CardTypeID As Integer, ByVal CustomerTypeID As Integer) As CustomerDetailsClass
        Dim ds As System.Data.DataSet
        Dim oCustomerDetails As New CustomerDetailsClass
        Dim oStatus As New CustWebStatus
        Dim dt As DataTable
        Dim dr As DataRow
        Dim i As Integer

        InitApp()
    
        ds = _CustomerDetails(GUID, CustomerID, CustomerTypeID, CardTypeID)

        dt = ds.Tables("Status")
        oStatus.StatusCode = dt.Rows(0).Item("StatusCode")
        oStatus.Description = dt.Rows(0).Item("Description")
        oCustomerDetails.Status = oStatus
    
        If oStatus.StatusCode = "0" Then
            dt = ds.Tables("Customer")
            If dt.Rows.Count > 0 Then
                Dim oCustomer As New CustWebCustomer
                oCustomer.CustomerPK = dt.Rows(0).Item("CustomerPK")
                oCustomer.FirstName = MyCommon.NZ(dt.Rows(0).Item("FirstName"), "")
                oCustomer.LastName = MyCommon.NZ(dt.Rows(0).Item("LastName"), "")
                oCustomer.Employee = MyCommon.NZ(dt.Rows(0).Item("Employee").ToString.ToLower, "")
                oCustomer.CustomerStatusID = MyCommon.NZ(dt.Rows(0).Item("CustomerStatusID"), 0)
                oCustomer.CurrentYearSTD = MyCommon.NZ(dt.Rows(0).Item("CurrYearSTD"), 0)
                oCustomer.LastYearSTD = MyCommon.NZ(dt.Rows(0).Item("LastYearSTD"), 0)
                oCustomer.CustomerTypeID = MyCommon.NZ(dt.Rows(0).Item("CustomerTypeID"), 0)
                oCustomer.HouseholdPK = MyCommon.NZ(dt.Rows(0).Item("HHPK"), 0)
                oCustomer.Address = MyCommon.NZ(dt.Rows(0).Item("Address"), "")
                oCustomer.City = MyCommon.NZ(dt.Rows(0).Item("City"), "")
                oCustomer.State = MyCommon.NZ(dt.Rows(0).Item("State"), "")
                oCustomer.Zip = MyCommon.NZ(dt.Rows(0).Item("Zip"), "")
                oCustomer.Country = MyCommon.NZ(dt.Rows(0).Item("Country"), "")
                oCustomer.Phone = MyCommon.NZ(dt.Rows(0).Item("Phone"), "")
                oCustomer.CardID = MyCommon.NZ(dt.Rows(0).Item("InitialCardID"), "")
                oCustomerDetails.Customer = oCustomer
            End If
            dt = ds.Tables("PointsProgram")
            If dt.Rows.Count > 0 Then
                ReDim oCustomerDetails.PointsPrograms(dt.Rows.Count - 1)
                Dim pp As CustomerPointsProgram
                i = 0
                For Each dr In dt.Rows
                    pp = New CustomerPointsProgram
                    pp.ID = dr.Item("ProgramID")
                    pp.Balance = dr.Item("Balance")
                    oCustomerDetails.PointsPrograms(i) = pp
                    i += 1
                Next
            End If
        End If

        Return oCustomerDetails
    End Function
   

 Private Function _CustomerDetailsFilteredResponse(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As Integer, Optional ByVal CardTypeID As Integer = -1) As System.Data.DataSet
        Dim dt, dt2, dtCustAttrIDs, dtCustAttrDesc As System.Data.DataTable
        Dim dtStatus, dtBalances, dtCustomerAttributes As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim ResultSet As New System.Data.DataSet("CustomerDetails")
        Dim CustomerPK As Long = 0
        Dim BalRetMsg As String = ""
        Dim BalRetCode As StatusCodes = StatusCodes.SUCCESS
        Dim AltColumn As String = "Not Set"
        Dim CMInstalled As Boolean = False

        WriteDebug("CustomerDetails - CardNumber: " & Copient.MaskHelper.MaskCard(CustomerID, CardTypeID) & ". CardTypeID: " & CardTypeID, DebugState.BeginTime)

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        If GUID = "" Then
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_GUID
            row.Item("Description") = "Failure: Invalid GUID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _CustomerDetailsFilteredResponse = ResultSet
            Exit Function
        End If
        If CustomerID = "" Then
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERID
            row.Item("Description") = "Failure: Invalid Customer ID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _CustomerDetailsFilteredResponse = ResultSet
            Exit Function
        End If
        If CustomerTypeID = -1 Then
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERTYPEID
            row.Item("Description") = "Failure: Invalid Customer Type"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _CustomerDetailsFilteredResponse = ResultSet
            Exit Function
        End If
        If GUID.Contains(Chr(34)) = True Or GUID.Contains("'") = True Then
            row = dtStatus.NewRow()
            row.Item("StatusCode") = 1
            row.Item("Description") = "Failure: Invalid GUID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _CustomerDetailsFilteredResponse = ResultSet
            Exit Function
        End If
        If CustomerID.Contains(Chr(34)) = True Or CustomerID.Contains("'") = True Then
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERID
            row.Item("Description") = "Failure: Invalid Customer ID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _CustomerDetailsFilteredResponse = ResultSet
            Exit Function
        End If

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
            WriteDebug("Connected LogixRT and LogixXS CustomerDetails", DebugState.CurrentTime)
      
            If Not IsValidGUID(GUID) Then
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
            ElseIf (CustomerID.Length < 1) Then
                'Bad customer ID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERID
                row.Item("Description") = "Failure: Invalid customer ID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
            Else
                'First test the CustomerTypeID for alternate column 
                CMInstalled = MyCommon.IsEngineInstalled(0)
                If (CMInstalled) And (CustomerTypeID = 4) Then
                       
                    'find the alternate search column via systemobjects
                    AltColumn = MyCommon.Fetch_SystemOption(60)
                    Dim split As String() = AltColumn.Split(New [Char]() {"."c})

                    'Find the Customer PK
                    MyCommon.QueryStr = "select " & split(0) & ".CustomerPK from " & split(0) & " with (NoLock) " & _
                                        "where " & split(1) & "='" &  MyCryptLib.SQL_StringEncrypt(CustomerID) & "';"

                Else
                    'Pad the customer ID
                    If CardTypeID <> -1 Then
                        CustomerID = MyCommon.Pad_ExtCardID(CustomerID, CardTypeID)
                    Else
                        CustomerID = MyCommon.Pad_ExtCardID(CustomerID, CardTypes.CUSTOMER)
                    End If
                    'Find the Customer PK
                    WriteDebug("Starting Find of CustomerPK [LogixXS]", DebugState.CurrentTime)
                    MyCommon.QueryStr = "select CID.CustomerPK from CardIDs as CID with (NoLock) " & _
                                        "INNER JOIN CardTypes AS CT with (NoLock) ON CID.CardTypeID = CT.CardTypeID " & _
                                        "WHERE CT.CustTypeID=" & CustomerTypeID & " AND CID.ExtCardID='" & MyCryptLib.SQL_StringEncrypt(CustomerID) & "';"
                End If

                dt = MyCommon.LXS_Select
                WriteDebug("Finished Find of CustomerPK with row count=" & dt.Rows.Count, DebugState.CurrentTime)
                If dt.Rows.Count = 0 Then
                    'Customer not found
                    If CustomerTypeID = 0 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_CUSTOMER
                        row.Item("Description") = "Failure: Customer " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    ElseIf CustomerTypeID = 1 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_HOUSEHOLD
                        row.Item("Description") = "Failure: Household " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    ElseIf CustomerTypeID = 2 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_CAM
                        row.Item("Description") = "Failure: CAM " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    Else
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERTYPEID
                        row.Item("Description") = "Failure: Invalid customer type."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    End If
                Else
                    
                    'Constructing the select columns which should be sent to SP as per user filter
                    Dim CustomerDtColumns As String()
                    Dim PointsProgramDtColumns As String()
                    Dim CustomerAttributesDtColumns As String()
                    MyCommon.QueryStr = "SELECT FilteredAttributes FROM FilterOutputColumns with (NoLock) WHERE ReferenceId = 2"
                    Dim fdt As DataTable = MyCommon.LRT_Select
                    If (fdt.Rows.Count > 0) Then
                        For Each row2 In fdt.Rows
                            Dim filterstr As String = MyCommon.NZ(row2.Item("FilteredAttributes"), "")
                            Dim sArr As String() = filterstr.Split("|")
                            For Each s As String In sArr
                                Dim sArr1 As String() = s.Split("-")
                                Dim tableName As String = sArr1(0)
                                Dim selectedColumns As String() = sArr1(1).Split(",")
                                Select Case (tableName)
                                    Case "Customer"
                                        CustomerDtColumns = selectedColumns
                                    Case "PointsProgram"
                                        PointsProgramDtColumns = selectedColumns
                                    Case "CustomerAttributes"
                                        CustomerAttributesDtColumns = selectedColumns
                                    Case Else
                                End Select
                            Next
                        Next
                    End If
                    
                    CustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)
                    'Get the customer's details
                    WriteDebug("Starting Get of Customer Details [LogixXS]", DebugState.CurrentTime)
                    If (MyCommon.Extract_Val(MyCommon.Fetch_SystemOption(6)) >= 2) And (CustomerTypeID = 0) Then
                        'Savings To Date is turned on for household level - we need to return the household's STD along with the customer's STD
                        If (CMInstalled) And (CustomerTypeID = 4) Then
                            MyCommon.QueryStr = "select " & _
                                IIf(CustomerDtColumns.Contains("CustomerPK"), "C.CustomerPK,", "") & _
                                 IIf(CustomerDtColumns.Contains("FirstName"), "C.FirstName,", "") & _
								                 IIf(CustomerDtColumns.Contains("MiddleName"), " C.MiddleName,", "") & _
                                  IIf(CustomerDtColumns.Contains("LastName"), " C.LastName,", "") & _
                                   IIf(CustomerDtColumns.Contains("Employee"), "C.Employee,", "") & _
                                    IIf(CustomerDtColumns.Contains("CustomerStatusID"), "C.CustomerStatusID,", "") & _
                                     IIf(CustomerDtColumns.Contains("CurrYearSTD"), "C.CurrYearSTD,", "") & _
                                      IIf(CustomerDtColumns.Contains("LastYearSTD"), "C.LastYearSTD,", "") & _
                                       IIf(CustomerDtColumns.Contains("PrimaryCardID"), " cid.ExtCardID as PrimaryCardID,", "") & _
                                        IIf(CustomerDtColumns.Contains("HHCurrYearSTD"), "isnull(HHCust.CurrYearSTD, 0) as HHCurrYearSTD,", "") & _
                                         IIf(CustomerDtColumns.Contains("HHLastYearSTD"), " isnull(HHCust.LastYearSTD, 0) as HHLastYearSTD,", "") & _
                                          IIf(CustomerDtColumns.Contains("CustomerTypeID"), "C.CustomerTypeID,", "") & _
                                            IIf(CustomerDtColumns.Contains("HHPK") AndAlso Not CustomerDtColumns.Contains("HHID"), "C.HHPK,", "") & _
                                             IIf(CustomerDtColumns.Contains("Address"), "E.Address,", "") & _
                                              IIf(CustomerDtColumns.Contains("City"), "E.City,", "") & _
                                               IIf(CustomerDtColumns.Contains("State"), "E.State,", "") & _
                                                IIf(CustomerDtColumns.Contains("Zip"), "E.Zip,", "") & _
                                                 IIf(CustomerDtColumns.Contains("Country"), "E.Country,", "") & _
                                                  IIf(CustomerDtColumns.Contains("Phone"), "E.PhoneAsEntered as Phone,", "") & _
                        												   IIf(CustomerDtColumns.Contains("Mobile"), "E.MobilePhoneAsEntered as Mobile,", "") & _
                        												    IIf(CustomerDtColumns.Contains("Email"), "E.Email,", "") & _
                                                    IIf(CustomerDtColumns.Contains("DOB"), "E.DOB,", "") & _
                                                     IIf(CustomerDtColumns.Contains("HHID"), "'' as HHID, C.HHPK,", "") & _
                                                     IIf(CustomerDtColumns.Contains("CustomerID"), "'" & CustomerID & "' as CustomerID, ", "") & _
                                                      IIf(CustomerDtColumns.Contains("InitialCardID"), "C.InitialCardID ", "")
                            MyCommon.QueryStr = MyCommon.QueryStr.Trim(",") & " from Customers as C inner join CardIDs as cid on C.CustomerPK = cid.CustomerPK "
                            If (CustomerDtColumns.Contains("Address") OrElse CustomerDtColumns.Contains("City") OrElse CustomerDtColumns.Contains("State") OrElse CustomerDtColumns.Contains("Zip") OrElse CustomerDtColumns.Contains("Country") OrElse CustomerDtColumns.Contains("Phone") OrElse CustomerDtColumns.Contains("Mobile") OrElse CustomerDtColumns.Contains("Email") OrElse CustomerDtColumns.Contains("DOB")) Then
                                MyCommon.QueryStr = MyCommon.QueryStr & " left join CustomerEXT as E with (NoLock) on C.CustomerPK=E.CustomerPK "
                            End If
                            If (CustomerDtColumns.Contains("HHCurrYearSTD") OrElse CustomerDtColumns.Contains("HHLastYearSTD")) Then
                                MyCommon.QueryStr = MyCommon.QueryStr & " Left join Customers as HHCust with (NoLock) on C.HHPK=HHCust.CustomerPK "
                            End If
                            MyCommon.QueryStr = MyCommon.QueryStr & " where cid.cardtypeid = 0 and C.CustomerPK=" & CustomerPK.ToString & ";"
                        Else
                            MyCommon.QueryStr = "select " & _
                                IIf(CustomerDtColumns.Contains("CustomerPK"), "C.CustomerPK,", "") & _
                                 IIf(CustomerDtColumns.Contains("FirstName"), "C.FirstName,", "") & _
								                 IIf(CustomerDtColumns.Contains("MiddleName"), " C.MiddleName,", "") & _
                                  IIf(CustomerDtColumns.Contains("LastName"), " C.LastName,", "") & _
                                   IIf(CustomerDtColumns.Contains("Employee"), "C.Employee,", "") & _
                                    IIf(CustomerDtColumns.Contains("CustomerStatusID"), "C.CustomerStatusID,", "") & _
                                     IIf(CustomerDtColumns.Contains("CurrYearSTD"), "C.CurrYearSTD,", "") & _
                                      IIf(CustomerDtColumns.Contains("LastYearSTD"), "C.LastYearSTD,", "") & _
                                       IIf(CustomerDtColumns.Contains("PrimaryCardID"), " C.InitialCardID as PrimaryCardID,", "") & _
                                        IIf(CustomerDtColumns.Contains("HHCurrYearSTD"), "isnull(HHCust.CurrYearSTD, 0) as HHCurrYearSTD,", "") & _
                                         IIf(CustomerDtColumns.Contains("HHLastYearSTD"), " isnull(HHCust.LastYearSTD, 0) as HHLastYearSTD,", "") & _
                                          IIf(CustomerDtColumns.Contains("CustomerTypeID"), "C.CustomerTypeID,", "") & _
                                           IIf(CustomerDtColumns.Contains("HHPK") AndAlso Not CustomerDtColumns.Contains("HHID"), "C.HHPK,", "") & _
                                             IIf(CustomerDtColumns.Contains("Address"), "E.Address,", "") & _
                                              IIf(CustomerDtColumns.Contains("City"), "E.City,", "") & _
                                               IIf(CustomerDtColumns.Contains("State"), "E.State,", "") & _
                                                IIf(CustomerDtColumns.Contains("Zip"), "E.Zip,", "") & _
                                                 IIf(CustomerDtColumns.Contains("Country"), "E.Country,", "") & _
                                                  IIf(CustomerDtColumns.Contains("Phone"), "E.PhoneAsEntered as Phone,", "") & _
												                          IIf(CustomerDtColumns.Contains("Mobile"), "E.MobilePhoneAsEntered as Mobile,", "") & _
                                                   IIf(CustomerDtColumns.Contains("Email"), "E.Email,", "") & _
                                                    IIf(CustomerDtColumns.Contains("DOB"), "E.DOB,", "") & _
                                                     IIf(CustomerDtColumns.Contains("HHID"), "'' as HHID, C.HHPK,", "") & _
                                                     IIf(CustomerDtColumns.Contains("CustomerID"), "'" & CustomerID & "' as CustomerID,", "") & _
                                                       IIf(CustomerDtColumns.Contains("InitialCardID"), "C.InitialCardID ", "")
                            
                            MyCommon.QueryStr = MyCommon.QueryStr.Trim(",") & "from Customers as C with (NoLock) "
                            If (CustomerDtColumns.Contains("Address") OrElse CustomerDtColumns.Contains("City") OrElse CustomerDtColumns.Contains("State") OrElse CustomerDtColumns.Contains("Zip") OrElse CustomerDtColumns.Contains("Country") OrElse CustomerDtColumns.Contains("Phone") OrElse CustomerDtColumns.Contains("Mobile") OrElse CustomerDtColumns.Contains("Email") OrElse CustomerDtColumns.Contains("DOB")) Then
                                MyCommon.QueryStr = MyCommon.QueryStr & " left join CustomerEXT as E with (NoLock) on C.CustomerPK=E.CustomerPK "
                            End If
                            If (CustomerDtColumns.Contains("HHCurrYearSTD") OrElse CustomerDtColumns.Contains("HHLastYearSTD")) Then
                                MyCommon.QueryStr = MyCommon.QueryStr & " Left join Customers as HHCust with (NoLock) on C.HHPK=HHCust.CustomerPK "
                            End If
                            MyCommon.QueryStr = MyCommon.QueryStr & " where C.CustomerPK=" & CustomerPK.ToString & ";"
                        End If
                    Else
                        If (CMInstalled) And (CustomerTypeID = 4) Then
                            MyCommon.QueryStr = "select " & _
                               IIf(CustomerDtColumns.Contains("CustomerPK"), "C.CustomerPK,", "") & _
                                IIf(CustomerDtColumns.Contains("FirstName"), "C.FirstName,", "") & _
								                 IIf(CustomerDtColumns.Contains("MiddleName"), " C.MiddleName,", "") & _
                                 IIf(CustomerDtColumns.Contains("LastName"), " C.LastName,", "") & _
                                  IIf(CustomerDtColumns.Contains("Employee"), "C.Employee,", "") & _
                                   IIf(CustomerDtColumns.Contains("CustomerStatusID"), "C.CustomerStatusID,", "") & _
                                    IIf(CustomerDtColumns.Contains("CurrYearSTD"), "C.CurrYearSTD,", "") & _
                                     IIf(CustomerDtColumns.Contains("LastYearSTD"), "C.LastYearSTD,", "") & _
                                      IIf(CustomerDtColumns.Contains("PrimaryCardID"), " cid.ExtCardID as PrimaryCardID,", "") & _
                                       IIf(CustomerDtColumns.Contains("CurrYearSTD"), "C.CurrYearSTD,", "") & _
                                        IIf(CustomerDtColumns.Contains("LastYearSTD"), " C.LastYearSTD,", "") & _
                                         IIf(CustomerDtColumns.Contains("CustomerTypeID"), "C.CustomerTypeID,", "") & _
                                           IIf(CustomerDtColumns.Contains("HHPK") AndAlso Not CustomerDtColumns.Contains("HHID"), "C.HHPK,", "") & _
                                            IIf(CustomerDtColumns.Contains("Address"), "E.Address,", "") & _
                                             IIf(CustomerDtColumns.Contains("City"), "E.City,", "") & _
                                              IIf(CustomerDtColumns.Contains("State"), "E.State,", "") & _
                                               IIf(CustomerDtColumns.Contains("Zip"), "E.Zip,", "") & _
                                                IIf(CustomerDtColumns.Contains("Country"), "E.Country,", "") & _
                                                 IIf(CustomerDtColumns.Contains("Phone"), "E.PhoneAsEntered as Phone,", "") & _
												                         IIf(CustomerDtColumns.Contains("Mobile"), "E.MobilePhoneAsEntered as Mobile,", "") & _
                                                  IIf(CustomerDtColumns.Contains("Email"), "E.Email,", "") & _
                                                   IIf(CustomerDtColumns.Contains("DOB"), "E.DOB,", "") & _
                                                    IIf(CustomerDtColumns.Contains("HHID"), "'' as HHID, C.HHPK,", "") & _
                                                    IIf(CustomerDtColumns.Contains("CustomerID"), "'" & CustomerID & "' as CustomerID, ", "") & _
                                                       IIf(CustomerDtColumns.Contains("InitialCardID"), "C.InitialCardID ", "")
                            MyCommon.QueryStr = MyCommon.QueryStr.Trim(",") & " from Customers as C inner join CardIDs as cid on C.CustomerPK = cid.CustomerPK "
                            If (CustomerDtColumns.Contains("Address") OrElse CustomerDtColumns.Contains("City") OrElse CustomerDtColumns.Contains("State") OrElse CustomerDtColumns.Contains("Zip") OrElse CustomerDtColumns.Contains("Country") OrElse CustomerDtColumns.Contains("Phone") OrElse CustomerDtColumns.Contains("Mobile") OrElse CustomerDtColumns.Contains("Email") OrElse CustomerDtColumns.Contains("DOB")) Then
                                MyCommon.QueryStr = MyCommon.QueryStr & " left join CustomerEXT as E with (NoLock) on C.CustomerPK=E.CustomerPK "
                            End If
                            MyCommon.QueryStr = MyCommon.QueryStr & " where cid.cardtypeid = 0 and C.CustomerPK=" & CustomerPK.ToString & ";"
                        Else
                            MyCommon.QueryStr = "select " & _
                              IIf(CustomerDtColumns.Contains("CustomerPK"), "C.CustomerPK,", "") & _
                               IIf(CustomerDtColumns.Contains("FirstName"), "C.FirstName,", "") & _
							                 IIf(CustomerDtColumns.Contains("MiddleName"), " C.MiddleName,", "") & _
                                IIf(CustomerDtColumns.Contains("LastName"), " C.LastName,", "") & _
                                 IIf(CustomerDtColumns.Contains("Employee"), "C.Employee,", "") & _
                                  IIf(CustomerDtColumns.Contains("CustomerStatusID"), "C.CustomerStatusID,", "") & _
                                   IIf(CustomerDtColumns.Contains("CurrYearSTD"), "C.CurrYearSTD,", "") & _
                                    IIf(CustomerDtColumns.Contains("LastYearSTD"), "C.LastYearSTD,", "") & _
                                     IIf(CustomerDtColumns.Contains("PrimaryCardID"), " cid.ExtCardID as PrimaryCardID,", "") & _
                                       IIf(CustomerDtColumns.Contains("CustomerTypeID"), "C.CustomerTypeID,", "") & _
                                          IIf(CustomerDtColumns.Contains("HHPK") AndAlso Not CustomerDtColumns.Contains("HHID"), "C.HHPK,", "") & _
                                           IIf(CustomerDtColumns.Contains("Address"), "E.Address,", "") & _
                                            IIf(CustomerDtColumns.Contains("City"), "E.City,", "") & _
                                             IIf(CustomerDtColumns.Contains("State"), "E.State,", "") & _
                                              IIf(CustomerDtColumns.Contains("Zip"), "E.Zip,", "") & _
                                               IIf(CustomerDtColumns.Contains("Country"), "E.Country,", "") & _
                                                IIf(CustomerDtColumns.Contains("Phone"), "E.PhoneAsEntered as Phone,", "") & _
												 IIf(CustomerDtColumns.Contains("Mobile"), "E.MobilePhoneAsEntered as Mobile,", "") & _
                                                 IIf(CustomerDtColumns.Contains("Email"), "E.Email,", "") & _
                                                  IIf(CustomerDtColumns.Contains("DOB"), "E.DOB,", "") & _
                                                   IIf(CustomerDtColumns.Contains("HHID"), "'' as HHID, C.HHPK, ", "") & _
                                                   IIf(CustomerDtColumns.Contains("CustomerID"), "'" & CustomerID & "' as CustomerID,", "") & _
                                                    IIf(CustomerDtColumns.Contains("InitialCardID"), "C.InitialCardID ", "")
                            MyCommon.QueryStr = MyCommon.QueryStr.Trim(",") & " from Customers as C inner join CardIDs as cid on C.CustomerPK = cid.CustomerPK "
                            If (CustomerDtColumns.Contains("Address") OrElse CustomerDtColumns.Contains("City") OrElse CustomerDtColumns.Contains("State") OrElse CustomerDtColumns.Contains("Zip") OrElse CustomerDtColumns.Contains("Country") OrElse CustomerDtColumns.Contains("Phone") OrElse CustomerDtColumns.Contains("Mobile") OrElse CustomerDtColumns.Contains("Email") OrElse CustomerDtColumns.Contains("DOB")) Then
                                MyCommon.QueryStr = MyCommon.QueryStr & " left join CustomerEXT as E with (NoLock) on C.CustomerPK=E.CustomerPK "
                            End If
                            MyCommon.QueryStr = MyCommon.QueryStr & " where C.CustomerPK=" & CustomerPK.ToString & ";"
                        End If
                    End If
                    dt = MyCommon.LXS_Select
                    WriteDebug("Finished Get of Customer Details FilteredResponse with row count=" & dt.Rows.Count, DebugState.CurrentTime)

                    If dt.Rows.Count > 0 Then
                        If (CustomerDtColumns.Contains("HHID") AndAlso MyCommon.NZ(dt.Rows(0).Item("HHPK"), 0) > 0) Then
                            WriteDebug("Starting Get TOP ExtCardID from CardIDs [LogixXS]", DebugState.CurrentTime)
                            MyCommon.QueryStr = "select top 1 ExtCardID from CardIDs with (NoLock) where CustomerPK=" & dt.Rows(0).Item("HHPK") '& ";"
                            dt2 = MyCommon.LXS_Select
                            WriteDebug("Finished Get TOP ExtCardID from CardIDs with row count=" & dt2.Rows.Count, DebugState.CurrentTime)
                            If dt2.Rows.Count > 0 Then
                                dt.Rows(0).Item("HHID") = MyCryptLib.SQL_StringDecrypt(dt2.Rows(0).Item("ExtCardID").ToString())
                            End If
                        End If
                        If (CustomerDtColumns.Contains("HHID") AndAlso Not CustomerDtColumns.Contains("HHPK") AndAlso dt.Columns.Contains("HHPK")) Then
                            dt.Columns.Remove("HHPK")
                            dt.AcceptChanges()
                        End If
                        If (CustomerDtColumns.Contains("DOB") AndAlso Not String.IsNullOrEmpty(Convert.ToString(dt.Rows(0).Item("DOB")))) Then
                            Dim MyLookup1 As New Copient.CustomerLookup
                            Dim dob1 As Date = MyLookup1.ParseDateOfBirth(MyCryptLib.SQL_StringDecrypt(MyCommon.NZ(dt.Rows(0).Item("DOB"), "")))
                            If dob1 <> Nothing Then
                              dt.Rows(0).Item("DOB") = dob1.ToString("yyyy-MM-dd")
                            End If
                        End If
                        If Not String.IsNullOrWhiteSpace((dt.Rows(0).Item("Phone").ToString())) Then
                            dt.Rows(0).Item("Phone") = MyCryptLib.SQL_StringDecrypt(dt.Rows(0).Item("Phone"))
                        End If
                        If Not String.IsNullOrWhiteSpace((dt.Rows(0).Item("Mobile").ToString())) Then
                            dt.Rows(0).Item("Mobile") = MyCryptLib.SQL_StringDecrypt(dt.Rows(0).Item("Mobile"))
                        End If
                        If Not String.IsNullOrWhiteSpace((dt.Rows(0).Item("Email").ToString())) Then
                            dt.Rows(0).Item("Email") = MyCryptLib.SQL_StringDecrypt(dt.Rows(0).Item("Email"))
                        End If
                        If Not String.IsNullOrWhiteSpace((dt.Rows(0).Item("PrimaryCardID").ToString())) Then
                            dt.Rows(0).Item("PrimaryCardID") = MyCryptLib.SQL_StringDecrypt(dt.Rows(0).Item("PrimaryCardID"))
                        End If
                        If Not String.IsNullOrWhiteSpace((dt.Rows(0).Item("InitialCardID").ToString())) Then
                            dt.Rows(0).Item("InitialCardID") = MyCryptLib.SQL_StringDecrypt(dt.Rows(0).Item("InitialCardID"))
                        End If
                       
                        If (Not PointsProgramDtColumns Is Nothing AndAlso PointsProgramDtColumns.Length > 0) Then
                            WriteDebug("Enter CustomerLookup.Send_PointsProgramBalances()", DebugState.CurrentTime)
                            dtBalances = Send_PointsProgramBalances_FilteredResponse(CustomerPK, String.Join(",", PointsProgramDtColumns), BalRetCode, BalRetMsg)
                            WriteDebug("Exit CustomerLookup.Send_PointsProgramBalances()", DebugState.CurrentTime)
                        End If
                        If BalRetCode = StatusCodes.SUCCESS Then
                            If (Not CustomerAttributesDtColumns Is Nothing AndAlso CustomerAttributesDtColumns.Length > 0) Then
                                WriteDebug("Starting SELECT AttributeTypeID,AttributeValueID from CustomerAttributes [LogixXS]", DebugState.CurrentTime)
                              If MyCommon.Fetch_SystemOption(110) <> "0" Then
								                MyCommon.QueryStr = "select AttributeTypeID,AttributeValueID from CustomerAttributes with (noLock) where CustomerPK=" + CustomerPK.ToString + " and Deleted=0"
                                dtCustAttrIDs = MyCommon.LXS_Select
                                WriteDebug("Finished SELECT AttributeTypeID,AttributeValueID from CustomerAttributes with row count=" & dtCustAttrIDs.Rows.Count, DebugState.CurrentTime)
                                
								                 If dtCustAttrIDs.Rows.Count > 0 Then
                                    'Details found
                                    dtCustomerAttributes = New DataTable
                                    dtCustomerAttributes.TableName = "CustomerAttributes"
                                    If (CustomerAttributesDtColumns.Contains("CustomerPK")) Then dtCustomerAttributes.Columns.Add("CustomerPK", System.Type.GetType("System.Int64"))
                                    If (CustomerAttributesDtColumns.Contains("ExtCardID")) Then dtCustomerAttributes.Columns.Add("ExtCardID", System.Type.GetType("System.String"))
                                    If (CustomerAttributesDtColumns.Contains("CardTypeID")) Then dtCustomerAttributes.Columns.Add("CardTypeID", System.Type.GetType("System.Int32"))
                                    If (CustomerAttributesDtColumns.Contains("TypeID")) Then dtCustomerAttributes.Columns.Add("TypeID", System.Type.GetType("System.Int32"))
                                    If (CustomerAttributesDtColumns.Contains("TypeDescription")) Then dtCustomerAttributes.Columns.Add("TypeDescription", System.Type.GetType("System.String"))
                                    If (CustomerAttributesDtColumns.Contains("ValueID")) Then dtCustomerAttributes.Columns.Add("ValueID", System.Type.GetType("System.Int32"))
                                    If (CustomerAttributesDtColumns.Contains("ValueDescription")) Then dtCustomerAttributes.Columns.Add("ValueDescription", System.Type.GetType("System.String"))
                                    For i As Integer = 0 To dtCustAttrIDs.Rows.Count - 1
                                        If (CustomerAttributesDtColumns.Contains("ValueDescription") OrElse CustomerAttributesDtColumns.Contains("TypeDescription")) Then
                                            WriteDebug("Starting FOR- SELECT AV.Description as ValueDescription,AT.Description as TypeDescription from AttributeValues [LogixRT]", DebugState.CurrentTime)
                                            MyCommon.QueryStr = "select " & _
                                               IIf(CustomerAttributesDtColumns.Contains("ValueDescription"), "Av.Description as ValueDescription,", "") & _
                                                IIf(CustomerAttributesDtColumns.Contains("TypeDescription"), "At.Description as TypeDescription  ", "")
                                            MyCommon.QueryStr = MyCommon.QueryStr.Trim(",") & " from AttributeValues Av with (noLock) " & _
                                                                    " inner join AttributeTypes At with (noLock)  on Av.AttributeTypeID=At.AttributeTypeID where " & _
                                                                    " Av.AttributeValueID=" + dtCustAttrIDs.Rows(i)("AttributeValueID").ToString + " and " & _
                                                                    " Av.AttributeTypeID=" + dtCustAttrIDs.Rows(i)("AttributeTypeID").ToString + " and Av.Deleted=0 and At.Deleted=0 "
                                            dtCustAttrDesc = MyCommon.LRT_Select
                                            WriteDebug("Finished FOR- SELECT AV.Description as ValueDescription,AT.Description as TypeDescription from AttributeValues with row count=" & dtCustAttrDesc.Rows.Count, DebugState.CurrentTime)
                                        End If
                                        If Not dtCustAttrDesc Is Nothing AndAlso dtCustAttrDesc.Rows.Count > 0 Then
                                            row = dtCustomerAttributes.NewRow()
                                            If (CustomerAttributesDtColumns.Contains("CustomerPK")) Then row.Item("CustomerPK") = CustomerPK.ToString
                                            If (CustomerAttributesDtColumns.Contains("ExtCardID")) Then row.Item("ExtCardID") = MyCryptLib.SQL_StringDecrypt(CustomerID.ToString)
                                            If (CustomerAttributesDtColumns.Contains("CardTypeID")) Then row.Item("CardTypeID") = CustomerTypeID.ToString
                                            If (CustomerAttributesDtColumns.Contains("TypeID")) Then row.Item("TypeID") = dtCustAttrIDs.Rows(i)("AttributeTypeID").ToString
                                            If (CustomerAttributesDtColumns.Contains("TypeDescription")) Then row.Item("TypeDescription") = dtCustAttrDesc.Rows(0)("TypeDescription").ToString
                                            If (CustomerAttributesDtColumns.Contains("ValueID")) Then row.Item("ValueID") = dtCustAttrIDs.Rows(i)("AttributeValueID").ToString
                                            If (CustomerAttributesDtColumns.Contains("ValueDescription")) Then row.Item("ValueDescription") = dtCustAttrDesc.Rows(0)("ValueDescription").ToString
                                            dtCustomerAttributes.Rows.Add(row)
                                            dtCustomerAttributes.AcceptChanges()
                                        Else
                                            row = dtCustomerAttributes.NewRow()
                                            If (CustomerAttributesDtColumns.Contains("CustomerPK")) Then row.Item("CustomerPK") = CustomerPK.ToString
                                            If (CustomerAttributesDtColumns.Contains("ExtCardID")) Then row.Item("ExtCardID") = CustomerID.ToString
                                            If (CustomerAttributesDtColumns.Contains("CardTypeID")) Then row.Item("CardTypeID") = CustomerTypeID.ToString
                                            If (CustomerAttributesDtColumns.Contains("TypeID")) Then row.Item("TypeID") = dtCustAttrIDs.Rows(i)("AttributeTypeID").ToString
                                            If (CustomerAttributesDtColumns.Contains("ValueID")) Then row.Item("ValueID") = dtCustAttrIDs.Rows(i)("AttributeValueID").ToString
                                            dtCustomerAttributes.Rows.Add(row)
                                            dtCustomerAttributes.AcceptChanges()
                                        End If
                                    Next
                                End If
                             
                                If Not dtCustomerAttributes Is Nothing AndAlso dtCustomerAttributes.Rows.Count > 0 Then
                                    dt.TableName = "Customer"
                                    dt.AcceptChanges()
                                    If (Not dtBalances Is Nothing) Then dtBalances.AcceptChanges()
                                    row = dtStatus.NewRow()
                                    row.Item("StatusCode") = StatusCodes.SUCCESS
                                    row.Item("Description") = "Success."
                                    dtStatus.Rows.Add(row)
                                    dtStatus.AcceptChanges()
                                    ResultSet.Tables.Add(dtStatus.Copy())
                                    ResultSet.Tables.Add(dt.Copy())
                                    If (Not dtBalances Is Nothing) Then ResultSet.Tables.Add(dtBalances.Copy())
                                    If (Not dtCustomerAttributes Is Nothing) Then ResultSet.Tables.Add(dtCustomerAttributes.Copy()) ' Ahold -enhancement
                                Else
                                    dt.TableName = "Customer"
                                    dt.AcceptChanges()
                                    row = dtStatus.NewRow()
                                    row.Item("StatusCode") = StatusCodes.NOTFOUND_CUSTSUPPLEMENTATTRIBUTE
                                    If (Not CustomerAttributesDtColumns Is Nothing) Then
                                        MyCommon.QueryStr = "select top 1 CustomerPK from CustomerAttributes with (noLock) where CustomerPK=" + CustomerPK.ToString + " and Deleted=0"
                                        dtCustAttrIDs = MyCommon.LXS_Select
                                        If dtCustAttrIDs.Rows.Count > 0 Then
                                            row.Item("Description") = "Success."
                                        Else
                                            row.Item("Description") = "Not found Customer attributes."
                                        End If
                                    Else
                                        row.Item("Description") = "Success."
                                    End If
                                    dtStatus.Rows.Add(row)
                                    dtStatus.AcceptChanges()
                                    ResultSet.Tables.Add(dtStatus.Copy())
                                    ResultSet.Tables.Add(dt.Copy())
                                    If (Not dtBalances Is Nothing) Then ResultSet.Tables.Add(dtBalances.Copy())
                                End If
                            Else
                                dt.TableName = "Customer"
                                dt.AcceptChanges()
                                row = dtStatus.NewRow()
                                row.Item("StatusCode") = StatusCodes.NOTFOUND_CUSTSUPPLEMENTATTRIBUTE
                                If (Not CustomerAttributesDtColumns Is Nothing) Then
                                    MyCommon.QueryStr = "select top 1 CustomerPK from CustomerAttributes with (noLock) where CustomerPK=" + CustomerPK.ToString + " and Deleted=0"
                                    dtCustAttrIDs = MyCommon.LXS_Select
                                    If dtCustAttrIDs.Rows.Count > 0 Then
                                        row.Item("Description") = "Success."
                                    Else
                                        row.Item("Description") = "Not found Customer attributes."
                                    End If
                                Else
                                    row.Item("Description") = "Success."
                                End If
                                dtStatus.Rows.Add(row)
                                dtStatus.AcceptChanges()
                                ResultSet.Tables.Add(dtStatus.Copy())
                                ResultSet.Tables.Add(dt.Copy())
                                If (Not dtBalances Is Nothing) Then ResultSet.Tables.Add(dtBalances.Copy())
                            End If
                        End If
                        ' Fix for RT6262 End
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            WriteDebug("CustomerDetails FilteredResponse Exception: " & ex.Message, DebugState.CurrentTime)
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
        WriteDebug("Disconnected LogixRT and LogixXS CustomerDetails FilteredResponse", DebugState.CurrentTime)
    
        WriteDebug("CustomerDetails FilteredResponse  - " & System.Net.Dns.GetHostName(), DebugState.EndTime)
        WriteLogLines()

        Return ResultSet
    End Function

    Private Function _CustomerDetails(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As Integer, Optional ByVal CardTypeID As Integer = -1) As System.Data.DataSet
        Dim dt, dt2, dtCustAttrIDs, dtCustAttrDesc As System.Data.DataTable
        Dim dtStatus, dtBalances, dtCustomerAttributes As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim ResultSet As New System.Data.DataSet("CustomerDetails")
        Dim CustomerPK As Long = 0
        Dim BalRetMsg As String = ""
        Dim BalRetCode As StatusCodes = StatusCodes.SUCCESS
        Dim AltColumn As String = "Not Set"
        Dim CMInstalled As Boolean = False

        WriteDebug("CustomerDetails - CardNumber: " & Copient.MaskHelper.MaskCard(CustomerID, CardTypeID) & ". CardTypeID: " & CardTypeID, DebugState.BeginTime)

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        If GUID = "" Then
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_GUID
            row.Item("Description") = "Failure: Invalid GUID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _CustomerDetails = ResultSet
            Exit Function
        End If
        If CustomerID = "" Then
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERID
            row.Item("Description") = "Failure: Invalid Customer ID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _CustomerDetails = ResultSet
            Exit Function
        End If
        If CustomerTypeID = -1 Then
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERTYPEID
            row.Item("Description") = "Failure: Invalid Customer Type"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _CustomerDetails = ResultSet
            Exit Function
        End If
        If GUID.Contains(Chr(34)) = True Or GUID.Contains("'") = True Then
            row = dtStatus.NewRow()
            row.Item("StatusCode") = 1
            row.Item("Description") = "Failure: Invalid GUID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _CustomerDetails = ResultSet
            Exit Function
        End If
        If CustomerID.Contains(Chr(34)) = True Or CustomerID.Contains("'") = True Then
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERID
            row.Item("Description") = "Failure: Invalid Customer ID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _CustomerDetails = ResultSet
            Exit Function
        End If

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
            WriteDebug("Connected LogixRT and LogixXS CustomerDetails", DebugState.CurrentTime)

            If Not IsValidGUID(GUID) Then
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
            ElseIf (CustomerID.Length < 1) Then
                'Bad customer ID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERID
                row.Item("Description") = "Failure: Invalid customer ID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
            Else
                'First test the CustomerTypeID for alternate column 
                CMInstalled = MyCommon.IsEngineInstalled(0)
                If (CMInstalled) And (CustomerTypeID = 4) Then
                       
                    'find the alternate search column via systemobjects
                    AltColumn = MyCommon.Fetch_SystemOption(60)
                    Dim split As String() = AltColumn.Split(New [Char]() {"."c})

                    'Find the Customer PK
                    MyCommon.QueryStr = "select " & split(0) & ".CustomerPK from " & split(0) & " with (NoLock) " & _
                                        "where " & split(1) & "='" & MyCryptLib.SQL_StringEncrypt(CustomerID) & "';"

                Else
                    'Pad the customer ID
                    If CardTypeID <> -1 Then
                        CustomerID = MyCommon.Pad_ExtCardID(CustomerID, CardTypeID)
                    Else
                        CustomerID = MyCommon.Pad_ExtCardID(CustomerID, CardTypes.CUSTOMER)
                    End If
                    'Find the Customer PK
                    WriteDebug("Starting Find of CustomerPK [LogixXS]", DebugState.CurrentTime)
                    MyCommon.QueryStr = "select CID.CustomerPK from CardIDs as CID with (NoLock) " & _
                                        "INNER JOIN CardTypes AS CT with (NoLock) ON CID.CardTypeID = CT.CardTypeID " & _
                                        "WHERE CT.CustTypeID=" & CustomerTypeID & " AND CID.ExtCardID='" & MyCryptLib.SQL_StringEncrypt(CustomerID, True) & "';"
                End If

                dt = MyCommon.LXS_Select
                WriteDebug("Finished Find of CustomerPK with row count=" & dt.Rows.Count, DebugState.CurrentTime)
                If dt.Rows.Count = 0 Then
                    'Customer not found
                    If CustomerTypeID = 0 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_CUSTOMER
                        row.Item("Description") = "Failure: Customer " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    ElseIf CustomerTypeID = 1 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_HOUSEHOLD
                        row.Item("Description") = "Failure: Household " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    ElseIf CustomerTypeID = 2 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_CAM
                        row.Item("Description") = "Failure: CAM " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    Else
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERTYPEID
                        row.Item("Description") = "Failure: Invalid customer type."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    End If
                Else
                    CustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)
                    'Get the customer's details
                    WriteDebug("Starting Get of Customer Details [LogixXS]", DebugState.CurrentTime)
                    If (MyCommon.Extract_Val(MyCommon.Fetch_SystemOption(6)) >= 2) And (CustomerTypeID = 0) Then
                        'Savings To Date is turned on for household level - we need to return the household's STD along with the customer's STD
                        If (CMInstalled) And (CustomerTypeID = 4) Then
                            MyCommon.QueryStr = "select C.CustomerPK, C.FirstName, C.MiddleName, C.LastName, C.Employee, C.CustomerStatusID, C.CurrYearSTD, C.LastYearSTD, cid.ExtCardIDOriginal as PrimaryCardID, " & _
                                     " isnull(HHCust.CurrYearSTD, 0) as HHCurrYearSTD, isnull(HHCust.LastYearSTD, 0) as HHLastYearSTD, " & _
                "C.CustomerTypeID, C.InitialCardIDOriginal as InitialCardID, C.HHPK, E.Address, E.City, E.State, E.Zip, E.Country, E.PhoneAsEntered as Phone, E.MobilePhoneAsEntered as Mobile,  E.Email, E.DOB, '' as HHID, '" & CustomerID & "' as CustomerID " & _
                                     "from Customers as C  " & _
                                     "inner join CardIDs as cid on C.CustomerPK = cid.CustomerPK " & _
                                     "left join CustomerEXT as E with (NoLock) on C.CustomerPK=E.CustomerPK " & _
                                     "Left join Customers as HHCust with (NoLock) on C.HHPK=HHCust.CustomerPK " & _
                                                 "where cid.cardtypeid = 0 and C.CustomerPK=" & CustomerPK.ToString & ";"
                        Else
                            MyCommon.QueryStr = "select C.CustomerPK, C.FirstName, C.MiddleName, C.LastName, C.Employee, C.CustomerStatusID, C.CurrYearSTD, C.LastYearSTD, C.InitialCardIDOriginal as PrimaryCardID, " & _
                                     " isnull(HHCust.CurrYearSTD, 0) as HHCurrYearSTD, isnull(HHCust.LastYearSTD, 0) as HHLastYearSTD, " & _
                                     "C.CustomerTypeID, C.HHPK, E.Address, E.City, E.State, E.Zip, E.Country, E.PhoneAsEntered as Phone, E.Email, E.MobilePhoneAsEntered as Mobile, E.DOB, '' as HHID, '" & CustomerID & "' as CustomerID,C.InitialCardIDOriginal as InitialCardID " & _
                                     "from Customers as C with (NoLock) " & _
                                     "left join CustomerEXT as E with (NoLock) on C.CustomerPK=E.CustomerPK " & _
                                     "Left join Customers as HHCust with (NoLock) on C.HHPK=HHCust.CustomerPK " & _
                                                 "where C.CustomerPK=" & CustomerPK.ToString & ";"
                        End If
                        
                    Else
                        If (CMInstalled) And (CustomerTypeID = 4) Then
                            MyCommon.QueryStr = "select C.CustomerPK, C.FirstName, C.MiddleName, C.LastName, C.Employee, cid.ExtCardIDOriginal as PrimaryCardID, C.CustomerStatusID, C.CurrYearSTD, C.LastYearSTD, " & _
                            "C.CustomerTypeID, C.InitialCardIDOriginal as InitialCardID, C.HHPK, E.Address, E.City, E.State, E.Zip, E.Country, E.PhoneAsEntered as Phone, E.MobilePhoneAsEntered as Mobile, E.Email, E.DOB, '' as HHID, " & _
                                     "'" & CustomerID & "' as CustomerID " & _
                                     "from Customers as C  with (NoLock) " & _
                                     "inner join CardIDs as cid on C.CustomerPK = cid.CustomerPK " & _
                                     "left join CustomerEXT as E with (NoLock) on C.CustomerPK=E.CustomerPK " & _
                                                 "where cid.cardtypeid = 0 and C.CustomerPK=" & CustomerPK.ToString & ";"
                        Else
                            MyCommon.QueryStr = "select C.CustomerPK, C.FirstName, C.MiddleName, C.LastName, C.Employee, C.InitialCardIDOriginal as PrimaryCardID, C.CustomerStatusID, C.CurrYearSTD, C.LastYearSTD, " & _
                                     "C.CustomerTypeID, C.HHPK, E.Address, E.City, E.State, E.Zip, E.Country, E.PhoneAsEntered as Phone, E.MobilePhoneAsEntered as Mobile, E.Email, E.DOB, '' as HHID, " & _
                                     "'" & CustomerID & "' as CustomerID ,C.InitialCardIDOriginal as InitialCardID " & _
                                     "from Customers as C with (NoLock) " & _
                                     "left join CustomerEXT as E with (NoLock) on C.CustomerPK=E.CustomerPK " & _
                                                 "where C.CustomerPK=" & CustomerPK.ToString & ";"
                        End If
                    End If
                    dt = MyCommon.LXS_Select
                    WriteDebug("Finished Get of Customer Details with row count=" & dt.Rows.Count, DebugState.CurrentTime)

                    If dt.Rows.Count > 0 Then
                        If MyCommon.NZ(dt.Rows(0).Item("HHPK"), 0) > 0 Then
                            WriteDebug("Starting Get TOP ExtCardID from CardIDs [LogixXS]", DebugState.CurrentTime)
                            MyCommon.QueryStr = "select top 1 ExtCardID from CardIDs with (NoLock) where CustomerPK=" & dt.Rows(0).Item("HHPK") '& ";"
                            dt2 = MyCommon.LXS_Select
                            WriteDebug("Finished Get TOP ExtCardID from CardIDs with row count=" & dt2.Rows.Count, DebugState.CurrentTime)
                            If dt2.Rows.Count > 0 Then
                                dt.Rows(0).Item("HHID") = MyCryptLib.SQL_StringDecrypt(dt2.Rows(0).Item("ExtCardID").ToString())
                            End If
                        End If
                        If Not String.IsNullOrEmpty(Convert.ToString(dt.Rows(0).Item("DOB"))) Then
                            Dim MyLookup1 As New Copient.CustomerLookup
                            Dim dob1 As Date = MyLookup1.ParseDateOfBirth(MyCryptLib.SQL_StringDecrypt(MyCommon.NZ(dt.Rows(0).Item("DOB"), "")))
                            If dob1 <> Nothing Then
                              dt.Rows(0).Item("DOB") = dob1.ToString("yyyy-MM-dd")
                            End If
                        End If
                        If Not String.IsNullOrWhiteSpace((dt.Rows(0).Item("Phone").ToString())) Then
                            dt.Rows(0).Item("Phone") = MyCryptLib.SQL_StringDecrypt(dt.Rows(0).Item("Phone"))
                        End If
                        If Not String.IsNullOrWhiteSpace((dt.Rows(0).Item("Mobile").ToString())) Then
                            dt.Rows(0).Item("Mobile") = MyCryptLib.SQL_StringDecrypt(dt.Rows(0).Item("Mobile"))
                        End If
                        If Not String.IsNullOrWhiteSpace((dt.Rows(0).Item("Email").ToString())) Then
                            dt.Rows(0).Item("Email") = MyCryptLib.SQL_StringDecrypt(dt.Rows(0).Item("Email"))
                        End If
                        If Not String.IsNullOrWhiteSpace((dt.Rows(0).Item("PrimaryCardID").ToString())) Then
                            dt.Rows(0).Item("PrimaryCardID") = MyCryptLib.SQL_StringDecrypt(dt.Rows(0).Item("PrimaryCardID"))
                        End If
                        If Not String.IsNullOrWhiteSpace((dt.Rows(0).Item("InitialCardID").ToString())) Then
                            dt.Rows(0).Item("InitialCardID") = MyCryptLib.SQL_StringDecrypt(dt.Rows(0).Item("InitialCardID"))
                        End If
                        WriteDebug("Enter CustomerLookup.Send_PointsProgramBalances()", DebugState.CurrentTime)
                        dtBalances = Send_PointsProgramBalances(CustomerPK, BalRetCode, BalRetMsg)
                        WriteDebug("Exit CustomerLookup.Send_PointsProgramBalances()", DebugState.CurrentTime)
                        If BalRetCode = StatusCodes.SUCCESS Then
                            WriteDebug("Starting SELECT AttributeTypeID,AttributeValueID from CustomerAttributes [LogixXS]", DebugState.CurrentTime)
                            If MyCommon.Fetch_SystemOption(110) = "0" Then
                              'send customer details and program balances
                              dt.TableName = "Customer"
                              dt.AcceptChanges()
                              dtBalances.AcceptChanges()
                              row = dtStatus.NewRow()
                              row.Item("StatusCode") = StatusCodes.SUCCESS
                              row.Item("Description") = "Success."
                              dtStatus.Rows.Add(row)
                              dtStatus.AcceptChanges()
                              ResultSet.Tables.Add(dtStatus.Copy())
                              ResultSet.Tables.Add(dt.Copy())
                              ResultSet.Tables.Add(dtBalances.Copy())
                            End If
                        ElseIf BalRetCode = StatusCodes.FAILED_BALANCE_LOOKUP Then
                            ' send back the customer details without the program balances
                            dt.TableName = "Customer"
                            dt.AcceptChanges()
                            row = dtStatus.NewRow()
                            row.Item("StatusCode") = StatusCodes.FAILED_BALANCE_LOOKUP
                            row.Item("Description") = BalRetMsg
                            dtStatus.Rows.Add(row)
                            dtStatus.AcceptChanges()
                            ResultSet.Tables.Add(dtStatus.Copy())
                            ResultSet.Tables.Add(dt.Copy())
                        Else
                            ' send back an error message
                            row = dtStatus.NewRow()
                            row.Item("StatusCode") = BalRetCode
                            row.Item("Description") = BalRetMsg
                            dtStatus.Rows.Add(row)
                            dtStatus.AcceptChanges()
                            ResultSet.Tables.Add(dtStatus.Copy())
                        End If

                        ' Fix for RT6262 Start
                        If MyCommon.Fetch_SystemOption(110) <> "0" Then
                            MyCommon.QueryStr = "select AttributeTypeID,AttributeValueID from CustomerAttributes with (noLock) where CustomerPK=" + CustomerPK.ToString + " and Deleted=0"
                            dtCustAttrIDs = MyCommon.LXS_Select
                            WriteDebug("Finished SELECT AttributeTypeID,AttributeValueID from CustomerAttributes with row count=" & dtCustAttrIDs.Rows.Count, DebugState.CurrentTime)
                            If dtCustAttrIDs.Rows.Count > 0 Then
                                'Details found
                                dtCustomerAttributes = New DataTable
                                dtCustomerAttributes.TableName = "CustomerAttributes"
                                dtCustomerAttributes.Columns.Add("CustomerPK", System.Type.GetType("System.Int64"))
                                dtCustomerAttributes.Columns.Add("ExtCardID", System.Type.GetType("System.String"))
                                dtCustomerAttributes.Columns.Add("CardTypeID", System.Type.GetType("System.Int32"))
                                dtCustomerAttributes.Columns.Add("TypeID", System.Type.GetType("System.Int32"))
                                dtCustomerAttributes.Columns.Add("TypeDescription", System.Type.GetType("System.String"))
                                dtCustomerAttributes.Columns.Add("ValueID", System.Type.GetType("System.Int32"))
                                dtCustomerAttributes.Columns.Add("ValueDescription", System.Type.GetType("System.String"))
                                dtCustomerAttributes.Columns.Add("InitialCardID", System.Type.GetType("System.String"))
                                For i As Integer = 0 To dtCustAttrIDs.Rows.Count - 1
                                    WriteDebug("Starting FOR- SELECT AV.Description as ValueDescription,AT.Description as TypeDescription from AttributeValues [LogixRT]", DebugState.CurrentTime)
                                    MyCommon.QueryStr = "select Av.Description as ValueDescription ,At.Description as TypeDescription from AttributeValues Av with (noLock) " & _
                                                            " inner join AttributeTypes At with (noLock)  on Av.AttributeTypeID=At.AttributeTypeID where " & _
                                                            " Av.AttributeValueID=" + dtCustAttrIDs.Rows(i)("AttributeValueID").ToString + " and " & _
                                                            " Av.AttributeTypeID=" + dtCustAttrIDs.Rows(i)("AttributeTypeID").ToString + " and Av.Deleted=0 and At.Deleted=0 "
                                    dtCustAttrDesc = MyCommon.LRT_Select
                                    WriteDebug("Finished FOR- SELECT AV.Description as ValueDescription,AT.Description as TypeDescription from AttributeValues with row count=" & dtCustAttrDesc.Rows.Count, DebugState.CurrentTime)
                                    If dtCustAttrDesc.Rows.Count > 0 Then
                                        row = dtCustomerAttributes.NewRow()
                                        row.Item("CustomerPK") = CustomerPK.ToString
                                        row.Item("ExtCardID") = MyCryptLib.SQL_StringDecrypt(CustomerID.ToString)
                                        row.Item("CardTypeID") = CustomerTypeID.ToString
                                        row.Item("TypeID") = dtCustAttrIDs.Rows(i)("AttributeTypeID").ToString
                                        row.Item("TypeDescription") = dtCustAttrDesc.Rows(0)("TypeDescription").ToString
                                        row.Item("ValueID") = dtCustAttrIDs.Rows(i)("AttributeValueID").ToString
                                        row.Item("ValueDescription") = dtCustAttrDesc.Rows(0)("ValueDescription").ToString
                                        dtCustomerAttributes.Rows.Add(row)
                                        dtCustomerAttributes.AcceptChanges()
                                    End If
                                Next
                                If dtCustomerAttributes.Rows.Count > 0 Then
                                    dt.TableName = "Customer"
                                    dt.AcceptChanges()
                                    dtBalances.AcceptChanges()
                                    row = dtStatus.NewRow()
                                    row.Item("StatusCode") = StatusCodes.SUCCESS
                                    row.Item("Description") = "Success."
                                    dtStatus.Rows.Add(row)
                                    dtStatus.AcceptChanges()
                                    ResultSet.Tables.Add(dtStatus.Copy())
                                    ResultSet.Tables.Add(dt.Copy())
                                    ResultSet.Tables.Add(dtBalances.Copy())
                                    ResultSet.Tables.Add(dtCustomerAttributes.Copy()) '  -enhancement
                                Else
                                    dt.TableName = "Customer"
                                    dt.AcceptChanges()
                                    row = dtStatus.NewRow()
                                    row.Item("StatusCode") = StatusCodes.NOTFOUND_CUSTSUPPLEMENTATTRIBUTE
                                    row.Item("Description") = "Not found Customer attributes."
                                    dtStatus.Rows.Add(row)
                                    dtStatus.AcceptChanges()
                                    ResultSet.Tables.Add(dtStatus.Copy())
                                    ResultSet.Tables.Add(dt.Copy())
                                    ResultSet.Tables.Add(dtBalances.Copy())
                                End If
                            Else
                                dt.TableName = "Customer"
                                dt.AcceptChanges()
                                row = dtStatus.NewRow()
                                row.Item("StatusCode") = StatusCodes.NOTFOUND_CUSTSUPPLEMENTATTRIBUTE
                                row.Item("Description") = "Not found Customer attributes."
                                dtStatus.Rows.Add(row)
                                dtStatus.AcceptChanges()
                                ResultSet.Tables.Add(dtStatus.Copy())
                                ResultSet.Tables.Add(dt.Copy())
                                ResultSet.Tables.Add(dtBalances.Copy())
                            End If
                        End If
                        ' Fix for RT6262 End
                    End If
                End If
            End If
        Catch ex As Exception
            WriteDebug("CustomerDetails Exception: " & ex.Message, DebugState.CurrentTime)
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
        WriteDebug("Disconnected LogixRT and LogixXS CustomerDetails", DebugState.CurrentTime)
    
        WriteDebug("CustomerDetails - " & System.Net.Dns.GetHostName(), DebugState.EndTime)
        WriteLogLines()

        Return ResultSet
    End Function
  
    '" AMS Integration- Product Enhancements FIS I" -<2.3.17>
    <WebMethod()> _
    Public Function UpdateCustSupplementAttributes(ByVal GUID As String, ByVal CardID As String, ByVal CardType As String, ByVal AttributeName As String, ByVal Value As String) As DataSet
        Dim iCardType As Integer = -1
        Try
            InitApp()
            iCardType = CInt(CardType)
        Catch ex As Exception
            iCardType = -1
        End Try
        Return _UpdateCustSupplementDetails(GUID, CardID, iCardType, AttributeName, Value)
    End Function
    '" AMS Integration- Product Enhancements FIS I" -<2.3.17>
    <WebMethod()> _
    Public Function UpdateCustSupplementAttributes_ByCardID(ByVal GUID As String, ByVal CardID As String, ByVal CustomerType As String, ByVal AttributeName As String, ByVal Value As String, ByVal CardTypeID As Integer) As DataSet
        Dim iCardType As Integer = -1
        Try
            InitApp()
            iCardType = CInt(CustomerType)
        Catch ex As Exception
            iCardType = -1
        End Try
        Return _UpdateCustSupplementDetails(GUID, CardID, CustomerType, AttributeName, Value, CardTypeID)
    End Function
  
    Private Function _UpdateCustSupplementDetails(ByVal GUID As String, ByVal CardID As String, ByVal CardType As Integer, _
                                                  ByVal AttributeName As String, ByVal Value As String, Optional ByVal CardTypeID As Integer = -1) As DataSet
        Dim ResultSet As New System.Data.DataSet("CustomerSupplementDetails")
        Dim dt As System.Data.DataTable
        Dim dtStatus As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim CustomerPK As Long = 0
        Dim ifield As Integer = 0

        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))


        ' Check the exception scenarios
        If GUID = "" Then
            'GUID Value is empty
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_GUID
            row.Item("Description") = "Failure: Invalid GUID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _UpdateCustSupplementDetails = ResultSet
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
            _UpdateCustSupplementDetails = ResultSet
            Exit Function
        End If
        If CardType = -1 Then
            'CardType Value is empty
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERTYPEID
            row.Item("Description") = "Failure: Invalid Customer Type"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _UpdateCustSupplementDetails = ResultSet
            Exit Function
        End If
        If AttributeName = "" Then
            'Attribute Name is empty
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_CUSTSUPPLEMENTATTRIBUTE
            row.Item("Description") = "Failure: Invalid Supplement Attribute"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _UpdateCustSupplementDetails = ResultSet
            Exit Function
        End If
        If Value.Trim = "" Then
            'Value is empty
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_SUPPATTRIBUTEVALUE
            row.Item("Description") = "Failure: Invalid Supplement Attribute value"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _UpdateCustSupplementDetails = ResultSet
            Exit Function
        End If

        If GUID.Contains(Chr(34)) = True Or GUID.Contains("'") = True Then
            'GUID Value is empty
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_GUID
            row.Item("Description") = "Failure: Invalid GUID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _UpdateCustSupplementDetails = ResultSet
            Exit Function
        End If
        If CardID.Contains(Chr(34)) = True Or CardID.Contains("'") = True Then
            'CardID Value is empty
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERID
            row.Item("Description") = "Failure: Invalid Customer ID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _UpdateCustSupplementDetails = ResultSet
            Exit Function
        End If
        If AttributeName.Contains(Chr(34)) = True Or AttributeName.Contains("'") = True Then
            'CardID Value is empty
            'Attribute Name is empty
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_CUSTSUPPLEMENTATTRIBUTE
            row.Item("Description") = "Failure: Invalid Supplement Attribute"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _UpdateCustSupplementDetails = ResultSet
            Exit Function
        End If
        If Value.Contains(Chr(34)) = True Or Value.Contains("'") = True Then
            'Value is empty
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_SUPPATTRIBUTEVALUE
            row.Item("Description") = "Failure: Invalid Supplement Attribute value"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _UpdateCustSupplementDetails = ResultSet
            Exit Function
        End If


        Try
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
            If Not IsValidGUID(GUID) Then
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
                row.Item("Description") = "Failure: Invalid Customer ID"
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
            Else
                ' Check whether the Supplement Attribute Name is valid or not
                MyCommon.QueryStr = "select FieldID from CustomerSupplementalFields with (NoLock) where Name = '" & AttributeName.Trim() & "' and Deleted = 0"
                dt = MyCommon.LXS_Select
                If dt.Rows.Count > 0 Then
                    ' File ID for the attribute
                    ifield = CInt(dt.Rows(0)(0))
                    'Pad the customer ID
                    If CardTypeID <> -1 Then
                        CardID = MyCommon.Pad_ExtCardID(CardID, CardTypeID)
                    Else
                        CardID = MyCommon.Pad_ExtCardID(CardID, CardTypes.CUSTOMER)
                    End If
                    MyCommon.QueryStr = "select CID.CustomerPK from CardIDs as CID with (NoLock) " & _
                                        "INNER JOIN CardTypes AS CT with (NoLock) ON CID.CardTypeID = CT.CardTypeID " & _
                                        "WHERE CT.CustTypeID=" & CardType & " AND CID.ExtCardID='" & MyCryptLib.SQL_StringEncrypt(CardID, True) & "';"

                    dt = MyCommon.LXS_Select
                    If dt.Rows.Count > 0 Then
                        'Get the CustomerPK
                        CustomerPK = dt.Rows(0)(0)
                        'Check whether the suppliment attribute need to add or modify for the customer
                        MyCommon.QueryStr = "select COUNT(0) from CustomerSupplemental with (NoLock) " & _
                                            "where CustomerPK = " & CustomerPK.ToString & " and FieldID = " & ifield.ToString & " and Deleted=0"
                        dt = MyCommon.LXS_Select
                        If dt.Rows(0)(0) > 0 Then
                            'Update the value to the suppliment attribute
                            MyCommon.QueryStr = "Update CustomerSupplemental with (RowLock) SET Value= '" & Value.ToString & "' WHERE " & _
                                                " CustomerPK= " & CustomerPK.ToString & " AND FieldID = " & ifield.ToString & " and Deleted=0;"
                            MyCommon.LXS_Execute()
                            row = dtStatus.NewRow()
                            row.Item("StatusCode") = StatusCodes.SUCCESS
                            row.Item("Description") = "Success: Updated Customer Supplement Attributes."
                            dtStatus.Rows.Add(row)
                            dtStatus.AcceptChanges()
                            ResultSet.Tables.Add(dtStatus.Copy())
                        Else
                            'Adding the value to the suppliment attribute
                            MyCommon.QueryStr = "INSERT INTO CustomerSupplemental(CustomerPK, FieldID,Value,Deleted) VALUES" & _
                                                " ( " & CustomerPK.ToString & " ," & ifield.ToString & ",'" & Value.ToString & "',0) ;"
                            MyCommon.LXS_Execute()
                            row = dtStatus.NewRow()
                            row.Item("StatusCode") = StatusCodes.SUCCESS
                            row.Item("Description") = "Success: Created Customer Supplement Attributes."
                            dtStatus.Rows.Add(row)
                            dtStatus.AcceptChanges()
                            ResultSet.Tables.Add(dtStatus.Copy())
                        End If
                    Else
                        'Customer PK Not found
                        If CardType = 0 Then
                            row = dtStatus.NewRow()
                            row.Item("StatusCode") = StatusCodes.NOTFOUND_CUSTOMER
                            row.Item("Description") = "Failure: Customer " & CardID & " not found."
                            dtStatus.Rows.Add(row)
                            dtStatus.AcceptChanges()
                            ResultSet.Tables.Add(dtStatus.Copy())
                        ElseIf CardType = 1 Then
                            row = dtStatus.NewRow()
                            row.Item("StatusCode") = StatusCodes.NOTFOUND_HOUSEHOLD
                            row.Item("Description") = "Failure: Household " & CardID & " not found."
                            dtStatus.Rows.Add(row)
                            dtStatus.AcceptChanges()
                            ResultSet.Tables.Add(dtStatus.Copy())
                        ElseIf CardType = 2 Then
                            row = dtStatus.NewRow()
                            row.Item("StatusCode") = StatusCodes.NOTFOUND_CAM
                            row.Item("Description") = "Failure: CAM " & CardID & " not found."
                            dtStatus.Rows.Add(row)
                            dtStatus.AcceptChanges()
                            ResultSet.Tables.Add(dtStatus.Copy())
                        Else
                            row = dtStatus.NewRow()
                            row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERTYPEID
                            row.Item("Description") = "Failure: Invalid customer type."
                            dtStatus.Rows.Add(row)
                            dtStatus.AcceptChanges()
                            ResultSet.Tables.Add(dtStatus.Copy())
                        End If
                    End If
                Else
                    'Invalid Attribute.
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = StatusCodes.INVALID_CUSTSUPPLEMENTATTRIBUTE
                    row.Item("Description") = "Failure: Invalid Supplement Attribute"
                    dtStatus.Rows.Add(row)
                    dtStatus.AcceptChanges()
                    ResultSet.Tables.Add(dtStatus.Copy())
                End If
            End If
        Catch ex As Exception
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
            row.Item("Description") = "Failure: Application " & ex.ToString
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
        Finally
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try
        Return ResultSet
    End Function

    '" AMS Integration- Product Enhancements FIS I" -<2.3.17>
    <WebMethod()> _
    Public Function GetCustSupplementAttributes(ByVal GUID As String, ByVal CardID As String, ByVal CardType As String) As DataSet
        Dim iCardType As Integer = -1
        Try
            InitApp()
            iCardType = CInt(CardType)
        Catch ex As Exception
            iCardType = -1
        End Try
        Return _GetCustSupplementDetails(GUID, CardID, iCardType)
    End Function
    '" AMS Integration- Product Enhancements FIS I" -<2.3.17>
    <WebMethod()> _
    Public Function GetCustSupplementAttributes_ByCardID(ByVal GUID As String, ByVal CardID As String, ByVal CustomerType As String, ByVal CardTypeID As String) As DataSet
        Dim iCardTypeID As Integer = -1
        Try
            InitApp()
            iCardTypeID = CInt(CustomerType)
        Catch ex As Exception
            iCardTypeID = -1
        End Try
        Return _GetCustSupplementDetails(GUID, CardID, CustomerType, CardTypeID)
    End Function

    '" AMS Integration- Product Enhancements FIS I" -<2.3.17>
    Private Function _GetCustSupplementDetails(ByVal GUID As String, ByVal CardID As String, ByVal CardType As Integer, Optional ByVal CardTypeID As Integer = -1) As DataSet
        Dim ResultSet As New System.Data.DataSet("CustomerSupplementDetails")
        Dim dt As DataTable
        Dim dtStatus, dtCustSuppTable As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim CustomerPK As Long = 0

        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        ' Check the exception scenarios
        If GUID = "" Then
            'GUID Value is empty
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_GUID
            row.Item("Description") = "Failure: Invalid GUID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _GetCustSupplementDetails = ResultSet
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
            _GetCustSupplementDetails = ResultSet
            Exit Function
        End If
        If CardType = -1 Then
            'CardType Value is empty
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERTYPEID
            row.Item("Description") = "Failure: Invalid Customer Type"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _GetCustSupplementDetails = ResultSet
            Exit Function
        End If

        If GUID.Contains("'") = True Or GUID.Contains(Chr(34)) = True Then
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_GUID
            row.Item("Description") = "Failure: Invalid GUID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _GetCustSupplementDetails = ResultSet
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
            _GetCustSupplementDetails = ResultSet
            Exit Function
        End If
        Try
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
            If Not IsValidGUID(GUID) Then
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                _GetCustSupplementDetails = ResultSet
            ElseIf (CardID.Length < 1) Then
                'Bad customer ID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERID
                row.Item("Description") = "Failure: Invalid Customer ID"
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                _GetCustSupplementDetails = ResultSet
            Else
                'Pad the customer ID
                If CardTypeID <> -1 Then
                    CardID = MyCommon.Pad_ExtCardID(CardID, CardTypeID)
                Else
                    CardID = MyCommon.Pad_ExtCardID(CardID, CardTypes.CUSTOMER)
                End If
                MyCommon.QueryStr = "select CID.CustomerPK from CardIDs as CID with (NoLock) " & _
                                    "INNER JOIN CardTypes AS CT with (NoLock) ON CID.CardTypeID = CT.CardTypeID " & _
                                    "WHERE CT.CustTypeID=" & CardType & " AND CID.ExtCardID='" & MyCryptLib.SQL_StringEncrypt(CardID, True) & "';"

                dt = MyCommon.LXS_Select
                If dt.Rows.Count > 0 Then
                    'Get the CustomerPK
                    CustomerPK = dt.Rows(0)(0)

                    MyCommon.QueryStr = "select '" + CardID + "' as CardID,'" + CardType.ToString + "' as CardType,cs.FieldID,csf.Name,cs.Value " & _
                                        " from CustomerSupplemental cs with (NoLock) Inner Join CustomerSupplementalFields csf with (NoLock) " & _
                                        " on cs.FieldID=csf.FieldID where cs.CustomerPK=" + CustomerPK.ToString + " and cs.Deleted=0 and csf.Deleted=0 ;"

                    dtCustSuppTable = MyCommon.LXS_Select
                    'Get Customer Supplement Attributes from LogixXS DB.
                    If dtCustSuppTable.Rows.Count = 0 Then
                        ' No records found
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_RECORDS
                        row.Item("Description") = "Failure: No records found"
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                        _GetCustSupplementDetails = ResultSet
                    Else
                        ' CSFields records found
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.SUCCESS
                        row.Item("Description") = "Success:"
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        dtCustSuppTable.TableName = "CustSupplementAttributes"
                        dtCustSuppTable.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                        ResultSet.Tables.Add(dtCustSuppTable.Copy())
                        _GetCustSupplementDetails = ResultSet
                    End If
                Else
                    'Customer PK Not found
                    If CardType = 0 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_CUSTOMER
                        row.Item("Description") = "Failure: Customer " & CardID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                        _GetCustSupplementDetails = ResultSet
                    ElseIf CardType = 1 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_HOUSEHOLD
                        row.Item("Description") = "Failure: Household " & CardID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                        _GetCustSupplementDetails = ResultSet
                    ElseIf CardType = 2 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_CAM
                        row.Item("Description") = "Failure: CAM " & CardID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                        _GetCustSupplementDetails = ResultSet
                    Else
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERTYPEID
                        row.Item("Description") = "Failure: Invalid customer type."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                        _GetCustSupplementDetails = ResultSet
                    End If
                End If
            End If
        Catch ex As Exception
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
            row.Item("Description") = "Failure: Application " & ex.ToString
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _GetCustSupplementDetails = ResultSet
        Finally
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try
    End Function

    '"AMS_FIS_hbc_J" - <2.3.14>
    <WebMethod()> _
    Public Function GetOffersByProduct(ByVal GUID As String, ByVal ProductId As String, ByVal ProductType As String) As DataSet
        Dim iProductType As Integer = -1
        Try
            InitApp()
            iProductType = CInt(ProductType)
        Catch ex As Exception
            iProductType = -1
        End Try
        Return _GetOffersByProduct(GUID, ProductId, iProductType)
    End Function
	
    '"AMS_FIS_hbc_J" - <2.3.14>
    Public Function _GetOffersByProduct(ByVal GUID As String, ByVal ProductId As String, ByVal ProductType As Integer) As DataSet
        Dim ResultSet As New System.Data.DataSet("OffersByProduct")
        Dim bOpenedRTConnection As Boolean = False
        Dim dtStatus As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim CMInstalled As Boolean = False
        Dim CPEInsatlled As Boolean = False
        Dim UEInstalled As Boolean = False
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
            _GetOffersByProduct = ResultSet
            Exit Function
        End If
        If ProductId = "" Then
            'ProductID Value is empty
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_PRODUCTID
            row.Item("Description") = "Failure: Invalid ProductID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _GetOffersByProduct = ResultSet
            Exit Function
        End If
        If ProductType = -1 Then
            'ProductType is empty
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_PRODUCTTYPE
            row.Item("Description") = "Failure: Invalid ProductType"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _GetOffersByProduct = ResultSet
            Exit Function
        End If
        If GUID.Contains("'") = True Or GUID.Contains(Chr(34)) = True Then
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_GUID
            row.Item("Description") = "Failure: Invalid GUID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _GetOffersByProduct = ResultSet
            Exit Function
        End If
        If ProductId.Contains("'") = True Or ProductId.Contains(Chr(34)) = True Then
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_PRODUCTID
            row.Item("Description") = "Failure: Invalid ProductID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _GetOffersByProduct = ResultSet
            Exit Function
        End If
        Dim dtOffersChannel As System.Data.DataTable
        Dim dt As DataTable
        Dim ProductGroupID As Long = 0
        Dim ProdGroupIDs As String = ""
        Dim dtOffersByProduct As System.Data.DataTable = New DataTable()
        Dim tempOffersByProductCM As DataTable = New DataTable()
        Dim tempOffersByProduct As DataTable = New DataTable()
        Dim selectQueryCM As StringBuilder = New StringBuilder()
        Dim joinQueryCM As StringBuilder = New StringBuilder()
        Dim whereQueryCM As StringBuilder = New StringBuilder()
        Dim selectQuery As StringBuilder = New StringBuilder()
        Dim joinQuery As StringBuilder = New StringBuilder()
        Dim whereQuery As StringBuilder = New StringBuilder()
        Try

            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                bOpenedRTConnection = True
                MyCommon.Open_LogixRT()
            End If

            If Not IsValidGUID(GUID) Then
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
            Else
                'Fetch product id based on ProductId and producttype id from products table
                MyCommon.QueryStr = "select ProductID from products with (NoLock) " & _
                                            "where ExtProductID='" & ProductId.ToString & "' and ProductTypeID=" & ProductType.ToString & " ;"

                dt = MyCommon.LRT_Select
                If dt.Rows.Count > 0 Then
                    'Product ID found
                    ProductGroupID = dt.Rows(0)(0)
                    'Fetch Product Group id based on the product Id from ProductGroupItems
                    MyCommon.QueryStr = "select PG.ProductGroupID from prodgroupitems PGI with (NoLock) Inner Join ProductGroups PG with (NoLock) " & _
                                        "on PG.ProductGroupID=PGI.ProductGroupID where PGI.ProductID='" & ProductGroupID & "' and PGI.Deleted=0 " & _
                                        " and PG.Deleted=0"
                    dt = MyCommon.LRT_Select
                    If dt.Rows.Count > 0 Then
                        ProdGroupIDs = "("
                        For intIndex As Integer = 0 To dt.Rows.Count - 1
                            If intIndex = 0 Then
                                ProdGroupIDs = ProdGroupIDs & dt.Rows(intIndex)(0)
                            Else
                                ProdGroupIDs = ProdGroupIDs & "," & dt.Rows(intIndex)(0)
                            End If
                        Next
                        ProdGroupIDs = ProdGroupIDs & ")"
                       
                        CMInstalled = MyCommon.IsEngineInstalled(0)
                        CPEInsatlled = MyCommon.IsEngineInstalled(2)
                        UEInstalled = MyCommon.IsEngineInstalled(9)
                        If (CMInstalled) Then
                            'Fetch the Start date, End date, Offer name, Offer id, Offer reward amount, Offer reward amount type, Folder name from RT DB
                            selectQueryCM.Append("SELECT  O.ProdStartDate, O.ProdEndDate, O.Name AS 'OfferName', O.OfferID ,RT.RewardAmount AS 'OfferRewardAmount', AT.Description as 'OfferRewardAmountType', " & _
                                                " isnull(F.FolderName,'') as 'FolderName', isnull(L.LocationName,'') as 'OfferLocation', isnull(CG.Name,'') as 'CustomerGroup' ")
                            If (MyCommon.Fetch_CM_SystemOption(85) = "1") Then
                                selectQueryCM.Append(" ,OAF.DisplayStartDate, OAF.DisplayEndDate ")
                            End If
                            joinQueryCM.Append("From Offers O WITH (NoLock) " & _
                                             "inner join (Offerrewards ORS WITH (NoLock) " & _
                                                                   " inner join RewardTiers RT WITH (NoLock) on ORS.RewardID= RT.RewardID " & _
                                                                   " inner join AmountTypes AT WITH (NoLock) on ORS.RewardAmountTypeID = AT.AmountTypeID and ORS.Deleted=0 " & _
                                                                   " )on O.OfferID= ORS.OfferID and ORS.ProductGroupID in " & ProdGroupIDs & " inner join (OfferConditions OC  WITH (NoLock) " & _
                                                                   " inner join CustomerGroups CG WITH (NoLock) on OC.LinkID= CG.CustomerGroupID and OC.Deleted=0 and CG.Deleted=0) on OC.OfferID=O.OfferID " & _
                                               " inner join (OfferLocations OL WITH (NoLock) " & _
                                                   " inner Join  (LocGroupItems LGI WITH (NoLock) inner join Locations L WITH (NoLock) " & _
                                                                               " on LGI.LocationID=L.LocationID and LGI.Deleted=0 and L.Deleted=0 " & _
                                                                       "  ) on OL.LocationGroupID=LGI.LocationGroupID and OL.Deleted=0 " & _
                                                                   " )on O.OfferID = OL.OfferID  " & _
                                               " left outer join (FolderItems FI WITH (NoLock) " & _
                                                                               " inner join Folders F with (NoLock) on Fi.FolderID=f.FolderID) " & _
                                               " on FI.LinkID=O.OfferID and o.istemplate=0 ")
                            If (MyCommon.Fetch_CM_SystemOption(85) = "1") Then
                                joinQueryCM.Append(" inner join OfferAccessoryFields OAF with (NOLOCK) on OAF.OfferID = O.OfferID")
                            End If
                            whereQueryCM.Append(" where o.Deleted=0 and DATEDIFF(DD,GETDATE(),O.ProdEndDate)>=0 ORDER BY O.OfferID")
                            MyCommon.QueryStr = selectQueryCM.ToString() & joinQueryCM.ToString() & whereQueryCM.ToString()
                            tempOffersByProductCM = MyCommon.LRT_Select
                            If tempOffersByProductCM.Rows.Count > 0 Then
                                dtOffersByProduct = tempOffersByProductCM
                            End If
                        End If
                        If (CPEInsatlled) Or (UEInstalled) Then
                            selectQuery.Append("select CI.StartDate,CI.EndDate,CI.IncentiveName as 'OfferName',CI.IncentiveID as 'OfferID', CD.DiscountAmount  as 'OfferRewardAmount'," & _
                                                " CAT.Name as 'OfferRewardAmountType',isnull(F.FolderName,'') as 'FolderName',isnull(LG.Name,'') as 'OfferLocation',isnull(CG.Name,'') as 'CustomerGroup'")
                            If (MyCommon.Fetch_UE_SystemOption(143) = "1") Then
                                selectQuery.Append(",OAF.DisplayStartDate, OAF.DisplayEndDate ")
                            End If
                            
                            joinQuery.Append("FROM CPE_Incentives CI WITH (NOLOCK) Left Outer JOIN (CPE_RewardOptions CR WITH (NOLOCK)" & _
                                                " Left Outer JOIN ( CPE_Deliverables DL with (NoLock) INNER JOIN" & _
                                                " (CPE_Discounts CD with (NoLock) INNER JOIN CPE_AmountTypes cat with (NoLock)" & _
                                                " on CD.AmountTypeID =cat.AmountTypeID) on DL.OutputID=CD.DiscountID" & _
                                                " and DL.RewardOptionPhase=3 and DL.DeliverableTypeID=2 )on CR.RewardOptionID=DL.RewardOptionID" & _
                                                " and CR.TouchResponse=0 and CR.Deleted=0 " & _
                                                " LEFT OUTER JOIN (CPE_IncentiveCustomerGroups CICG Inner Join CustomerGroups CG" & _
                                                " on CICG.CustomerGroupID = CG.CustomerGroupID )On CR.RewardOptionID=CICG.RewardOptionID" & _
                                                " )ON CI.IncentiveId=CR.IncentiveID" & _
                                                " LEFT OUTER JOIN (FolderItems FI with (NoLock) inner join Folders F with (NoLock)" & _
                                                " on Fi.FolderID=f.FolderID)on FI.LinkID=ci.IncentiveID" & _
                                                " LEFT OUTER JOIN (OfferLocations OL WITH (NOLOCK) INNER JOIN LocationGroups LG WITH (NOLOCK)" & _
                                                " ON OL.LocationGroupID=LG.LocationGroupID AND isnull(OL.Deleted,0)=0 and isNull(LG.Deleted,0)=0)" & _
                                                " ON CI.IncentiveId=OL.OfferID")
                            If (MyCommon.Fetch_UE_SystemOption(143) = "1") Then
                                joinQuery.Append(" inner join OfferAccessoryFields OAF with (NOLOCK) on OAF.OfferID = CI.IncentiveId ")
                            End If
                            whereQuery.Append(" WHERE CI.Deleted=0 AND CI.istemplate=0" & _
                                               " AND DATEDIFF(DD,GETDATE(),CI.EndDate)>=0 and CD.DiscountedProductGroupID in " & ProdGroupIDs & " ORDER BY CI.IncentiveId")
                            MyCommon.QueryStr = selectQuery.ToString() & joinQuery.ToString() & whereQuery.ToString()
                            tempOffersByProduct = MyCommon.LRT_Select
                            If tempOffersByProduct.Rows.Count > 0 Then
                                If dtOffersByProduct.Rows.Count > 0 Then
                                    For Each tempRow As DataRow In tempOffersByProduct.Rows
                                        dtOffersByProduct.ImportRow(tempRow)
                                    Next
                                Else
                                    dtOffersByProduct = tempOffersByProduct
                                End If
                            End If
                        End If
            
                        If dtOffersByProduct.Rows.Count > 0 Then
                            '' Offers information found
                            row = dtStatus.NewRow()
                            row.Item("StatusCode") = StatusCodes.SUCCESS
                            row.Item("Description") = "Success:"
                            dtStatus.Rows.Add(row)
                            dtStatus.AcceptChanges()
                            ResultSet.Tables.Add(dtStatus.Copy())
                            dtOffersByProduct.TableName = "OffersByProduct"
                            dtOffersByProduct.AcceptChanges()
                            ResultSet.Tables.Add(dtOffersByProduct.Copy())

                            Dim DisplayOfferAd As Integer = 0
                            Try
                                DisplayOfferAd = MyCommon.Extract_Val(MyCommon.Fetch_SystemOption(129))
                            Catch ex As Exception
                                DisplayOfferAd = 0
                            End Try
                            If DisplayOfferAd = 1 Then
                                Try
                                    Dim OfferIDs As String = ""
                                    Dim dtOffersAd As System.Data.DataTable
                                    For i As Integer = 0 To dtOffersByProduct.Rows.Count - 1
                                        If i = 0 Then
                                            OfferIDs = "(" & dtOffersByProduct.Rows(i)("OfferID") & ","
                                        Else
                                            If OfferIDs.Contains(dtOffersByProduct.Rows(i)("OfferID") & ",") = False Then
                                                OfferIDs = OfferIDs & dtOffersByProduct.Rows(i)("OfferID") & ","
                                            End If
                                        End If
                                    Next
                                    If OfferIDs <> "" AndAlso DisplayOfferAd Then
                                        OfferIDs = Mid(OfferIDs, 1, Len(OfferIDs) - 1) & ")"
                                        MyCommon.QueryStr = "Select OAF.OfferID, OAF.CopyText,isnull(ADC.AdFieldDescription,'') 'CoverageMethod'," & _
                                                            " isnull(OAF.Page,0)'Page',isnull(OAF.Block,0)'Block'," & _
                                                            " isnull(ADS.AdFieldDescription,'FLYER') 'SaleEventType' from OfferAdFields OAF with (nolock)" & _
                                                            " Left Outer join AdDetails ADC with (nolock) on OAF.CoverageMethodID=ADC.AdFieldValue" & _
                                                            " Left Outer join AdDetails ADS with (nolock) on OAF.SaleEventTypeID=ADS.AdFieldValue " & _
                                                            " where OAF.OfferID in " & OfferIDs & " and OAF.Deleted=0"
                                        dtOffersAd = MyCommon.LRT_Select
                                        dtOffersAd.TableName = "OfferAdFields"
                                        dtOffersAd.AcceptChanges()
                                        ResultSet.Tables.Add(dtOffersAd.Copy())
                                    End If
                                    If OfferIDs <> "" Then
                                        MyCommon.QueryStr = "select MediaTypeID, CAST(CAST(N'' AS XML).value('xs:base64Binary(sql:column(""MediaData""))','VARBINARY(MAX)') AS VARCHAR(MAX)) as MediaData, MediaFormatID, LanguageID " & _
                                                               "from ChannelOfferAssets with (NoLock) " & _
                                                               "where OfferID in " & OfferIDs & " and ChannelID=5;"
                                        dtOffersChannel = MyCommon.LRT_Select
                                        dtOffersChannel.TableName = "OfferChannelFields"
                                        dtOffersChannel.AcceptChanges()
                                        ResultSet.Tables.Add(dtOffersChannel.Copy())
                                    End If
                                Catch ex As Exception
                                End Try
                            End If
                        Else
                            'No records found
                            row = dtStatus.NewRow()
                            row.Item("StatusCode") = StatusCodes.NOTFOUND_RECORDS
                            row.Item("Description") = "Failure: NOT Found Records"
                            dtStatus.Rows.Add(row)
                            dtStatus.AcceptChanges()
                            ResultSet.Tables.Add(dtStatus.Copy())
                        End If
                    Else
                        'Product Group not found for the input parameters
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_PRODUCTGROUP
                        row.Item("Description") = "Failure: NOT Found ProductGoup"
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    End If
                Else
                    'Invalid Product ID or Invalid Product Group ID
                    MyCommon.QueryStr = "select ProductID from products with (NoLock) " & _
                                        "where ExtProductID='" & ProductId.ToString & "';"

                    dt = MyCommon.LRT_Select
                    If dt.Rows.Count > 0 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.INVALID_PRODUCTTYPE
                        row.Item("Description") = "Failure: Invalid ProductType"
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    Else
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.INVALID_PRODUCTID
                        row.Item("Description") = "Failure: Invalid ProductID"
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    End If

                End If
            End If
        Catch ex As Exception
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
            row.Item("Description") = "Failure: Application " & ex.ToString
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
        Finally
            If bOpenedRTConnection = True And MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try
        Return ResultSet
    End Function

    <WebMethod()> _
    Public Function GetActiveOffersByProduct(ByVal GUID As String, ByVal ProductId As String, ByVal ProductType As String) As DataSet
        Dim iProductType As Integer = -1
        Try
            InitApp()
            iProductType = CInt(ProductType)
        Catch ex As Exception
            iProductType = -1
        End Try
        Return _GetActiveOffersByProduct(GUID, ProductId, iProductType)
    End Function
	
    '"AMS_FIS_hbc_J" - <2.3.14>
    Public Function _GetActiveOffersByProduct(ByVal GUID As String, ByVal ProductId As String, ByVal ProductType As Integer) As DataSet
        Dim ResultSet As New System.Data.DataSet("OffersByProduct")
        Dim bOpenedRTConnection As Boolean = False
        Dim dtStatus As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim sFileName As String 
        Dim OfferID As String = ""
        Dim Offerstatus As String = ""
        Dim CMInstalled As Boolean = False
        Dim CPEInsatlled As Boolean = False
        Dim UEInstalled As Boolean = False
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))
        sFileName = sLogFileName & "." & MyCommon.Leading_Zero_Fill(Year(Today), 4) & MyCommon.Leading_Zero_Fill(Month(Today), 2) & MyCommon.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
        If GUID = "" Then
            'GUID Value is empty
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_GUID
            row.Item("Description") = "Failure: Invalid GUID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _GetActiveOffersByProduct = ResultSet
            Exit Function
        End If
        If ProductId = "" Then
            'ProductID Value is empty
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_PRODUCTID
            row.Item("Description") = "Failure: Invalid ProductID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _GetActiveOffersByProduct = ResultSet
            Exit Function
        End If
        If ProductType = -1 Then
            'ProductType is empty
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_PRODUCTTYPE
            row.Item("Description") = "Failure: Invalid ProductType"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _GetActiveOffersByProduct = ResultSet
            Exit Function
        End If
        If GUID.Contains("'") = True Or GUID.Contains(Chr(34)) = True Then
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_GUID
            row.Item("Description") = "Failure: Invalid GUID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _GetActiveOffersByProduct = ResultSet
            Exit Function
        End If
        If ProductId.Contains("'") = True Or ProductId.Contains(Chr(34)) = True Then
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_PRODUCTID
            row.Item("Description") = "Failure: Invalid ProductID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _GetActiveOffersByProduct = ResultSet
            Exit Function
        End If
        Dim dtOffersChannel As System.Data.DataTable
        Dim dt As DataTable
        Dim ProductGroupID As Long = 0
        Dim ProdGroupIDs As String = ""
        Dim dtOffersByProduct As System.Data.DataTable = New DataTable()
        Dim tempOffersByProductCM As DataTable = New DataTable()
        Dim tempOffersByProduct As DataTable = New DataTable()
        Dim selectQueryCM As StringBuilder = New StringBuilder()
        Dim joinQueryCM As StringBuilder = New StringBuilder()
        Dim whereQueryCM As StringBuilder = New StringBuilder()
        Dim selectQuery As StringBuilder = New StringBuilder()
        Dim joinQuery As StringBuilder = New StringBuilder()
        Dim whereQuery As StringBuilder = New StringBuilder()
        Try

            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                bOpenedRTConnection = True
                MyCommon.Open_LogixRT()
            End If

            If Not IsValidGUID(GUID) Then
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
            Else
                'Fetch product id based on ProductId and producttype id from products table
                MyCommon.QueryStr = "select ProductID from products with (NoLock) " & _
                                            "where ExtProductID='" & ProductId.ToString & "' and ProductTypeID=" & ProductType.ToString & " ;"

                dt = MyCommon.LRT_Select
                If dt.Rows.Count > 0 Then
                    'Product ID found
                    ProductGroupID = dt.Rows(0)(0)
                    'Fetch Product Group id based on the product Id from ProductGroupItems
                    MyCommon.QueryStr = "select PG.ProductGroupID from prodgroupitems PGI with (NoLock) Inner Join ProductGroups PG with (NoLock) " & _
                                        "on PG.ProductGroupID=PGI.ProductGroupID where PGI.ProductID='" & ProductGroupID & "' and PGI.Deleted=0 " & _
                                        " and PG.Deleted=0"
                    dt = MyCommon.LRT_Select
                    If dt.Rows.Count > 0 Then
                        ProdGroupIDs = "("
                        For intIndex As Integer = 0 To dt.Rows.Count - 1
                            If intIndex = 0 Then
                                ProdGroupIDs = ProdGroupIDs & dt.Rows(intIndex)(0)
                            Else
                                ProdGroupIDs = ProdGroupIDs & "," & dt.Rows(intIndex)(0)
                            End If
                        Next
                        ProdGroupIDs = ProdGroupIDs & ")"
                       
                        CMInstalled = MyCommon.IsEngineInstalled(0)
                        CPEInsatlled = MyCommon.IsEngineInstalled(2)
                        UEInstalled = MyCommon.IsEngineInstalled(9)
                        If (CMInstalled) Then
                            'Fetch the Start date, End date, Offer name, Offer id, Offer reward amount, Offer reward amount type, Folder name from RT DB
                            selectQueryCM.Append("SELECT  O.ProdStartDate, O.ProdEndDate, O.Name AS 'OfferName', O.OfferID ,RT.RewardAmount AS 'OfferRewardAmount', AT.Description as 'OfferRewardAmountType', " & _
                                                " isnull(F.FolderName,'') as 'FolderName', isnull(L.LocationName,'') as 'OfferLocation', isnull(CG.Name,'') as 'CustomerGroup' ")
                            If (MyCommon.Fetch_CM_SystemOption(85) = "1") Then
                                selectQueryCM.Append(" ,OAF.DisplayStartDate, OAF.DisplayEndDate ")
                            End If
                            joinQueryCM.Append("From Offers O WITH (NoLock) " & _
                                             "inner join (Offerrewards ORS WITH (NoLock) " & _
                                                                   " inner join RewardTiers RT WITH (NoLock) on ORS.RewardID= RT.RewardID " & _
                                                                   " inner join AmountTypes AT WITH (NoLock) on ORS.RewardAmountTypeID = AT.AmountTypeID and ORS.Deleted=0 " & _
                                                                   " )on O.OfferID= ORS.OfferID and ORS.ProductGroupID in " & ProdGroupIDs & " inner join (OfferConditions OC  WITH (NoLock) " & _
                                                                   " inner join CustomerGroups CG WITH (NoLock) on OC.LinkID= CG.CustomerGroupID and OC.Deleted=0 and CG.Deleted=0) on OC.OfferID=O.OfferID " & _
                                               " inner join (OfferLocations OL WITH (NoLock) " & _
                                                   " inner Join  (LocGroupItems LGI WITH (NoLock) inner join Locations L WITH (NoLock) " & _
                                                                               " on LGI.LocationID=L.LocationID and LGI.Deleted=0 and L.Deleted=0 " & _
                                                                       "  ) on OL.LocationGroupID=LGI.LocationGroupID and OL.Deleted=0 " & _
                                                                   " )on O.OfferID = OL.OfferID  " & _
                                               " left outer join (FolderItems FI WITH (NoLock) " & _
                                                                               " inner join Folders F with (NoLock) on Fi.FolderID=f.FolderID) " & _
                                               " on FI.LinkID=O.OfferID and o.istemplate=0 ")
                            If (MyCommon.Fetch_CM_SystemOption(85) = "1") Then
                                joinQueryCM.Append(" inner join OfferAccessoryFields OAF with (NOLOCK) on OAF.OfferID = O.OfferID")
                            End If
                            whereQueryCM.Append(" where o.Deleted=0 and DATEDIFF(DD,GETDATE(),O.ProdEndDate)>=0 ORDER BY O.OfferID")
                            MyCommon.QueryStr = selectQueryCM.ToString() & joinQueryCM.ToString() & whereQueryCM.ToString()
                            tempOffersByProductCM = MyCommon.LRT_Select
                            dtOffersByProduct = tempOffersByProduct.Clone()
                            For Each tempRow As DataRow In tempOffersByProduct.Rows
                                OfferID = tempRow(3)
                                Offerstatus = logixInc.GetOfferStatus(OfferID, 1)
                                If Offerstatus.IndexOf("Active") > 0 Then
                                    dtOffersByProduct.ImportRow(tempRow)
                                    MyCommon.Write_Log(sFileName, "No records are added", True)
                                End If
                            Next
                        End If
                        If (CPEInsatlled) Or (UEInstalled) Then
                            selectQuery.Append("select CI.StartDate,CI.EndDate,CI.IncentiveName as 'OfferName',CI.IncentiveID as 'OfferID', CD.DiscountAmount  as 'OfferRewardAmount'," & _
                                                " CAT.Name as 'OfferRewardAmountType',isnull(F.FolderName,'') as 'FolderName',isnull(LG.Name,'') as 'OfferLocation',isnull(CG.Name,'') as 'CustomerGroup'")
                            If (MyCommon.Fetch_UE_SystemOption(143) = "1") Then
                                selectQuery.Append(",OAF.DisplayStartDate, OAF.DisplayEndDate ")
                            End If
                            
                            joinQuery.Append("FROM CPE_Incentives CI WITH (NOLOCK) Left Outer JOIN (CPE_RewardOptions CR WITH (NOLOCK)" & _
                                                " Left Outer JOIN ( CPE_Deliverables DL with (NoLock) INNER JOIN" & _
                                                " (CPE_Discounts CD with (NoLock) INNER JOIN CPE_AmountTypes cat with (NoLock)" & _
                                                " on CD.AmountTypeID =cat.AmountTypeID) on DL.OutputID=CD.DiscountID" & _
                                                " and DL.RewardOptionPhase=3 and DL.DeliverableTypeID=2 )on CR.RewardOptionID=DL.RewardOptionID" & _
                                                " and CR.TouchResponse=0 and CR.Deleted=0 " & _
                                                " LEFT OUTER JOIN (CPE_IncentiveCustomerGroups CICG Inner Join CustomerGroups CG" & _
                                                " on CICG.CustomerGroupID = CG.CustomerGroupID )On CR.RewardOptionID=CICG.RewardOptionID" & _
                                                " )ON CI.IncentiveId=CR.IncentiveID" & _
                                                " LEFT OUTER JOIN (FolderItems FI with (NoLock) inner join Folders F with (NoLock)" & _
                                                " on Fi.FolderID=f.FolderID)on FI.LinkID=ci.IncentiveID" & _
                                                " LEFT OUTER JOIN (OfferLocations OL WITH (NOLOCK) INNER JOIN LocationGroups LG WITH (NOLOCK)" & _
                                                " ON OL.LocationGroupID=LG.LocationGroupID AND isnull(OL.Deleted,0)=0 and isNull(LG.Deleted,0)=0)" & _
                                                " ON CI.IncentiveId=OL.OfferID")
                            If (MyCommon.Fetch_UE_SystemOption(143) = "1") Then
                                joinQuery.Append(" inner join OfferAccessoryFields OAF with (NOLOCK) on OAF.OfferID = CI.IncentiveId ")
                            End If
                            whereQuery.Append(" WHERE CI.Deleted=0 AND CI.istemplate=0" & _
                                               " AND DATEDIFF(DD,GETDATE(),CI.EndDate)>=0 and CD.DiscountedProductGroupID in " & ProdGroupIDs & " ORDER BY CI.IncentiveId")
                            MyCommon.QueryStr = selectQuery.ToString() & joinQuery.ToString() & whereQuery.ToString()
                            tempOffersByProduct = MyCommon.LRT_Select
                            dtOffersByProduct = tempOffersByProduct.Clone()
                            For Each tempRow As DataRow In tempOffersByProduct.Rows
                                OfferID = tempRow(3)
                                Offerstatus = logixInc.GetOfferStatus(OfferID, 1)
                                If Offerstatus.IndexOf("Active") > 0 Then
                                    dtOffersByProduct.ImportRow(tempRow)
                                     MyCommon.Write_Log(sFileName, "No records are added", True)

                                End If
                            Next                               
                        End If            
                        If dtOffersByProduct.Rows.Count > 0 Then
                            '' Offers information found
                            row = dtStatus.NewRow()
                            row.Item("StatusCode") = StatusCodes.SUCCESS
                            row.Item("Description") = "Success:"
                            dtStatus.Rows.Add(row)
                            dtStatus.AcceptChanges()
                            ResultSet.Tables.Add(dtStatus.Copy())
                            dtOffersByProduct.TableName = "OffersByProduct"
                            dtOffersByProduct.AcceptChanges()
                            ResultSet.Tables.Add(dtOffersByProduct.Copy())
                           
                            Dim DisplayOfferAd As Integer = 0
                            Try
                                DisplayOfferAd = MyCommon.Extract_Val(MyCommon.Fetch_SystemOption(129))
                            Catch ex As Exception
                                DisplayOfferAd = 0
                            End Try
                            If DisplayOfferAd = 1 Then
                                Try
                                    Dim OfferIDs As String = ""
                                    Dim dtOffersAd As System.Data.DataTable
                                    For i As Integer = 0 To dtOffersByProduct.Rows.Count - 1
                                        If i = 0 Then
                                            OfferIDs = "(" & dtOffersByProduct.Rows(i)("OfferID") & ","
                                        Else
                                            If OfferIDs.Contains(dtOffersByProduct.Rows(i)("OfferID") & ",") = False Then
                                                OfferIDs = OfferIDs & dtOffersByProduct.Rows(i)("OfferID") & ","
                                            End If
                                        End If
                                    Next
                                    If OfferIDs <> "" AndAlso DisplayOfferAd Then
                                        OfferIDs = Mid(OfferIDs, 1, Len(OfferIDs) - 1) & ")"
                                        MyCommon.QueryStr = "Select OAF.OfferID, OAF.CopyText,isnull(ADC.AdFieldDescription,'') 'CoverageMethod'," & _
                                                            " isnull(OAF.Page,0)'Page',isnull(OAF.Block,0)'Block'," & _
                                                            " isnull(ADS.AdFieldDescription,'FLYER') 'SaleEventType' from OfferAdFields OAF with (nolock)" & _
                                                            " Left Outer join AdDetails ADC with (nolock) on OAF.CoverageMethodID=ADC.AdFieldValue" & _
                                                            " Left Outer join AdDetails ADS with (nolock) on OAF.SaleEventTypeID=ADS.AdFieldValue " & _
                                                            " where OAF.OfferID in " & OfferIDs & " and OAF.Deleted=0"
                                        dtOffersAd = MyCommon.LRT_Select
                                        dtOffersAd.TableName = "OfferAdFields"
                                        dtOffersAd.AcceptChanges()
                                        ResultSet.Tables.Add(dtOffersAd.Copy())
                                    End If
                                    If OfferIDs <> "" Then
                                        MyCommon.QueryStr = "select MediaTypeID, CAST(CAST(N'' AS XML).value('xs:base64Binary(sql:column(""MediaData""))','VARBINARY(MAX)') AS VARCHAR(MAX)) as MediaData, MediaFormatID, LanguageID " & _
                                                               "from ChannelOfferAssets with (NoLock) " & _
                                                               "where OfferID in " & OfferIDs & " and ChannelID=5;"
                                        dtOffersChannel = MyCommon.LRT_Select
                                        dtOffersChannel.TableName = "OfferChannelFields"
                                        dtOffersChannel.AcceptChanges()
                                        ResultSet.Tables.Add(dtOffersChannel.Copy())
                                    End If
                                Catch ex As Exception
                                End Try
                            End If
                        Else
                            'No records found
                            row = dtStatus.NewRow()
                            row.Item("StatusCode") = StatusCodes.NOTFOUND_RECORDS
                            row.Item("Description") = "Failure: NOT Found Records"
                            dtStatus.Rows.Add(row)
                            dtStatus.AcceptChanges()
                            ResultSet.Tables.Add(dtStatus.Copy())
                        End If
                    Else
                        'Product Group not found for the input parameters
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_PRODUCTGROUP
                        row.Item("Description") = "Failure: NOT Found ProductGoup"
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    End If
                Else
                    'Invalid Product ID or Invalid Product Group ID
                    MyCommon.QueryStr = "select ProductID from products with (NoLock) " & _
                                        "where ExtProductID='" & ProductId.ToString & "';"

                    dt = MyCommon.LRT_Select
                    If dt.Rows.Count > 0 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.INVALID_PRODUCTTYPE
                        row.Item("Description") = "Failure: Invalid ProductType"
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    Else
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.INVALID_PRODUCTID
                        row.Item("Description") = "Failure: Invalid ProductID"
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    End If

                End If
            End If
        Catch ex As Exception
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
            row.Item("Description") = "Failure: Application " & ex.ToString
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
        Finally
            If bOpenedRTConnection = True And MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try
        Return ResultSet
    End Function
    
    <WebMethod()> _
    Public Function GetCardAvailabilityInCustomerGroup(ByVal GUID As String, ByVal CardId As String, ByVal CardTypeID As String, ByVal CustomerGroupId As String) As DataSet
        Dim iCardTypeId As Integer = -1
        Dim lcustomergroupid As Long = -1

        InitApp()

        'If Not String.IsNullOrEmpty(CardTypeID) AndAlso Not CardTypeID.Trim = String.Empty AndAlso IsNumeric(CardTypeID) Then iCardTypeId = Convert.ToInt32(CardTypeID)
        'Bug 2448
        If String.IsNullOrEmpty(CardTypeID) Or CardTypeID.Trim = String.Empty Then
            iCardTypeId = -100
        ElseIf Not IsNumeric(CardTypeID) Then
            iCardTypeId = -1
        Else
            If Convert.ToInt32(CardTypeID) < 0 Then
                iCardTypeId = -1
            Else
                iCardTypeId = Convert.ToInt32(CardTypeID)
            End If
        End If

        If Not String.IsNullOrEmpty(CustomerGroupId) AndAlso Not CustomerGroupId.Trim = String.Empty AndAlso IsNumeric(CustomerGroupId) Then lcustomergroupid = Convert.ToInt64(CustomerGroupId)

        Return _GetCardAvailabilityInCustomerGroup(GUID, CardId, iCardTypeId, lcustomergroupid)

    End Function

    Private Function _GetCardAvailabilityInCustomerGroup(ByVal GUID As String, ByVal CardId As String, ByVal CardTypeId As Integer, ByVal customergroupid As Long) As DataSet
     
        Dim Retmsg As String = ""
        Dim row, CustomerPKrow As DataRow
        Dim dtstatus, dt As DataTable
        Dim dtcardavailability As DataTable = Nothing
        Dim resultset As New DataSet
        Dim dtCustomerPK As DataTable
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim lCustomerPKtemp As Long
        'Initialize the status table, which will report the success or failure of the operation
        dtstatus = New DataTable
        dtstatus.TableName = "Status"
        dtstatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtstatus.Columns.Add("Description", System.Type.GetType("System.String"))
        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            If Not IsValidGUID(GUID) Then
                'invalid GUID
                row = dtstatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtstatus.Rows.Add(row)

            ElseIf CardId.Length < 1 Then
                'Bad customer ID
                row = dtstatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERID
                row.Item("Description") = "Failure: Cards ID not provided."
                dtstatus.Rows.Add(row)
				
            ElseIf CardId.Contains(Chr(34)) = True Or CardId.Contains("'") = True Then
                'Card Id empty
                row = dtstatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERID
                row.Item("Description") = "Failure: Invalid CustomerID"
                dtstatus.Rows.Add(row)

            ElseIf CardId.Contains(Chr(34)) = True Or CardId.Contains("'") = True Then
                'Card Id empty
                row = dtstatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERID
                row.Item("Description") = "Failure: Invalid CustomerID"
                dtstatus.Rows.Add(row)

            ElseIf CardTypeId = -1 Then
                'Bad Card TypeID
                row = dtstatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERTYPEID
                row.Item("Description") = "Failure: Invalid CardType."
                dtstatus.Rows.Add(row)
			
            ElseIf CardTypeId = -100 Then
                'Bad Card TypeID
                row = dtstatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERTYPEID
                row.Item("Description") = "Failure: Card Type ID not provided."
                dtstatus.Rows.Add(row)
				
            ElseIf customergroupid = -1 Then
                'Bad Card TypeID
                row = dtstatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERGROUPID
                row.Item("Description") = "Failure: CustomerGroupID not provided."
                dtstatus.Rows.Add(row)
            Else
                CardId = MyCommon.Pad_ExtCardID(CardId, CardTypeId)

                MyCommon.QueryStr = "select CustomerPK from CardIDs with (NoLock) " & _
                                  "where ExtCardID='" & MyCryptLib.SQL_StringEncrypt(CardId, True) & "' and CardTypeID=" & CardTypeId.ToString & ";"

                dt = MyCommon.LXS_Select
                If dt.Rows.Count > 0 Then
                    dtCustomerPK = dt
                    'if customer id not found
                ElseIf CardTypeId = 0 Then
                    ''For LoyaltyCards
                    RetCode = StatusCodes.NOTFOUND_CUSTOMER
                    Retmsg = "Failure: Customer " & CardId & " not found."
                ElseIf CardTypeId = 1 Then
                    ''For HouseHold Cards
                    RetCode = StatusCodes.NOTFOUND_HOUSEHOLD
                    Retmsg = "Failure: Household " & CardId & " not found."
                ElseIf CardTypeId = 2 Then
                    ''For CAM Cards
                    RetCode = StatusCodes.NOTFOUND_CAM
                    Retmsg = "Failure: CAM " & CardId & " not found."
                Else
                    ''For any other card type except of card type 0,1,2
                    RetCode = StatusCodes.INVALID_CUSTOMERTYPEID
                    Retmsg = "Failure: Invalid customer type."
                End If

                If RetCode = StatusCodes.SUCCESS Then
                    dtcardavailability = New DataTable("CustomerGroups")
                    dtcardavailability.Columns.Add("Status", System.Type.GetType("System.String"))
                    ''check for existence of customer in provided customer grouop
                    For Each CustomerPKrow In dtCustomerPK.Rows
                        lCustomerPKtemp = MyCommon.NZ(CustomerPKrow.Item("CustomerPK"), 0)
                        MyCommon.QueryStr = "select CustomerGroupID from groupmembership with (Nolock) where deleted=0 and CustomerPK=" & lCustomerPKtemp.ToString & " and CustomerGroupID=" & customergroupid.ToString & ";"
                        row = dtcardavailability.NewRow()
                        If MyCommon.LXS_Select.Rows.Count > 0 Then
                            row.Item(0) = "True"
                            dtcardavailability.Rows.Add(row)
                        Else
                            row.Item(0) = "False"
                            dtcardavailability.Rows.Add(row)
                        End If
                    Next
                    If dtcardavailability.Rows.Count > 0 Then dtcardavailability.AcceptChanges()

                    row = dtstatus.NewRow()
                    row.Item("StatusCode") = StatusCodes.SUCCESS
                    row.Item("Description") = "Success."
                    dtstatus.Rows.Add(row)
                Else     ''''If customer not found then add the row to status table
                    row = dtstatus.NewRow()
                    row.Item("StatusCode") = RetCode
                    row.Item("Description") = Retmsg
                    dtstatus.Rows.Add(row)
                End If
            End If
            If dtstatus IsNot Nothing AndAlso dtstatus.Rows.Count > 0 Then
                dtstatus.AcceptChanges()
                resultset.Tables.Add(dtstatus.Copy())
            End If
            If dtcardavailability IsNot Nothing Then resultset.Tables.Add(dtcardavailability.Copy())

        Catch ex As Exception
            row = dtstatus.NewRow()
            row.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
            row.Item("Description") = "Failure: Application " & ex.ToString
            dtstatus.Rows.Add(row)
            dtstatus.AcceptChanges()
            resultset.Tables.Add(dtstatus.Copy())
        Finally
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try
        Return resultset
    End Function


    <WebMethod()> _
    Public Function GetCustomerGroupRecordByCard(ByVal GUID As String, ByVal CardId As String, ByVal CardTypeID As String) As DataSet
        'Dim iCardTypeId As Integer = -1
        'Try
        '  InitApp()
        '  iCardTypeId = CInt(CardTypeID)
        'Catch ex As Exception
        '  iCardTypeId = -1
        'End Try
        'Return _GetCustomerGroupRecordByCard(GUID, CardId, iCardTypeId)
        Return _GetCustomerGroupRecordByCard(GUID, CardId, CardTypeID)
    End Function

    Private Function _GetCustomerGroupRecordByCard(ByVal GUID As String, ByVal CardId As String, ByVal CardTypeId As String) As DataSet
        Dim dtStatus As DataTable
        Dim dtcustgroupinfo As DataTable = Nothing
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim row, CustomerPKrow, dr As DataRow
        Dim ResultSet As New System.Data.DataSet("GetCustomerGroupRecordByCard")
        Dim dt As DataTable
        Dim dtCustomerPK As DataTable
        Dim RetMsg As String = ""
        Dim lCustomerPKtemp As Long
        Dim CustomerGroupID As Long
        Dim CustomerGroupName As String

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))
        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            If Not IsValidGUID(GUID) Then
                'invalid GUID
                RetCode = StatusCodes.INVALID_GUID
                RetMsg = "Failure: Invalid GUID."
                row = dtStatus.NewRow()
                row.Item("StatusCode") = RetCode
                row.Item("Description") = RetMsg
                dtStatus.Rows.Add(row)
                
            ElseIf Not IsValidCustomerCard(CardId, CardTypeId, RetCode, RetMsg) Then
                
                'Bad customer ID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = RetCode
                row.Item("Description") = RetMsg
                dtStatus.Rows.Add(row)
                
                
            Else
                Dim iCardTypeId As Integer = -1
                Try
                    InitApp()
                    iCardTypeId = CInt(CardTypeId)
                Catch ex As Exception
                    iCardTypeId = -1
                End Try
                CardId = MyCommon.Pad_ExtCardID(CardId, iCardTypeId)

                MyCommon.QueryStr = "select CustomerPK from CardIDs with (NoLock) " & _
                          "where ExtCardID='" & MyCryptLib.SQL_StringEncrypt(CardId, True) & "' and CardTypeID=" & CardTypeId & ";"

                dt = MyCommon.LXS_Select
                If dt.Rows.Count > 0 Then
                    dtCustomerPK = dt
                    'if customer id not found
                    ''For Loyalty Cards
                ElseIf iCardTypeId = 0 Then
                    RetCode = StatusCodes.NOTFOUND_CUSTOMER
                    RetMsg = "Failure: Customer " & CardId & " not found."
                    ''For HouseHold Cards
                ElseIf iCardTypeId = 1 Then
                    RetCode = StatusCodes.NOTFOUND_HOUSEHOLD
                    RetMsg = "Failure: Household " & CardId & " not found."
                    ''For CAM Cards
                ElseIf iCardTypeId = 2 Then
                    RetCode = StatusCodes.NOTFOUND_CAM
                    RetMsg = "Failure: CAM " & CardId & " not found."
                Else
                    ''For any other card type except of card type 0,1,2
                    RetCode = StatusCodes.INVALID_CUSTOMERTYPEID
                    RetMsg = "Failure: Customer not found."
                End If
                If RetCode = StatusCodes.SUCCESS Then
                    dtcustgroupinfo = New DataTable("CustomerGroups")
                    dtcustgroupinfo.Columns.Add("CustomerGroupId", System.Type.GetType("System.Int64"))
                    dtcustgroupinfo.Columns.Add("CustomerGroupName", System.Type.GetType("System.String"))
                    ''For getting corresponding CustomerGroupID,CustomerGroupName mapped with the provided Customer
                    For Each CustomerPKrow In dtCustomerPK.Rows
                        lCustomerPKtemp = MyCommon.NZ(CustomerPKrow.Item("CustomerPK"), 0)
                        MyCommon.QueryStr = " select CustomerGroupID from GroupMembership with (Nolock) " & _
                                            " where CustomerPK=" & lCustomerPKtemp.ToString & " and deleted=0;"
                        dt = MyCommon.LXS_Select
                        If dt.Rows.Count > 0 Then
                            For Each dr In dt.Rows
                                CustomerGroupID = dr.Item(0)
                                MyCommon.QueryStr = "select Name from CustomerGroups with(Nolock) where CustomerGroupid = " & CustomerGroupID & " and deleted=0"
                                dt=MyCommon.LRT_Select
                                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                                CustomerGroupName = MyCommon.NZ(dt.Rows(0).Item("Name"), "")
                                row = dtcustgroupinfo.NewRow()
                                row.Item("CustomerGroupId") = CustomerGroupID
                                row.Item("CustomerGroupName") = CustomerGroupName
                                dtcustgroupinfo.Rows.Add(row)
                                End If
                             
                            Next
                        End If

                    Next
                    If dtcustgroupinfo.Rows.Count > 0 Then
                        dtcustgroupinfo.AcceptChanges()
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.SUCCESS
                        row.Item("Description") = "Success."
                        dtStatus.Rows.Add(row)
                    Else
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_RECORDS
                        row.Item("Description") = "Failure: NOT Found Records"
                        dtStatus.Rows.Add(row)
                    End If

                Else     ''''If customer not found the add the row to status table
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = RetCode
                    row.Item("Description") = RetMsg
                    dtStatus.Rows.Add(row)
                End If

            End If
            If dtStatus IsNot Nothing AndAlso dtStatus.Rows.Count > 0 Then
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
            End If
            If dtcustgroupinfo IsNot Nothing Then ResultSet.Tables.Add(dtcustgroupinfo.Copy())

        Catch ex As Exception
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
            row.Item("Description") = "Failure: Application " & ex.ToString
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
        Finally
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try
        Return ResultSet
    End Function

    <WebMethod()> _
    Public Function GroupMembershipUpdate(ByVal GUID As String, ByVal CardID As String, ByVal CustomerGroupID As Long, ByVal Operation As Integer) As String
        Dim sStatus As String = "Ok"
        Dim sError As String = "Error: "
        Dim sDefaultFirstName As String
        Dim sDefaultLastName As String
        Dim iStatus As Integer
        Dim IsValid As Boolean = False
        Dim ConnInc As New Copient.ConnectorInc
        'added
        Dim RetCode As StatusCodes
        Dim RetMsg As String = ""
        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            IsValid = ConnInc.IsValidConnectorGUID(MyCommon, 3, GUID)
            If IsValid Then
                If IsValidCustomerCard(CardID, 0, RetCode, RetMsg) Then
                    If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
                    Select Case Operation
                        Case 1
                            sDefaultFirstName = MyCommon.Fetch_CM_SystemOption(20)
                            sDefaultLastName = MyCommon.Fetch_CM_SystemOption(21)

                            MyCommon.QueryStr = "dbo.pa_LogixServ_GroupMembership_Insert"
                            MyCommon.Open_LXSsp()
                            MyCommon.LXSsp.Parameters.Add("@ExtCustomerID", SqlDbType.NVarChar, 400).Value = MyCryptLib.SQL_StringEncrypt(CardID, True)
                            MyCommon.LXSsp.Parameters.Add("@CardTypeID", SqlDbType.Int).Value = 0
                            MyCommon.LXSsp.Parameters.Add("@CustomerGroupID", SqlDbType.BigInt, 8).Value = CustomerGroupID
                            MyCommon.LXSsp.Parameters.Add("@LogixTransNum", SqlDbType.Char, 36).Value = ""
                            MyCommon.LXSsp.Parameters.Add("@FirstName", SqlDbType.NVarChar, 50).Value = sDefaultFirstName
                            MyCommon.LXSsp.Parameters.Add("@LastName", SqlDbType.NVarChar, 50).Value = sDefaultLastName
                            MyCommon.LXSsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
                            MyCommon.LXSsp.Parameters.Add("@ExtCustomerIDOriginal", SqlDbType.NVarChar, 400).Value = MyCryptLib.SQL_StringEncrypt(CardID)
                            MyCommon.LXSsp.ExecuteNonQuery()
                            iStatus = MyCommon.LXSsp.Parameters("@Status").Value
                            MyCommon.Close_LXSsp()
                            If iStatus <> 0 Then
                                ''WriteLog("Customer is already a member of Group: '" & CustomerGroupID & "'", MessageType.Info)
                                sStatus = sError & "Customer is already a member of Group: '" & CustomerGroupID & "'"
                            End If
                            UpdateCustomerCount(CardID, 0)
                        Case -1
                            MyCommon.QueryStr = "dbo.pa_LogixServ_GroupMembership_Delete"
                            MyCommon.Open_LXSsp()
                            MyCommon.LXSsp.Parameters.Add("@ExtCustomerID", SqlDbType.NVarChar, 400).Value = MyCryptLib.SQL_StringEncrypt(CardID, True)
                            MyCommon.LXSsp.Parameters.Add("@CardTypeID", SqlDbType.Int).Value = 0
                            MyCommon.LXSsp.Parameters.Add("@CustomerGroupID", SqlDbType.BigInt, 8).Value = CustomerGroupID
                            MyCommon.LXSsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
                            MyCommon.LXSsp.ExecuteNonQuery()
                            iStatus = MyCommon.LXSsp.Parameters("@Status").Value
                            MyCommon.Close_LXSsp()
                            If iStatus <> 0 Then
                                'WriteLog("Customer was not a member of Group: '" & CustomerGroupID & "'", MessageType.Info)
                                sStatus = sError & "Customer was not a member of Group: '" & CustomerGroupID & "'"
                            End If
                            UpdateCustomerCount(CardID, 0)
                        Case Else
                            sStatus = sError & "Invalid Operation"
                    End Select
                Else
                    sStatus = sError & "Invalid CardID"
                End If
            Else
                sStatus = sError & "Invalid GUID"
            End If
        Catch ex As Exception
            sStatus = sError & ex.Message
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try

        Return sStatus
    End Function
  
    Private Sub UpdateCustomerCount(ByVal sId As String, ByVal iCardTypeId As Integer)
        Dim iStatus As Integer = 0
        Dim CmInstalled As Boolean = MyCommon.IsEngineInstalled(0)

        If CmInstalled Then
            MyCommon.QueryStr = "dbo.pa_LogixServ_CustCountUpdate"
            MyCommon.Open_LXSsp()
            MyCommon.LXSsp.Parameters.Add("@CardID", SqlDbType.NVarChar, 400).Value = MyCryptLib.SQL_StringEncrypt(sId)
            MyCommon.LXSsp.Parameters.Add("@CardTypeID", SqlDbType.Int).Value = iCardTypeId
            MyCommon.LXSsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
            MyCommon.LXSsp.ExecuteNonQuery()
            iStatus = MyCommon.LXSsp.Parameters("@Status").Value
            MyCommon.Close_LXSsp()
            If iStatus <> 0 Then
                Throw New ApplicationException("Card NOT found for Card ID '" & sId & "' of type '" & iCardTypeId & "' in procedure 'pa_LogixServ_CustCountUpdate'")
            End If
        End If
    End Sub

	<WebMethod()> _
    Public Function MembershipEditByGroupName(ByVal GUID As String, ByVal Mode As String, ByVal CustomerGroupName As String, ByVal ExtCardId As String, ByVal CardTypeID As Integer) As System.Data.DataSet
      InitApp()
	
      Return _MembershipEditByGroupName(GUID, Mode, CustomerGroupName, ExtCardId, CardTypeID)
    End Function
	
    <WebMethod()> _
    Public Function MembershipEdit(ByVal GUID As String, ByVal Mode As String, ByVal CustomerGroupID As String, ByVal CustomerID As String, ByVal CustomerTypeID As String) As System.Data.DataSet
        InitApp()
        'added
        'Dim iCustomerTypeID As Integer = -1
        'If Not String.IsNullOrEmpty(CustomerTypeID) AndAlso Not CustomerTypeID.Trim = String.Empty AndAlso IsNumeric(CustomerTypeID) Then iCustomerTypeID = Convert.ToInt32(CustomerTypeID)
       

        Return _MembershipEdit(GUID, Mode, CustomerGroupID, CustomerID, CustomerTypeID)
    End Function

    Public Class MembershipEditClass
        Public Status As CustWebStatus
    End Class

    <WebMethod()> _
    Public Function MembershipEdit_Class(ByVal GUID As String, ByVal Mode As String, ByVal CustomerGroupID As String, ByVal CustomerID As String, ByVal CustomerTypeID As String) As MembershipEditClass
        Dim ds As System.Data.DataSet
        Dim oMembershipEdit As New MembershipEditClass
        Dim oStatus As New CustWebStatus
        Dim dt As DataTable
        'added
        ' Dim iCustomerTypeID As Integer = -1
        'If Not String.IsNullOrEmpty(CustomerTypeID) AndAlso Not CustomerTypeID.Trim = String.Empty AndAlso IsNumeric(CustomerTypeID) Then iCustomerTypeID = Convert.ToInt32(CustomerTypeID)
       

        InitApp()

        ds = _MembershipEdit(GUID, Mode, CustomerGroupID, CustomerID, CustomerTypeID)

        dt = ds.Tables("Status")
        oStatus.StatusCode = dt.Rows(0).Item("StatusCode")
        oStatus.Description = dt.Rows(0).Item("Description")
        oMembershipEdit.Status = oStatus
        Return oMembershipEdit
    End Function
    <WebMethod()> _
    Public Function MembershipEdit_Class_ByCardID(ByVal GUID As String, ByVal Mode As String, ByVal CustomerGroupID As String, ByVal CustomerID As String, ByVal CustomerTypeID As String, ByVal CardTypeID As Integer) As MembershipEditClass
        Dim ds As System.Data.DataSet
        Dim oMembershipEdit As New MembershipEditClass
        Dim oStatus As New CustWebStatus
        Dim dt As DataTable
        'added
        'Dim iCustomerTypeID As Integer = -1
        'If Not String.IsNullOrEmpty(CustomerTypeID) AndAlso Not CustomerTypeID.Trim = String.Empty AndAlso IsNumeric(CustomerTypeID) Then iCustomerTypeID = Convert.ToInt32(CustomerTypeID)
       

        InitApp()

        ds = _MembershipEdit(GUID, Mode, CustomerGroupID, CustomerID, CustomerTypeID, CardTypeID)

        dt = ds.Tables("Status")
        oStatus.StatusCode = dt.Rows(0).Item("StatusCode")
        oStatus.Description = dt.Rows(0).Item("Description")
        oMembershipEdit.Status = oStatus
        Return oMembershipEdit
    End Function
    
  Private Function _MembershipEditByGroupName(ByVal GUID As String, ByVal Mode As String, ByVal CustomerGroupName As String, ByVal CustomerID As String, ByVal CustomerTypeID As Integer, Optional ByVal CardTypeID As Integer = -1) As System.Data.DataSet
    Dim dt As System.Data.DataTable
    Dim dtStatus As System.Data.DataTable
    Dim row As System.Data.DataRow
    Dim ResultSet As New System.Data.DataSet("MembershipEdit")
    Dim CustomerPK As Long = 0
    Dim CustomerGroupID As Long = 0
	
    'Initialize the status table, which will report the success or failure of the operation
    dtStatus = New DataTable
    dtStatus.TableName = "Status"
    dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
    dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

    Try
      If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
      If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

      If Not IsValidGUID(GUID) Then
        'Wrong GUID
        row = dtStatus.NewRow()
        row.Item("StatusCode") = StatusCodes.INVALID_GUID
        row.Item("Description") = "Failure: Invalid GUID."
        dtStatus.Rows.Add(row)
        dtStatus.AcceptChanges()
        ResultSet.Tables.Add(dtStatus.Copy())
      ElseIf (CustomerID.Length < 1) Then
        'Bad customer ID
        row = dtStatus.NewRow()
        row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERID
        row.Item("Description") = "Failure: Invalid customer ID."
        dtStatus.Rows.Add(row)
        dtStatus.AcceptChanges()
        ResultSet.Tables.Add(dtStatus.Copy())
      Else
        'Pad the customer ID
        If CardTypeID <> -1 Then
          CustomerID = MyCommon.Pad_ExtCardID(CustomerID, CardTypeID)
        Else
          CustomerID = MyCommon.Pad_ExtCardID(CustomerID, CardTypes.CUSTOMER)
        End If
        'Find the Customer PK
                MyCommon.QueryStr = "select CID.CustomerPK from CardIDs as CID with (NoLock) " & _
                                    "where ExtCardID='" & MyCryptLib.SQL_StringEncrypt(CustomerID, True) & "' and CardTypeID in " & _
                                    "(select CardTypeID from CardTypes with (NoLock) where CustTypeID=" & CustomerTypeID & ");"
        dt = MyCommon.LXS_Select
        If dt.Rows.Count = 0 Then		  
		  If MyCommon.Fetch_InterfaceOption(11) AndAlso Mode.ToUpper = "OPTIN" Then		    
            Dim NewCustomer As New Copient.Customer
			Dim MyLookup As New Copient.CustomerLookup
			Dim RetCode = StatusCodes.Success
			NewCustomer.AddCard(New Copient.Card(CustomerID, CustomerTypeID))
            MyLookup.AddCustomer(NewCustomer, RetCode)
            If RetCode = StatusCodes.SUCCESS Then
			   CustomerPK = MyLookup.FindCustomerPKFromExtID(CustomerID, CustomerTypeID)
			Else			  
              row = dtStatus.NewRow()
              row.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
              row.Item("Description") = "Could not create a new customer " & CustomerID & " of type " & CustomerTypeID
              dtStatus.Rows.Add(row)
              dtStatus.AcceptChanges()
              ResultSet.Tables.Add(dtStatus.Copy())
			End If   
		  Else
		    'Customer not found
            If CustomerTypeID = 0 Then
              row = dtStatus.NewRow()
              row.Item("StatusCode") = StatusCodes.NOTFOUND_CUSTOMER
              row.Item("Description") = "Failure: Customer " & CustomerID & " not found."
              dtStatus.Rows.Add(row)
              dtStatus.AcceptChanges()
              ResultSet.Tables.Add(dtStatus.Copy())
            ElseIf CustomerTypeID = 1 Then
              row = dtStatus.NewRow()
              row.Item("StatusCode") = StatusCodes.NOTFOUND_HOUSEHOLD
              row.Item("Description") = "Failure: Household " & CustomerID & " not found."
              dtStatus.Rows.Add(row)
              dtStatus.AcceptChanges()
              ResultSet.Tables.Add(dtStatus.Copy())
            ElseIf CustomerTypeID = 2 Then
              row = dtStatus.NewRow()
              row.Item("StatusCode") = StatusCodes.NOTFOUND_CAM
              row.Item("Description") = "Failure: CAM " & CustomerID & " not found."
              dtStatus.Rows.Add(row)
              dtStatus.AcceptChanges()
              ResultSet.Tables.Add(dtStatus.Copy())
            Else
              row = dtStatus.NewRow()
              row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERTYPEID
              row.Item("Description") = "Failure: Invalid customer type."
              dtStatus.Rows.Add(row)
              dtStatus.AcceptChanges()
              ResultSet.Tables.Add(dtStatus.Copy())
		    End If
          End If          
        Else
		  CustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)
		End If
        If CustomerPK > 0 Then  
		  
          'Find the customer group's details
          MyCommon.QueryStr = "select CustomerGroupID from CustomerGroups with (NoLock) where Name='" & CustomerGroupName & "';"
          dt = MyCommon.LRT_Select
          If dt.Rows.Count = 0 Then
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERGROUPNAME
            row.Item("Description") = "Failure: Customer group " & CustomerGroupName & " could not be found."
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
          Else
		    CustomerGroupID = MyCommon.NZ(dt.Rows(0).Item("CustomerGroupID"),0)
            If Mode.ToUpper = "OPTIN" Then			  
              'Execute opt-in procedures:
              MyCommon.QueryStr = "dbo.pt_GroupMembership_Insert_ByPK"
              MyCommon.Open_LXSsp()
              MyCommon.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.NVarChar, 26).Value = CustomerPK
              MyCommon.LXSsp.Parameters.Add("@CustomerGroupID", SqlDbType.BigInt).Value = CustomerGroupID
              MyCommon.LXSsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
              MyCommon.LXSsp.ExecuteNonQuery()
              If MyCommon.LXSsp.Parameters("@Status").Value = 0 Then
                MyCommon.Activity_Log(4, CustomerGroupID, 1, Copient.PhraseLib.Lookup("history.cgroup-optin", 1) & " " & CustomerID)
                MyCommon.QueryStr = "update CustomerGroups with (RowLock) set LastUpdate=getdate() where CustomerGroupID=" & CustomerGroupID & ";"
                MyCommon.LRT_Execute()
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.SUCCESS
                row.Item("Description") = "Success: Customer " & CustomerID & " opted into customer group " & CustomerGroupName & "."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                UpdateCustomerCountViaPK(CustomerPK)
              ElseIf MyCommon.LXSsp.Parameters("@Status").Value = -1 Then
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.FAILED_OPTIN
                row.Item("Description") = "Failure: Unable to opt customer " & CustomerID & " into customer group " & CustomerGroupName & "."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
              End If
            ElseIf Mode.ToUpper = "OPTOUT" Then
              'Execute opt-out procedures:
              MyCommon.QueryStr = "dbo.pt_GroupMembership_Delete_ByPK"
              MyCommon.Open_LXSsp()
              MyCommon.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.NVarChar, 26).Value = CustomerPK
              MyCommon.LXSsp.Parameters.Add("@CustomerGroupID", SqlDbType.BigInt).Value = CustomerGroupID
              MyCommon.LXSsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
              MyCommon.LXSsp.ExecuteNonQuery()
              If MyCommon.LXSsp.Parameters("@Status").Value = 0 Then
                MyCommon.Activity_Log(4, CustomerGroupID, 1, Copient.PhraseLib.Lookup("history.cgroup-optout", 1) & " " & CustomerID)
                MyCommon.QueryStr = "update CustomerGroups with (RowLock) set LastUpdate=getdate() where CustomerGroupID=" & CustomerGroupID & ";"
                MyCommon.LRT_Execute()
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.SUCCESS
                row.Item("Description") = "Success: Customer " & CustomerID & " opted out of customer group " & CustomerGroupName & "."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                UpdateCustomerCountViaPK(CustomerPK)
              ElseIf MyCommon.LXSsp.Parameters("@Status").Value = -1 Then
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.FAILED_OPTOUT
                row.Item("Description") = "Failure: Unable to opt customer " & CustomerID & " out of customer group " & CustomerGroupName & "."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
              End If
            Else
              row = dtStatus.NewRow()
              row.Item("StatusCode") = StatusCodes.INVALID_MODE
              row.Item("Description") = "Failure: Invalid mode.  Specify ""optin"" or ""optout""."
              dtStatus.Rows.Add(row)
              dtStatus.AcceptChanges()
              ResultSet.Tables.Add(dtStatus.Copy())
            End If
          End If
        End If
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
	
    Private Function _MembershipEdit(ByVal GUID As String, ByVal Mode As String, ByVal CustomerGroupID As String, ByVal CustomerID As String, ByVal CustomerTypeID As String, Optional ByVal CardTypeID As Integer = -1) As System.Data.DataSet
        Dim dt As System.Data.DataTable
        Dim dtStatus As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim ResultSet As New System.Data.DataSet("MembershipEdit")
        Dim CustomerPK As Long
        'added
        Dim RetCode As StatusCodes
        Dim RetMsg As String = ""

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))
    
        WriteDebug("_MembershipEdit - CardNumber: " & Copient.MaskHelper.MaskCard(CustomerID, CardTypeID) & ". CardTypeID: " & CardTypeID, DebugState.BeginTime)

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            If Not IsValidGUID(GUID) Then
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                'ElseIf (CustomerID.Length < 1) Then
            ElseIf Not IsValidCustomerCard(CustomerID, CustomerTypeID, RetCode, RetMsg) Then
                
                'Bad customer ID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = RetCode
                row.Item("Description") = RetMsg
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
               
                ''Bad customer ID
                'row = dtStatus.NewRow()
                'row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERID
                'row.Item("Description") = "Failure: Invalid customer ID."
                'dtStatus.Rows.Add(row)
                'dtStatus.AcceptChanges()
                'ResultSet.Tables.Add(dtStatus.Copy())
            Else
                Dim iCustomerTypeID As Integer = -1
                If Not String.IsNullOrEmpty(CustomerTypeID) AndAlso Not CustomerTypeID.Trim = String.Empty AndAlso IsNumeric(CustomerTypeID) Then iCustomerTypeID = Convert.ToInt32(CustomerTypeID)
                
                'Pad the customer ID
                If CardTypeID <> -1 Then
                    CustomerID = MyCommon.Pad_ExtCardID(CustomerID, CardTypeID)
                Else
                    CustomerID = MyCommon.Pad_ExtCardID(CustomerID, CardTypes.CUSTOMER)
                End If
                'Find the Customer PK
                MyCommon.QueryStr = "select CID.CustomerPK from CardIDs as CID with (NoLock) " & _
                                    "INNER JOIN CardTypes AS CT with (NoLock) ON CID.CardTypeID = CT.CardTypeID " & _
                                            "WHERE CT.CustTypeID=" & iCustomerTypeID & " AND CID.ExtCardID='" & MyCryptLib.SQL_StringEncrypt(CustomerID, True) & "';"
                dt = MyCommon.LXS_Select
                If dt.Rows.Count = 0 Then
                    'Customer not found
                    If iCustomerTypeID = 0 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_CUSTOMER
                        row.Item("Description") = "Failure: Customer " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    ElseIf iCustomerTypeID = 1 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_HOUSEHOLD
                        row.Item("Description") = "Failure: Household " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    ElseIf iCustomerTypeID = 2 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_CAM
                        row.Item("Description") = "Failure: CAM " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    Else
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERTYPEID
                        row.Item("Description") = "Failure: Invalid customer type."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    End If
                Else
                    CustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)
                    'Find the customer group's details
                    MyCommon.QueryStr = "select Name from CustomerGroups with (NoLock) where CustomerGroupID=" & CustomerGroupID & ";"
                    dt = MyCommon.LRT_Select
                    If dt.Rows.Count = 0 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERGROUPID
                        row.Item("Description") = "Failure: Customer group " & CustomerGroupID & " could not be found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    Else
                        If Mode = "optin" Then
                            'Execute opt-in procedures:
                            AddHHToGroup = MyCommon.Fetch_SystemOption(138)
                            MyCommon.QueryStr = "dbo.pt_GroupMembership_Insert_ByPK"
                            MyCommon.Open_LXSsp()
                            MyCommon.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.NVarChar, 26).Value = CustomerPK
                            MyCommon.LXSsp.Parameters.Add("@CustomerGroupID", SqlDbType.BigInt).Value = CustomerGroupID
                            MyCommon.LXSsp.Parameters.Add("@AddHHToGroup", SqlDbType.Bit).Value = AddHHToGroup							
                            MyCommon.LXSsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
                            MyCommon.LXSsp.ExecuteNonQuery()
                            If MyCommon.LXSsp.Parameters("@Status").Value = 0 Then
                                MyCommon.Activity_Log(4, CustomerGroupID, 1, Copient.PhraseLib.Lookup("history.cgroup-optin", 1) & " " & CustomerID)
                                MyCommon.QueryStr = "update CustomerGroups with (RowLock) set LastUpdate=getdate() where CustomerGroupID=" & CustomerGroupID & ";"
                                MyCommon.LRT_Execute()
                                row = dtStatus.NewRow()
                                row.Item("StatusCode") = StatusCodes.SUCCESS
                                row.Item("Description") = "Success: Customer " & CustomerID & " opted into customer group " & CustomerGroupID & "."
                                dtStatus.Rows.Add(row)
                                dtStatus.AcceptChanges()
                                ResultSet.Tables.Add(dtStatus.Copy())
                                UpdateCustomerCountViaPK(CustomerPK)
                            ElseIf MyCommon.LXSsp.Parameters("@Status").Value = -1 Then
                                row = dtStatus.NewRow()
                                row.Item("StatusCode") = StatusCodes.FAILED_OPTIN
                                row.Item("Description") = "Failure: Unable to opt customer " & CustomerID & " into customer group " & CustomerGroupID & "."
                                dtStatus.Rows.Add(row)
                                dtStatus.AcceptChanges()
                                ResultSet.Tables.Add(dtStatus.Copy())
                            End If
                        ElseIf Mode = "optout" Then
                            'Execute opt-out procedures:
                            MyCommon.QueryStr = "dbo.pt_GroupMembership_Delete_ByPK"
                            MyCommon.Open_LXSsp()
                            MyCommon.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.NVarChar, 26).Value = CustomerPK
                            MyCommon.LXSsp.Parameters.Add("@CustomerGroupID", SqlDbType.BigInt).Value = CustomerGroupID
                            MyCommon.LXSsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
                            MyCommon.LXSsp.ExecuteNonQuery()
                            If MyCommon.LXSsp.Parameters("@Status").Value = 0 Then
                                MyCommon.Activity_Log(4, CustomerGroupID, 1, Copient.PhraseLib.Lookup("history.cgroup-optout", 1) & " " & CustomerID)
                                MyCommon.QueryStr = "update CustomerGroups with (RowLock) set LastUpdate=getdate() where CustomerGroupID=" & CustomerGroupID & ";"
                                MyCommon.LRT_Execute()
                                row = dtStatus.NewRow()
                                row.Item("StatusCode") = StatusCodes.SUCCESS
                                row.Item("Description") = "Success: Customer " & CustomerID & " opted out of customer group " & CustomerGroupID & "."
                                dtStatus.Rows.Add(row)
                                dtStatus.AcceptChanges()
                                ResultSet.Tables.Add(dtStatus.Copy())
                                UpdateCustomerCountViaPK(CustomerPK)
                            ElseIf MyCommon.LXSsp.Parameters("@Status").Value = -1 Then
                                row = dtStatus.NewRow()
                                row.Item("StatusCode") = StatusCodes.FAILED_OPTOUT
                                row.Item("Description") = "Failure: Unable to opt customer " & CustomerID & " out of customer group " & CustomerGroupID & "."
                                dtStatus.Rows.Add(row)
                                dtStatus.AcceptChanges()
                                ResultSet.Tables.Add(dtStatus.Copy())
                            End If
                        Else
                            row = dtStatus.NewRow()
                            row.Item("StatusCode") = StatusCodes.INVALID_MODE
                            row.Item("Description") = "Failure: Invalid mode.  Specify ""optin"" or ""optout""."
                            dtStatus.Rows.Add(row)
                            dtStatus.AcceptChanges()
                            ResultSet.Tables.Add(dtStatus.Copy())
                        End If
                    End If
                End If
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

        WriteDebug("_MembershipEdit - " & System.Net.Dns.GetHostName(), DebugState.EndTime)
        WriteLogLines()
    
        Return ResultSet
    End Function
  
    Private Sub UpdateCustomerCountViaPK(ByVal lCustomerPK As Long)
        Dim iStatus As Integer = 0
        Dim CmInstalled As Boolean = MyCommon.IsEngineInstalled(0)

        If CmInstalled Then
            Dim dt As DataTable

            ' check customer is member of household
            MyCommon.QueryStr = "select isnull(HHPK, 0) as HHPK from Customers with (NoLock) where CustomerPK=" & lCustomerPK & ";"
            dt = MyCommon.LXS_Select
            If dt.Rows.Count > 0 Then
                Dim lHHPK As Long = 0

                If Not Long.TryParse(MyCommon.NZ(dt.Rows(0).Item(0), "0"), lHHPK) Then lHHPK = 0
                If lHHPK > 0 Then
                    ' set to household PK
                    lCustomerPK = lHHPK
                End If
                MyCommon.QueryStr = "update Customers with (RowLock) set UpdateCount=UpdateCount+1 where CustomerPK=" & lCustomerPK & ";"
                MyCommon.LXS_Execute()
            End If
        End If
    End Sub
  
    <WebMethod()> _
    Public Function OfferList(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As String) As System.Data.DataSet
        Dim dt As System.Data.DataTable
        Dim dtOffers As System.Data.DataTable
        Dim dtGroups As System.Data.DataTable
        Dim dtStatus As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim CustomerPK As Long
        Dim ResultSet As New System.Data.DataSet("OfferList")
        WriteDebug("OfferList - CardNumber: " & Copient.MaskHelper.MaskCard(CustomerID, CardTypes.CUSTOMER) & ". CardTypeID: " & CardTypes.CUSTOMER, DebugState.BeginTime)
        'added
        Dim RetCode As StatusCodes
        Dim RetMsg As String = ""

        InitApp()

        'Initialize the status table, which will report the success or failure of the CustomerDetails operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            If Not IsValidGUID(GUID) Then
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                'ElseIf (CustomerID.Length < 1) Then
            ElseIf Not IsValidCustomerCard(CustomerID, CustomerTypeID, RetCode, RetMsg) Then
                
                'Bad customer ID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = RetCode
                row.Item("Description") = RetMsg
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                
            Else
                Dim iCustomerTypeID As Integer = -1
                If Not String.IsNullOrEmpty(CustomerTypeID) AndAlso Not CustomerTypeID.Trim = String.Empty AndAlso IsNumeric(CustomerTypeID) Then iCustomerTypeID = Convert.ToInt32(CustomerTypeID)
                'Pad the customer ID
                CustomerID = MyCommon.Pad_ExtCardID(CustomerID, CardTypes.CUSTOMER)
                'Find the Customer PK
                MyCommon.QueryStr = "select CID.CustomerPK from CardIDs as CID with (NoLock) " & _
                                  "INNER JOIN CardTypes AS CT with (NoLock) ON CID.CardTypeID = CT.CardTypeID " & _
                                  "WHERE CT.CustTypeID=" & iCustomerTypeID & " AND CID.ExtCardID='" & MyCryptLib.SQL_StringEncrypt(CustomerID, True) & "';"
                dt = MyCommon.LXS_Select
                If dt.Rows.Count = 0 Then
                    'Customer not found
                    If iCustomerTypeID = 0 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_CUSTOMER
                        row.Item("Description") = "Failure: Customer " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    ElseIf iCustomerTypeID = 1 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_HOUSEHOLD
                        row.Item("Description") = "Failure: Household " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    ElseIf iCustomerTypeID = 2 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_CAM
                        row.Item("Description") = "Failure: CAM " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    Else
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERTYPEID
                        row.Item("Description") = "Failure: Invalid customer type."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    End If
                Else
                    CustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)

                    'Put a success into the Status table
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = 0
                    row.Item("Description") = "Success."
                    dtStatus.Rows.Add(row)

                    'Get the lists of offers
                    WriteDebug("Entering Send_XMLCurrentOffers", DebugState.CurrentTime)
                    dtOffers = Send_XMLCurrentOffers(CustomerPK)
                    WriteDebug("Exiting Send_XMLCurrentOffers", DebugState.CurrentTime)

                    WriteDebug("Entering Send_XMLGroupOffers", DebugState.CurrentTime)
                    dtGroups = Send_XMLGroupOffers(CustomerPK)
                    WriteDebug("Exiting Send_XMLGroupOffers", DebugState.CurrentTime)

                    dtOffers.TableName = "Offers"
                    dtGroups.TableName = "Groups"

                    'Populate and return the DataSet
                    ResultSet.Tables.Add(dtStatus.Copy())
                    ResultSet.Tables.Add(dtOffers.Copy())
                    ResultSet.Tables.Add(dtGroups.Copy())

                End If
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
        WriteDebug("OfferList - " & System.Net.Dns.GetHostName(), DebugState.EndTime)
        WriteLogLines()

        Return ResultSet
    End Function
    
    <WebMethod()> _
    Public Function OfferList_ByCardID(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As String, ByVal CardTypeID As Integer) As System.Data.DataSet
        Dim dt As System.Data.DataTable
        Dim dtOffers As System.Data.DataTable
        Dim dtGroups As System.Data.DataTable
        Dim dtStatus As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim CustomerPK As Long
        Dim ResultSet As New System.Data.DataSet("OfferList")
        'added
        Dim RetCode As StatusCodes
        Dim RetMsg As String = ""
        

        InitApp()

        'Initialize the status table, which will report the success or failure of the CustomerDetails operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            If Not IsValidGUID(GUID) Then
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                'ElseIf (CustomerID.Length < 1) Then
            ElseIf Not IsValidCustomerCard(CustomerID, CustomerTypeID, RetCode, RetMsg) Then
                
                'Bad customer ID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = RetCode
                row.Item("Description") = RetMsg
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
            
            Else
                Dim iCustomerTypeID As Integer = -1
                If Not String.IsNullOrEmpty(CustomerTypeID) AndAlso Not CustomerTypeID.Trim = String.Empty AndAlso IsNumeric(CustomerTypeID) Then iCustomerTypeID = Convert.ToInt32(CustomerTypeID)
               
                'Pad the customer ID
                CustomerID = MyCommon.Pad_ExtCardID(CustomerID, CardTypeID)

                'Find the Customer PK
                MyCommon.QueryStr = "select CID.CustomerPK from CardIDs as CID with (NoLock) " & _
                                    "where ExtCardID='" & MyCryptLib.SQL_StringEncrypt(CustomerID, True) & "' and CardTypeID = " & CardTypeID & ";"
                dt = MyCommon.LXS_Select
                If dt.Rows.Count = 0 Then
                    'Customer not found
                    If iCustomerTypeID = 0 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_CUSTOMER
                        row.Item("Description") = "Failure: Customer " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    ElseIf iCustomerTypeID = 1 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_HOUSEHOLD
                        row.Item("Description") = "Failure: Household " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    ElseIf iCustomerTypeID = 2 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_CAM
                        row.Item("Description") = "Failure: CAM " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    Else
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERTYPEID
                        row.Item("Description") = "Failure: Invalid customer type."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    End If
                Else
                    CustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)

                    'Put a success into the Status table
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = 0
                    row.Item("Description") = "Success."
                    dtStatus.Rows.Add(row)

                    'Get the lists of offers
                    dtOffers = Send_XMLCurrentOffers(CustomerPK)
                    dtGroups = Send_XMLGroupOffers(CustomerPK)

                    dtOffers.TableName = "Offers"
                    dtGroups.TableName = "Groups"

                    'Populate and return the DataSet
                    ResultSet.Tables.Add(dtStatus.Copy())
                    ResultSet.Tables.Add(dtOffers.Copy())
                    ResultSet.Tables.Add(dtGroups.Copy())

                End If
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

    'AMSPS-2320 : CR 118 - OfferListCMAdvanced method is developed for specific client.
    'This method will return all active, external offers which have “FOB Eligible” advanced option on offer-gen.aspx set to 1
   <WebMethod()> _
   Public Function OfferListCMAdvanced(ByVal GUID As String) As System.Data.DataSet
        InitApp()
        Return _OfferListCMAdvanced(GUID)
   End Function

    <WebMethod()> _
    Public Function OfferListCM(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As Integer) As System.Data.DataSet
        InitApp()
        'Dim iCustomerTypeID As Integer = -1
        'If Not String.IsNullOrEmpty(CustomerTypeID) AndAlso Not CustomerTypeID.Trim = String.Empty AndAlso IsNumeric(CustomerTypeID) Then iCustomerTypeID = Convert.ToInt32(CustomerTypeID)
       
        Return _OfferListCM(GUID, CustomerID, CustomerTypeID)
    End Function
    
    <WebMethod()> _
    Public Function OfferListCM_ByCardID(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As Integer, ByVal CardTypeID As Integer) As System.Data.DataSet
        InitApp()
        'Dim iCustomerTypeID As Integer = -1
        'If Not String.IsNullOrEmpty(CustomerTypeID) AndAlso Not CustomerTypeID.Trim = String.Empty AndAlso IsNumeric(CustomerTypeID) Then iCustomerTypeID = Convert.ToInt32(CustomerTypeID)
       
        Return _OfferListCM(GUID, CustomerID, CustomerTypeID, CardTypeID)
    End Function

    Public Class Offer
        Public OfferID As Long = 0
        Public Name As String = ""
        Public Description As String = ""
        Public OfferCategory As String = ""
        Public StartDate As String = ""
        Public EndDate As String = ""
        Public DisplayStartDate As String = Nothing
        Public DisplayEndDate As String = Nothing
        Public CustomerGroupID As Long = 0
        Public EmployeesOnly As String = "false"
        Public EmployeesExcluded As String = "false"
    End Class

    Public Class Program
        Public OfferID As String = ""
        Public ConditionOrReward As String = ""
        Public ProgramType As String = ""
        Public ProgramID As String = ""
        Public TierLevel As Integer = 0
        Public Amount As Decimal = 0.0
        Public Name As String = ""
    End Class

    Public Class OfferListCMClass
        Public Status As CustWebStatus
        Public Offers() As Offer
        Public Programs() As Program
    End Class

    <WebMethod()> _
    Public Function OfferListCM_Class(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As Integer) As OfferListCMClass
        Dim ds As System.Data.DataSet
        Dim oOfferList As New OfferListCMClass
        Dim oStatus As New CustWebStatus
        Dim dt As DataTable
        Dim dr As DataRow
        Dim i As Integer

        InitApp()

        ds = _OfferListCM(GUID, CustomerID, CustomerTypeID)

        dt = ds.Tables("Status")
        oStatus.StatusCode = dt.Rows(0).Item("StatusCode")
        oStatus.Description = dt.Rows(0).Item("Description")
        oOfferList.Status = oStatus
        If oStatus.StatusCode = "0" Then
            dt = ds.Tables("Offers")
            If dt.Rows.Count > 0 Then
                ReDim oOfferList.Offers(dt.Rows.Count - 1)
                Dim oOffer As Offer
                i = 0
                For Each dr In dt.Rows
                    oOffer = New Offer
                    oOffer.OfferID = dr.Item("OfferID")
                    oOffer.Name = dr.Item("Name")
                    oOffer.Description = dr.Item("Description")
                    oOffer.OfferCategory = dr.Item("OfferCategory")
                    oOffer.StartDate = dr.Item("StartDate")
                    oOffer.EndDate = dr.Item("EndDate")
                    If (MyCommon.Fetch_CM_SystemOption(85) = "1") Then
                        If Not IsDBNull(dr.Item("DisplayStartDate")) Then
                            oOffer.DisplayStartDate = dr.Item("DisplayStartDate")
                        Else
                            oOffer.DisplayStartDate = ""
                        End If
                        If Not IsDBNull(dr.Item("DisplayEndDate")) Then
                            oOffer.DisplayEndDate = dr.Item("DisplayEndDate")
                        Else
                            oOffer.DisplayEndDate = ""
                        End If
                    End If
                    oOffer.CustomerGroupID = dr.Item("CustomerGroupID")
                    oOffer.EmployeesOnly = dr.Item("EmployeesOnly").ToString.ToLower
                    oOffer.EmployeesExcluded = dr.Item("EmployeesExcluded").ToString.ToLower
                    oOfferList.Offers(i) = oOffer
                    i += 1
                Next
            End If
            dt = ds.Tables("Programs")
            If dt.Rows.Count > 0 Then
                ReDim oOfferList.Programs(dt.Rows.Count - 1)
                Dim p As Program
                i = 0
                For Each dr In dt.Rows
                    p = New Program
                    p.OfferID = dr.Item("OfferID")
                    p.ConditionOrReward = dr.Item("ConditionOrReward")
                    p.ProgramType = dr.Item("ProgramType")
                    p.ProgramID = dr.Item("ProgramID")
                    p.TierLevel = dr.Item("TierLevel")
                    p.Amount = dr.Item("Amount")
                    p.Name = dr.Item("Name")
                    oOfferList.Programs(i) = p
                    i += 1
                Next
            End If
        End If

        Return oOfferList

    End Function

    <WebMethod()> _
    Public Function ExpiredExternalOfferListCM_Class(ByVal GUID As String, ByVal ExtCardID As String, ByVal CardTypeID As Integer) As OfferListCMClass
        Dim ds As System.Data.DataSet
        Dim oOfferList As New OfferListCMClass
        Dim oStatus As New CustWebStatus
        Dim dt As DataTable
        Dim dr As DataRow
        Dim i As Integer
        'added
        'Dim iCardTypeID As Integer = -1
        'If Not String.IsNullOrEmpty(CardTypeID) AndAlso Not CardTypeID.Trim = String.Empty AndAlso IsNumeric(CardTypeID) Then iCardTypeID = Convert.ToInt32(CardTypeID)
    

        InitApp()

        ds = GetCustomerExternalExpiredCmOffers(GUID, ExtCardID, CardTypeID)

        dt = ds.Tables("Status")
        oStatus.StatusCode = dt.Rows(0).Item("StatusCode")
        oStatus.Description = dt.Rows(0).Item("Description")
        oOfferList.Status = oStatus
        If oStatus.StatusCode = "0" Then
            dt = ds.Tables("Offers")
            If dt.Rows.Count > 0 Then
                ReDim oOfferList.Offers(dt.Rows.Count - 1)
                Dim oOffer As Offer
                i = 0
                For Each dr In dt.Rows
                    oOffer = New Offer
                    oOffer.OfferID = dr.Item("OfferID")
                    oOffer.Name = dr.Item("Name")
                    oOffer.Description = dr.Item("Description")
                    oOffer.OfferCategory = dr.Item("OfferCategory")
                    oOffer.StartDate = dr.Item("StartDate")
                    oOffer.EndDate = dr.Item("EndDate")
                    oOffer.CustomerGroupID = dr.Item("CustomerGroupID")
                    oOffer.EmployeesOnly = dr.Item("EmployeesOnly").ToString.ToLower
                    oOffer.EmployeesExcluded = dr.Item("EmployeesExcluded").ToString.ToLower
                    oOfferList.Offers(i) = oOffer
                    i += 1
                Next
            End If
            dt = ds.Tables("Programs")
            If dt.Rows.Count > 0 Then
                ReDim oOfferList.Programs(dt.Rows.Count - 1)
                Dim p As Program
                i = 0
                For Each dr In dt.Rows
                    p = New Program
                    p.OfferID = dr.Item("OfferID")
                    p.ConditionOrReward = dr.Item("ConditionOrReward")
                    p.ProgramType = dr.Item("ProgramType")
                    p.ProgramID = dr.Item("ProgramID")
                    p.TierLevel = dr.Item("TierLevel")
                    p.Amount = dr.Item("Amount")
                    p.Name = dr.Item("Name")
                    oOfferList.Programs(i) = p
                    i += 1
                Next
            End If
        End If

        Return oOfferList

    End Function


    <WebMethod()> _
    Public Function ExternalOfferListCM_Class(ByVal GUID As String, ByVal ExtCardID As String, ByVal CardTypeID As Integer, ByVal IncludeRedeemed As Boolean, _
                                           ByVal IncludeUnRedeemed As Boolean, ByVal IncludeExpired As Boolean) As OfferListCMClass
        Dim ds As System.Data.DataSet
        Dim oOfferList As New OfferListCMClass
        Dim oStatus As New CustWebStatus
        Dim dt As DataTable
        Dim dr As DataRow
        Dim i As Integer
        'added
        'Dim iCardTypeID As Integer = -1
        'If Not String.IsNullOrEmpty(CardTypeID) AndAlso Not CardTypeID.Trim = String.Empty AndAlso IsNumeric(CardTypeID) Then iCardTypeID = Convert.ToInt32(CardTypeID)
        Try
            InitApp()

            If IncludeExpired = False Then
                ' _ExternalOfferListCM only returns offers that have not expired
                ds = _ExternalOfferListCM(GUID, ExtCardID, CardTypeID, IncludeRedeemed, IncludeUnRedeemed)

                dt = ds.Tables("Status")
                oStatus.StatusCode = dt.Rows(0).Item("StatusCode")
                oStatus.Description = dt.Rows(0).Item("Description")
                oOfferList.Status = oStatus
                If oStatus.StatusCode = "0" Then
                    dt = ds.Tables("Offers")
                    If dt.Rows.Count > 0 Then
                        ReDim oOfferList.Offers(dt.Rows.Count - 1)
                        Dim oOffer As Offer
                        i = 0
                        For Each dr In dt.Rows
                            oOffer = New Offer
                            oOffer.OfferID = dr.Item("OfferID")
                            oOffer.Name = dr.Item("Name")
                            oOffer.Description = dr.Item("Description")
                            oOffer.OfferCategory = dr.Item("OfferCategory")
                            oOffer.StartDate = dr.Item("StartDate")
                            oOffer.EndDate = dr.Item("EndDate")
                            oOffer.CustomerGroupID = dr.Item("CustomerGroupID")
                            oOffer.EmployeesOnly = dr.Item("EmployeesOnly").ToString.ToLower
                            oOffer.EmployeesExcluded = dr.Item("EmployeesExcluded").ToString.ToLower
                            oOfferList.Offers(i) = oOffer
                            i += 1
                        Next
                    End If
                    dt = ds.Tables("Programs")
                    If dt.Rows.Count > 0 Then
                        ReDim oOfferList.Programs(dt.Rows.Count - 1)
                        Dim p As Program
                        i = 0
                        For Each dr In dt.Rows
                            p = New Program
                            p.OfferID = dr.Item("OfferID")
                            p.ConditionOrReward = dr.Item("ConditionOrReward")
                            p.ProgramType = dr.Item("ProgramType")
                            p.ProgramID = dr.Item("ProgramID")
                            p.TierLevel = dr.Item("TierLevel")
                            p.Amount = dr.Item("Amount")
                            p.Name = dr.Item("Name")
                            oOfferList.Programs(i) = p
                            i += 1
                        Next
                    End If
                End If
            ElseIf (IncludeExpired = True) And (IncludeRedeemed = True) And (IncludeUnRedeemed = True) Then
                'Expired offer list- contains both redeemed and unredeemed offers
                oOfferList = ExpiredExternalOfferListCM_Class(GUID, ExtCardID, CardTypeID)
            End If

        Catch ex As Exception

            oStatus.StatusCode = StatusCodes.APPLICATION_EXCEPTION
            oStatus.Description = "Failure: Application " & ex.ToString
            oOfferList.Status = oStatus
        Finally
     
        End Try
        Return oOfferList

    End Function


    Private Function _OfferListCM(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As Integer, Optional ByVal CardTypeID As Integer = -1) As System.Data.DataSet
        Dim dt As System.Data.DataTable
        Dim dtOffers As System.Data.DataTable
        Dim dtPrograms As System.Data.DataTable
        Dim dtStatus As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim lCustomerPK As Long
        Dim lHouseholdPK As Long
        Dim bStatus As Boolean
        Dim ResultSet As New System.Data.DataSet("OfferListCM")
        'added
        Dim RetCode As StatusCodes
        Dim RetMsg As String = ""

        WriteDebug("OfferListCM - CardNumber: " & Copient.MaskHelper.MaskCard(CustomerID, CardTypeID) & ". CardTypeID: " & CardTypeID, DebugState.BeginTime)
        'Initialize the status table, which will report the success or failure of the CustomerDetails operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            If Not IsValidGUID(GUID) Then
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                'ElseIf (CustomerID.Length < 1) Then
            ElseIf Not IsValidCustomerCard(CustomerID, CustomerTypeID.ToString(), RetCode, RetMsg) Then
                
                'Bad customer ID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = RetCode
                row.Item("Description") = RetMsg
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                
            Else
                If CardTypeID <> -1 Then
                    'Pad the customer ID
                    CustomerID = MyCommon.Pad_ExtCardID(CustomerID, CardTypeID)
                Else
                    'Pad the customer ID
                    CustomerID = MyCommon.Pad_ExtCardID(CustomerID, CardTypes.CUSTOMER)
                End If
                'Find the Customer PK
                MyCommon.QueryStr = "select CID.CustomerPK, C.HHPK from CardIDs as CID with (NoLock) " & _
                                    "INNER JOIN Customers as C with (NoLock) on C.CustomerPK=CID.CustomerPK " & _
                                    "INNER JOIN CardTypes AS CT with (NoLock) ON CID.CardTypeID = CT.CardTypeID " & _
                                    "WHERE CT.CustTypeID=" & CustomerTypeID & " AND CID.ExtCardID='" & MyCryptLib.SQL_StringEncrypt(CustomerID, True) & "';"
                dt = MyCommon.LXS_Select
                If dt.Rows.Count = 0 Then
                    'Customer not found
                    If CustomerTypeID = 0 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_CUSTOMER
                        row.Item("Description") = "Failure: Customer " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    ElseIf CustomerTypeID = 1 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_HOUSEHOLD
                        row.Item("Description") = "Failure: Household " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    ElseIf CustomerTypeID = 2 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_CAM
                        row.Item("Description") = "Failure: CAM " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    Else
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERTYPEID
                        row.Item("Description") = "Failure: Invalid customer type."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    End If
                Else
                    lCustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)
                    lHouseholdPK = MyCommon.NZ(dt.Rows(0).Item("HHPK"), 0)

                    'Put a success into the Status table
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = 0
                    row.Item("Description") = "Success."
                    dtStatus.Rows.Add(row)


                    dtOffers = New DataTable
                    dtOffers.TableName = "Offers"
                    dtOffers.Columns.Add("OfferID", System.Type.GetType("System.Int64"))
                    dtOffers.Columns.Add("Name", System.Type.GetType("System.String"))
                    dtOffers.Columns.Add("Description", System.Type.GetType("System.String"))
                    dtOffers.Columns.Add("OfferCategory", System.Type.GetType("System.String"))
                    dtOffers.Columns.Add("StartDate", System.Type.GetType("System.DateTime"))
                    dtOffers.Columns.Add("EndDate", System.Type.GetType("System.DateTime"))
                    If (MyCommon.Fetch_CM_SystemOption(85) = "1") Then
                        dtOffers.Columns.Add("DisplayStartDate", System.Type.GetType("System.DateTime"))
                        dtOffers.Columns.Add("DisplayEndDate", System.Type.GetType("System.DateTime"))
                    End If
                    dtOffers.Columns.Add("CustomerGroupID", System.Type.GetType("System.Int64"))
                    dtOffers.Columns.Add("EmployeesOnly", System.Type.GetType("System.Boolean"))
                    dtOffers.Columns.Add("EmployeesExcluded", System.Type.GetType("System.Boolean"))

                    dtPrograms = New DataTable
                    dtPrograms.TableName = "Programs"
                    dtPrograms.Columns.Add("OfferID", System.Type.GetType("System.Int64"))
                    dtPrograms.Columns.Add("ConditionOrReward", System.Type.GetType("System.String"))
                    dtPrograms.Columns.Add("ProgramType", System.Type.GetType("System.String"))
                    dtPrograms.Columns.Add("ProgramID", System.Type.GetType("System.Int64"))
                    dtPrograms.Columns.Add("TierLevel", System.Type.GetType("System.Int32"))
                    dtPrograms.Columns.Add("Amount", System.Type.GetType("System.Decimal"))
                    dtPrograms.Columns.Add("Name", System.Type.GetType("System.String"))
                    dtPrograms.Columns.Add("Description", System.Type.GetType("System.String"))
                    dtPrograms.Columns.Add("Category", System.Type.GetType("System.String"))
                    dtPrograms.Columns.Add("WebMessage", System.Type.GetType("System.String"))

                    WriteDebug("Entering GetCustomerCmOffers", DebugState.CurrentTime)
                    bStatus = GetCustomerCmOffers(lCustomerPK, lHouseholdPK, dtOffers, dtPrograms)
                    WriteDebug("Exiting GetCustomerCmOffers", DebugState.CurrentTime)


                    'Populate and return the DataSet
                    ResultSet.Tables.Add(dtStatus.Copy())
                    ResultSet.Tables.Add(dtOffers.Copy())
                    ResultSet.Tables.Add(dtPrograms.Copy())

                End If
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

        WriteDebug("OfferListCM - " & System.Net.Dns.GetHostName(), DebugState.EndTime)
        WriteLogLines()
    
        Return ResultSet
    End Function

    'AMSPS-2320 : CR 118
    Private Function _OfferListCMAdvanced(ByVal GUID As String) As System.Data.DataSet
        
        Dim dtOffers As System.Data.DataTable
        Dim bFobEligibilityEnabled As Boolean = IIf(MyCommon.Fetch_CM_SystemOption(142) = "1", True, False)
        Dim dtStatus As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim ResultSet As New System.Data.DataSet("OfferListCMAdvanced")
        
        Dim RetMsg As String = ""

        WriteDebug("OfferListCMAdvanced", DebugState.BeginTime)
        
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            If Not IsValidGUID(GUID) Then
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
            Else
                If (bFobEligibilityEnabled) Then
                
                    'Put a success into the Status table
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = 0
                    row.Item("Description") = "Success."
                    dtStatus.Rows.Add(row)

                                
                    MyCommon.QueryStr = "pa_custwebConnector_getExternalFobEligibleOffers"
                    MyCommon.Open_LRTsp()
                    
                    WriteDebug("_OfferListCMAdvanced - Starting Query For Offers [LogixRT]", DebugState.CurrentTime)
                    dtOffers = MyCommon.LRTsp_select
                    WriteDebug("_OfferListCMAdvanced - Completed Query For Offers with row count=" & dtOffers.Rows.Count, DebugState.CurrentTime)


                    'Populate and return the DataSet
                    ResultSet.Tables.Add(dtStatus.Copy())
                    ResultSet.Tables.Add(dtOffers.Copy())
                
                    ResultSet.Tables(1).TableName = "Offers"
                Else
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = StatusCodes.NOTFOUND_RECORDS
                    row.Item("Description") = "Failure: Cannot return records as FOB eligible system option is not enabled. "
                    dtStatus.Rows.Add(row)
                    dtStatus.AcceptChanges()
                    ResultSet.Tables.Add(dtStatus.Copy())
                End If
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

        WriteDebug("OfferListCMAdvanced - " & System.Net.Dns.GetHostName(), DebugState.EndTime)
        WriteLogLines()
    
        Return ResultSet
    End Function

    'This method only returns those offers where the offer limits/advance limits have not been reached.

    <WebMethod()> _
    Public Function OfferListLimitsBasedCM(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As Integer) As System.Data.DataSet
        Dim dt As System.Data.DataTable
        Dim dtOffers As System.Data.DataTable
        Dim dtStatus As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim lCustomerPK As Long
        Dim lHouseholdPK As Long
        Dim bStatus As Boolean
        Dim bGetHHAndMemberOffers As Boolean = False
        'added
        
        Dim RetCode As StatusCodes
        Dim RetMsg As String = ""
        'Dim iCustomerTypeID As Integer = -1
        'If Not String.IsNullOrEmpty(CustomerTypeID) AndAlso Not CustomerTypeID.Trim = String.Empty AndAlso IsNumeric(CustomerTypeID) Then iCustomerTypeID = Convert.ToInt32(CustomerTypeID)   
       

        Dim ResultSet As New System.Data.DataSet("OfferListLimitsBasedCM")
        WriteDebug("OfferListLimitsBasedCM - CardNumber: " & Copient.MaskHelper.MaskCard(CustomerID, CardTypes.CUSTOMER) & ". CardTypeID:" & CardTypes.CUSTOMER, DebugState.BeginTime)

        InitApp()

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            If Not IsValidGUID(GUID) Then
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                'ElseIf (CustomerID.Length < 1) Then
            ElseIf Not IsValidCustomerCard(CustomerID, CustomerTypeID.ToString(), RetCode, RetMsg) Then
                
                'Bad customer ID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = RetCode
                row.Item("Description") = RetMsg
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
              
            Else
                'Pad the customer ID
                CustomerID = MyCommon.Pad_ExtCardID(CustomerID, CardTypes.CUSTOMER)
                'Find the Customer PK
                MyCommon.QueryStr = "select CID.CustomerPK, C.HHPK from CardIDs as CID with (NoLock) " & _
                                    "inner join Customers as C with (NoLock) on C.CustomerPK=CID.CustomerPK " & _
                                    "INNER JOIN CardTypes AS CT with (NoLock) ON CID.CardTypeID = CT.CardTypeID " & _
                                    "WHERE CT.CustTypeID=" & CustomerTypeID & " AND CID.ExtCardID='" & MyCryptLib.SQL_StringEncrypt(CustomerID, True) & "';"
                dt = MyCommon.LXS_Select
                If dt.Rows.Count = 0 Then
                    'Customer not found
                    If CustomerTypeID = 0 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_CUSTOMER
                        row.Item("Description") = "Failure: Customer " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    ElseIf CustomerTypeID = 1 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_HOUSEHOLD
                        row.Item("Description") = "Failure: Household " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    ElseIf CustomerTypeID = 2 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_CAM
                        row.Item("Description") = "Failure: CAM " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    Else
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERTYPEID
                        row.Item("Description") = "Failure: Invalid customer type."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    End If
                Else
                    lCustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)
                    lHouseholdPK = MyCommon.NZ(dt.Rows(0).Item("HHPK"), 0)

                    'Put a success into the Status table
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = 0
                    row.Item("Description") = "Success."
                    dtStatus.Rows.Add(row)


                    dtOffers = New DataTable
                    dtOffers.TableName = "Offers"
                    dtOffers.Columns.Add("OfferID", System.Type.GetType("System.Int64"))
                    dtOffers.Columns.Add("Name", System.Type.GetType("System.String"))
                    dtOffers.Columns.Add("Description", System.Type.GetType("System.String"))
                    dtOffers.Columns.Add("OfferCategory", System.Type.GetType("System.String"))
                    dtOffers.Columns.Add("StartDate", System.Type.GetType("System.DateTime"))
                    dtOffers.Columns.Add("EndDate", System.Type.GetType("System.DateTime"))
                    If (MyCommon.Fetch_CM_SystemOption(85) = "1") Then
                        dtOffers.Columns.Add("DisplayStartDate", System.Type.GetType("System.DateTime"))
                        dtOffers.Columns.Add("DisplayEndDate", System.Type.GetType("System.DateTime"))
                    End If
                    dtOffers.Columns.Add("CustomerGroupID", System.Type.GetType("System.Int64"))
                    dtOffers.Columns.Add("EmployeesOnly", System.Type.GetType("System.Boolean"))
                    dtOffers.Columns.Add("EmployeesExcluded", System.Type.GetType("System.Boolean"))
          
          
                    If ((MyCommon.Fetch_CM_SystemOption(24) = "1") And (CustomerTypeID = 1)) Then
                        bGetHHAndMemberOffers = True
                        lHouseholdPK = lCustomerPK ' Both parameters should be the customerpk of the household record, as the logic of the
                        ' SP pa_LogixServ_FetchCustGroups_MemberOrHousehold that gets called
                        ' later is based on HHPK<>0 
                    End If
                    WriteDebug("Entering GetCustomerCmOffersLimitsBased", DebugState.CurrentTime)
                    bStatus = GetCustomerCmOffersLimitsBased(lCustomerPK, lHouseholdPK, dtOffers, bGetHHAndMemberOffers)
                    WriteDebug("Exiting GetCustomerCmOffersLimitsBased", DebugState.CurrentTime)


                    'Populate and return the DataSet
                    ResultSet.Tables.Add(dtStatus.Copy())
                    ResultSet.Tables.Add(dtOffers.Copy())


                End If
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
        WriteDebug("OfferListLimitsBasedCM - " & System.Net.Dns.GetHostName(), DebugState.EndTime)
        WriteLogLines()

        Return ResultSet
    End Function

    <WebMethod()> _
    Public Function OfferListLimitsBasedCM_ByCardID(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As Integer, ByVal CardTypeID As Integer) As System.Data.DataSet
        Dim dt As System.Data.DataTable
        Dim dtOffers As System.Data.DataTable
        Dim dtStatus As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim lCustomerPK As Long
        Dim lHouseholdPK As Long
        Dim bStatus As Boolean
        Dim bGetHHAndMemberOffers As Boolean = False

        'added
        Dim RetCode As StatusCodes
        Dim RetMsg As String = ""
        'Dim iCustomerTypeID As Integer = -1
        'If Not String.IsNullOrEmpty(CustomerTypeID) AndAlso Not CustomerTypeID.Trim = String.Empty AndAlso IsNumeric(CustomerTypeID) Then iCustomerTypeID = Convert.ToInt32(CustomerTypeID)
       
        ''added
        'Dim iCardTypeID As Integer = -1
        'If Not String.IsNullOrEmpty(CardTypeID) AndAlso Not CardTypeID.Trim = String.Empty AndAlso IsNumeric(CardTypeID) Then iCardTypeID = Convert.ToInt32(CardTypeID)   
       
        
        

        Dim ResultSet As New System.Data.DataSet("OfferListLimitsBasedCM")
        WriteDebug("OfferListLimitsBasedCM_ByCardID - CardNumber: " & Copient.MaskHelper.MaskCard(CustomerID, CardTypeID) & ". CardTypeID: " & CardTypeID, DebugState.BeginTime)

        InitApp()

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            If Not IsValidGUID(GUID) Then
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                'ElseIf (CustomerID.Length < 1) Then
            ElseIf Not IsValidCustomerCard(CustomerID, CustomerTypeID.ToString(), RetCode, RetMsg) Then
                
                'Bad customer ID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = RetCode
                row.Item("Description") = RetMsg
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                
            Else
                'Pad the customer ID
                CustomerID = MyCommon.Pad_ExtCardID(CustomerID, CardTypeID)

                'Find the Customer PK
                MyCommon.QueryStr = "select CID.CustomerPK, C.HHPK from CardIDs as CID with (NoLock) " & _
                                    "inner join Customers as C with (NoLock) on C.CustomerPK=CID.CustomerPK " & _
                                                    "where CID.ExtCardID='" & MyCryptLib.SQL_StringEncrypt(CustomerID, True) & "' and CID.CardTypeID = " & CardTypeID & ";"
                dt = MyCommon.LXS_Select
                If dt.Rows.Count = 0 Then
                    'Customer not found
                    If CustomerTypeID = 0 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_CUSTOMER
                        row.Item("Description") = "Failure: Customer " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    ElseIf CustomerTypeID = 1 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_HOUSEHOLD
                        row.Item("Description") = "Failure: Household " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    ElseIf CustomerTypeID = 2 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_CAM
                        row.Item("Description") = "Failure: CAM " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    Else
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERTYPEID
                        row.Item("Description") = "Failure: Invalid customer type."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    End If
                Else
                    lCustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)
                    lHouseholdPK = MyCommon.NZ(dt.Rows(0).Item("HHPK"), 0)

                    'Put a success into the Status table
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = 0
                    row.Item("Description") = "Success."
                    dtStatus.Rows.Add(row)


                    dtOffers = New DataTable
                    dtOffers.TableName = "Offers"
                    dtOffers.Columns.Add("OfferID", System.Type.GetType("System.Int64"))
                    dtOffers.Columns.Add("Name", System.Type.GetType("System.String"))
                    dtOffers.Columns.Add("Description", System.Type.GetType("System.String"))
                    dtOffers.Columns.Add("OfferCategory", System.Type.GetType("System.String"))
                    dtOffers.Columns.Add("StartDate", System.Type.GetType("System.DateTime"))
                    dtOffers.Columns.Add("EndDate", System.Type.GetType("System.DateTime"))
                    If (MyCommon.Fetch_CM_SystemOption(85) = "1") Then
                        dtOffers.Columns.Add("DisplayStartDate", System.Type.GetType("System.DateTime"))
                        dtOffers.Columns.Add("DisplayEndDate", System.Type.GetType("System.DateTime"))
                    End If
                    dtOffers.Columns.Add("CustomerGroupID", System.Type.GetType("System.Int64"))
                    dtOffers.Columns.Add("EmployeesOnly", System.Type.GetType("System.Boolean"))
                    dtOffers.Columns.Add("EmployeesExcluded", System.Type.GetType("System.Boolean"))
          
          
                    If ((MyCommon.Fetch_CM_SystemOption(24) = "1") And (CardTypeID = 1)) Then
                        bGetHHAndMemberOffers = True
                        lHouseholdPK = lCustomerPK ' Both parameters should be the customerpk of the household record, as the logic of the
                        ' SP pa_LogixServ_FetchCustGroups_MemberOrHousehold that gets called
                        ' later is based on HHPK<>0 
                    End If
                    WriteDebug("Entering GetCustomerCmOffersLimitsBased", DebugState.CurrentTime)
                    bStatus = GetCustomerCmOffersLimitsBased(lCustomerPK, lHouseholdPK, dtOffers, bGetHHAndMemberOffers)
                    WriteDebug("Exiting GetCustomerCmOffersLimitsBased", DebugState.CurrentTime)


                    'Populate and return the DataSet
                    ResultSet.Tables.Add(dtStatus.Copy())
                    ResultSet.Tables.Add(dtOffers.Copy())


                End If
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
        WriteDebug("OfferListLimitsBasedCM_ByCardID - " & System.Net.Dns.GetHostName(), DebugState.EndTime)
        WriteLogLines()

        Return ResultSet
    End Function
    
    <WebMethod()> _
    Public Function OfferListLimitsBasedCPE(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As Integer) As System.Data.DataSet
        Dim dt As System.Data.DataTable
        Dim dtOffers As System.Data.DataTable
        Dim dtStatus As System.Data.DataTable
        Dim dtudfs As System.Data.DataTable
        Dim dtTransactions As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim lCustomerPK As Long
        Dim lHouseholdPK As Long
        Dim bStatus As Boolean
        Dim bGetHHAndMemberOffers As Boolean = False
        Dim iOpenedXS As Integer = 0
        Dim iOpenedRT As Integer = 0
        Dim ResultSet As New System.Data.DataSet("OfferListLimitsBasedCPE")

        InitApp()

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                MyCommon.Open_LogixRT()
                iOpenedRT = 1
            End If
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then
                MyCommon.Open_LogixXS()
                iOpenedXS = 1
            End If
            If Not IsValidGUID(GUID) Then
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
            ElseIf (CustomerID.Length < 1) Then
                'Bad customer ID8
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERID
                row.Item("Description") = "Failure: Invalid customer ID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
            Else
                'Pad the customer ID
                CustomerID = MyCommon.Pad_ExtCardID(CustomerID, CardTypes.CUSTOMER)
                'Find the Customer PK
                MyCommon.QueryStr = "select CID.CustomerPK, C.HHPK from CardIDs as CID with (NoLock) " & _
                                    "inner join Customers as C with (NoLock) on C.CustomerPK=CID.CustomerPK " & _
                                    "INNER JOIN CardTypes AS CT with (NoLock) ON CID.CardTypeID = CT.CardTypeID " & _
                                    "WHERE CT.CustTypeID=" & CustomerTypeID & " AND CID.ExtCardID='" & MyCryptLib.SQL_StringEncrypt(CustomerID, True) & "';"
                dt = MyCommon.LXS_Select
                If dt.Rows.Count = 0 Then
                    'Customer not found
                    If CustomerTypeID = 0 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_CUSTOMER
                        row.Item("Description") = "Failure: Customer " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    ElseIf CustomerTypeID = 1 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_HOUSEHOLD
                        row.Item("Description") = "Failure: Household " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    ElseIf CustomerTypeID = 2 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_CAM
                        row.Item("Description") = "Failure: CAM " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    Else
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERTYPEID
                        row.Item("Description") = "Failure: Invalid customer type."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    End If
                Else
                    lCustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)
                    lHouseholdPK = MyCommon.NZ(dt.Rows(0).Item("HHPK"), 0)

                    'Put a success into the Status table
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = 0
                    row.Item("Description") = "Success."
                    dtStatus.Rows.Add(row)


                    dtOffers = New DataTable
                    dtOffers.TableName = "Incentives"
                    dtOffers.Columns.Add("IncentiveID", System.Type.GetType("System.Int64"))
                    dtOffers.Columns.Add("IncentiveName", System.Type.GetType("System.String"))
                    dtOffers.Columns.Add("Description", System.Type.GetType("System.String"))
                    dtOffers.Columns.Add("StartDate", System.Type.GetType("System.DateTime"))
                    dtOffers.Columns.Add("EndDate", System.Type.GetType("System.DateTime"))
                    dtOffers.Columns.Add("CustomerGroupID", System.Type.GetType("System.Int64"))
                    dtOffers.Columns.Add("EmployeesOnly", System.Type.GetType("System.Boolean"))
                    dtOffers.Columns.Add("EmployeesExcluded", System.Type.GetType("System.Boolean"))
                    dtOffers.Columns.Add("UserDefinedFields", GetType(DataTable))
                   
                   
                    'Declaring 'dtudfs' table for userdefinedfields
                 
                    dtudfs = New DataTable
                    dtudfs.TableName = "UserDefinedField"
                    dtudfs.Columns.Add("ExternalID", System.Type.GetType("System.String"))
                    dtudfs.Columns.Add("UserDescription", System.Type.GetType("System.String"))
                    dtudfs.Columns.Add("Value", System.Type.GetType("System.String"))
                   
                    

                    
                    If ((MyCommon.Fetch_CM_SystemOption(24) = "1") And (CustomerTypeID = 1)) Then
                        bGetHHAndMemberOffers = True
                        lHouseholdPK = lCustomerPK ' Both parameters should be the customerpk of the household record, as the logic of the
                        ' SP pa_LogixServ_FetchCustGroups_MemberOrHousehold that gets called
                        ' later is based on HHPK<>0 
                    End If
                    bStatus = GetCustomerCpeOffersLimitsBased(lCustomerPK, lHouseholdPK, dtOffers, bGetHHAndMemberOffers, dtudfs)


                    'Populate and return the DataSet
                    ResultSet.Tables.Add(dtStatus.Copy())
                    ResultSet.Tables.Add(dtOffers.Copy())


                End If
            End If

        Catch ex As Exception
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
            row.Item("Description") = "Failure: Application " & ex.ToString
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed AndAlso iOpenedRT = 1 Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed AndAlso iOpenedXS = 1 Then MyCommon.Close_LogixXS()
        End Try

        Return ResultSet
    End Function
    <WebMethod()> _
    Public Function OfferListLimitsBasedCPE_ByCardID(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As Integer, ByVal CardTypeID As Integer) As System.Data.DataSet
        Dim dt As System.Data.DataTable
        Dim dtOffers As System.Data.DataTable
        Dim dtStatus As System.Data.DataTable
        Dim dtTransactions As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim dtudfs As New DataTable
        Dim lCustomerPK As Long
        Dim lHouseholdPK As Long
        Dim bStatus As Boolean
        Dim bGetHHAndMemberOffers As Boolean = False
        Dim iOpenedXS As Integer = 0
        Dim iOpenedRT As Integer = 0
        Dim ResultSet As New System.Data.DataSet("OfferListLimitsBasedCPE")

        InitApp()

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                MyCommon.Open_LogixRT()
                iOpenedRT = 1
            End If
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then
                MyCommon.Open_LogixXS()
                iOpenedXS = 1
            End If
            If Not IsValidGUID(GUID) Then
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
            ElseIf (CustomerID.Length < 1) Then
                'Bad customer ID8
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERID
                row.Item("Description") = "Failure: Invalid customer ID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
            Else
                'Pad the customer ID
                CustomerID = MyCommon.Pad_ExtCardID(CustomerID, CardTypeID)

                'Find the Customer PK
                MyCommon.QueryStr = "select CID.CustomerPK, C.HHPK from CardIDs as CID with (NoLock) " & _
                                    "inner join Customers as C with (NoLock) on C.CustomerPK=CID.CustomerPK " & _
                                    "where CID.ExtCardID='" & MyCryptLib.SQL_StringEncrypt(CustomerID, True) & "' and CID.CardTypeID = " & CardTypeID & ";"
                dt = MyCommon.LXS_Select
                If dt.Rows.Count = 0 Then
                    'Customer not found
                    If CustomerTypeID = 0 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_CUSTOMER
                        row.Item("Description") = "Failure: Customer " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    ElseIf CustomerTypeID = 1 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_HOUSEHOLD
                        row.Item("Description") = "Failure: Household " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    ElseIf CustomerTypeID = 2 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_CAM
                        row.Item("Description") = "Failure: CAM " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    Else
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERTYPEID
                        row.Item("Description") = "Failure: Invalid customer type."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    End If
                Else
                    lCustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)
                    lHouseholdPK = MyCommon.NZ(dt.Rows(0).Item("HHPK"), 0)

                    'Put a success into the Status table
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = 0
                    row.Item("Description") = "Success."
                    dtStatus.Rows.Add(row)


                    dtOffers = New DataTable
                    dtOffers.TableName = "Incentives"
                    dtOffers.Columns.Add("IncentiveID", System.Type.GetType("System.Int64"))
                    dtOffers.Columns.Add("IncentiveName", System.Type.GetType("System.String"))
                    dtOffers.Columns.Add("Description", System.Type.GetType("System.String"))
                    dtOffers.Columns.Add("StartDate", System.Type.GetType("System.DateTime"))
                    dtOffers.Columns.Add("EndDate", System.Type.GetType("System.DateTime"))
                    dtOffers.Columns.Add("CustomerGroupID", System.Type.GetType("System.Int64"))
                    dtOffers.Columns.Add("EmployeesOnly", System.Type.GetType("System.Boolean"))
                    dtOffers.Columns.Add("EmployeesExcluded", System.Type.GetType("System.Boolean"))
                    dtOffers.Columns.Add("UserDefinedFields", GetType(DataTable))
                    
                    'Declaring 'dtudfs' table for userdefinedfields
                    dtudfs = New DataTable
                    dtudfs.TableName = "UserDefinedField"
                    dtudfs.Columns.Add("ExternalID", System.Type.GetType("System.String"))
                    dtudfs.Columns.Add("UserDescription", System.Type.GetType("System.String"))
                    dtudfs.Columns.Add("Value", System.Type.GetType("System.String"))
                   
                    
                    If ((MyCommon.Fetch_CM_SystemOption(24) = "1") And (CardTypeID = 1)) Then
                        bGetHHAndMemberOffers = True
                        lHouseholdPK = lCustomerPK ' Both parameters should be the customerpk of the household record, as the logic of the
                        ' SP pa_LogixServ_FetchCustGroups_MemberOrHousehold that gets called
                        ' later is based on HHPK<>0 
                    End If
                    bStatus = GetCustomerCpeOffersLimitsBased(lCustomerPK, lHouseholdPK, dtOffers, bGetHHAndMemberOffers, dtudfs)


                    'Populate and return the DataSet
                    ResultSet.Tables.Add(dtStatus.Copy())
                    ResultSet.Tables.Add(dtOffers.Copy())


                End If
            End If

        Catch ex As Exception
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
            row.Item("Description") = "Failure: Application " & ex.ToString
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed AndAlso iOpenedRT = 1 Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed AndAlso iOpenedXS = 1 Then MyCommon.Close_LogixXS()
        End Try

        Return ResultSet
    End Function
  
    Private Function _ExternalOfferListCM(ByVal GUID As String, ByVal ExtCardID As String, ByVal CardTypeID As Integer, ByVal bIncludeRedeemed As Boolean, ByVal bIncludeUnRedeemed As Boolean) As System.Data.DataSet
        Dim dt As System.Data.DataTable
        Dim dtOffers As System.Data.DataTable
        Dim dtPrograms As System.Data.DataTable
        Dim dtStatus As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim lCustomerPK As Long
        Dim lHouseholdPK As Long
        Dim bStatus As Boolean
        Dim ResultSet As New System.Data.DataSet("OfferListCM")
        'added
        Dim RetCode As StatusCodes
        Dim RetMsg As String = ""

        'Initialize the status table, which will report the success or failure of the CustomerDetails operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            If Not IsValidGUID(GUID) Then
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
 
                'ElseIf (CustomerID.Length < 1) Then
            ElseIf Not IsValidCustomerCard(ExtCardID, CardTypeID, RetCode, RetMsg) Then
                
                'Bad customer ID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = RetCode
                row.Item("Description") = RetMsg
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())

            Else
                'Pad the customer ID
                ExtCardID = MyCommon.Pad_ExtCardID(ExtCardID, CardTypeID)
    
                'Find the Customer PK
                MyCommon.QueryStr = "select CID.CustomerPK, C.HHPK from CardIDs as CID with (NoLock) " & _
                                    "inner join Customers as C with (NoLock) on C.CustomerPK=CID.CustomerPK " & _
                                    "where CID.ExtCardID='" & MyCryptLib.SQL_StringEncrypt(ExtCardID) & "' and CID.CardTypeID = " & CardTypeID & " ;"

                dt = MyCommon.LXS_Select
                If dt.Rows.Count = 0 Then
                    'Customer not found
                    If CardTypeID = 0 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_CUSTOMER
                        row.Item("Description") = "Failure: Card " & ExtCardID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    ElseIf CardTypeID = 1 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_HOUSEHOLD
                        row.Item("Description") = "Failure: Household " & ExtCardID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    ElseIf CardTypeID = 2 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_CAM
                        row.Item("Description") = "Failure: CAM " & ExtCardID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    Else
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERTYPEID
                        row.Item("Description") = "Failure: Invalid card type."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    End If
                Else
                    lCustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)
                    lHouseholdPK = MyCommon.NZ(dt.Rows(0).Item("HHPK"), 0)

                    'Put a success into the Status table
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = 0
                    row.Item("Description") = "Success."
                    dtStatus.Rows.Add(row)


                    dtOffers = New DataTable
                    dtOffers.TableName = "Offers"
                    dtOffers.Columns.Add("OfferID", System.Type.GetType("System.Int64"))
                    dtOffers.Columns.Add("Name", System.Type.GetType("System.String"))
                    dtOffers.Columns.Add("Description", System.Type.GetType("System.String"))
                    dtOffers.Columns.Add("OfferCategory", System.Type.GetType("System.String"))
                    dtOffers.Columns.Add("StartDate", System.Type.GetType("System.DateTime"))
                    dtOffers.Columns.Add("EndDate", System.Type.GetType("System.DateTime"))
                    dtOffers.Columns.Add("CustomerGroupID", System.Type.GetType("System.Int64"))
                    dtOffers.Columns.Add("EmployeesOnly", System.Type.GetType("System.Boolean"))
                    dtOffers.Columns.Add("EmployeesExcluded", System.Type.GetType("System.Boolean"))

                    dtPrograms = New DataTable
                    dtPrograms.TableName = "Programs"
                    dtPrograms.Columns.Add("OfferID", System.Type.GetType("System.Int64"))
                    dtPrograms.Columns.Add("ConditionOrReward", System.Type.GetType("System.String"))
                    dtPrograms.Columns.Add("ProgramType", System.Type.GetType("System.String"))
                    dtPrograms.Columns.Add("ProgramID", System.Type.GetType("System.Int64"))
                    dtPrograms.Columns.Add("TierLevel", System.Type.GetType("System.Int32"))
                    dtPrograms.Columns.Add("Amount", System.Type.GetType("System.Decimal"))
                    dtPrograms.Columns.Add("Name", System.Type.GetType("System.String"))
                    dtPrograms.Columns.Add("Description", System.Type.GetType("System.String"))
                    dtPrograms.Columns.Add("Category", System.Type.GetType("System.String"))
                    dtPrograms.Columns.Add("WebMessage", System.Type.GetType("System.String"))

                    bStatus = GetCustomerExternalCmOffers(lCustomerPK, lHouseholdPK, dtOffers, dtPrograms, bIncludeRedeemed, bIncludeUnRedeemed)


                    'Populate and return the DataSet
                    ResultSet.Tables.Add(dtStatus.Copy())
                    ResultSet.Tables.Add(dtOffers.Copy())
                    ResultSet.Tables.Add(dtPrograms.Copy())

                End If
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

    <WebMethod()> _
    Public Function ExtendedExternalOfferListCM_Class(ByVal GUID As String, ByVal ExtCardID As String, ByVal CardTypeID As Integer, ByVal IncludeRedeemed As Boolean, _
                                          ByVal IncludeUnRedeemed As Boolean, ByVal IncludeExpired As Boolean, ByVal IncludeActive As Boolean) As OfferListCMClass
        Dim ds As System.Data.DataSet
        Dim oOfferList As New OfferListCMClass
        Dim oStatus As New CustWebStatus
        Dim dt As DataTable
        Dim dr As DataRow
        Dim i As Integer
        Dim bProcessMore As Boolean = True

        Try
            InitApp()

            'RT 5596; RT 5636; RT 5637 - If includeactive and includeexpired or if includeredeemed and includeunredeemed are false, throw an error
            If (IncludeExpired = False AndAlso IncludeActive = False) Then
                oStatus.StatusCode = StatusCodes.INVALID_MODE
                oStatus.Description = "Failure: Excluding all active and all expired offers will return no offers"
                oOfferList.Status = oStatus
                bProcessMore = False
            ElseIf (IncludeRedeemed = False AndAlso IncludeUnRedeemed = False) Then
                oStatus.StatusCode = StatusCodes.INVALID_MODE
                oStatus.Description = "Failure: Excluding all redeemed and all unredeemed offers will return no offers"
                oOfferList.Status = oStatus
                bProcessMore = False
            ElseIf IncludeExpired = False And IncludeActive = True Then
                'External offer list- does not include expired offers
                oOfferList = ExternalOfferListCM_Class(GUID, ExtCardID, CardTypeID, IncludeRedeemed, IncludeUnRedeemed, IncludeExpired)
                bProcessMore = False
            ElseIf (IncludeExpired = True) And (IncludeActive = False) And (IncludeRedeemed = True) And (IncludeUnRedeemed = True) Then
                'Expired offer list- contains both redeemed and unredeemed offers
                oOfferList = ExpiredExternalOfferListCM_Class(GUID, ExtCardID, CardTypeID)
                bProcessMore = False
                'Handle [Active redeemed and expired redeemed offers only] case
            ElseIf (IncludeExpired = True) And (IncludeActive = True) And (IncludeRedeemed = True) And (IncludeUnRedeemed = False) Then
                'Return all Redeemed offers
                ds = _ExtendedExternalOfferListCM(GUID, ExtCardID, CardTypeID, IncludeRedeemed, IncludeUnRedeemed, IncludeExpired, IncludeActive)
        
                'Handle [Expired unredeemed offers only] case
            ElseIf (IncludeExpired = True) And (IncludeActive = False) And (IncludeRedeemed = False) And (IncludeUnRedeemed = True) Then
                'Expired offer list- contains unredeemed offers
                ds = _ExtendedExternalOfferListCM(GUID, ExtCardID, CardTypeID, IncludeRedeemed, IncludeUnRedeemed, IncludeExpired, IncludeActive)
        
                'Handle [Expired redeemed offers only] case
            ElseIf (IncludeExpired = True) And (IncludeActive = False) And (IncludeRedeemed = True) And (IncludeUnRedeemed = False) Then
                'Expired offer list- contains redeemed offers
                ds = _ExtendedExternalOfferListCM(GUID, ExtCardID, CardTypeID, IncludeRedeemed, IncludeUnRedeemed, IncludeExpired, IncludeActive)
            End If

            'If oOfferList has already been returned from another class, skip over this
            If bProcessMore AndAlso ds IsNot Nothing Then
                dt = ds.Tables("Status")
                oStatus.StatusCode = dt.Rows(0).Item("StatusCode")
                oStatus.Description = dt.Rows(0).Item("Description")
                oOfferList.Status = oStatus
                If oStatus.StatusCode = "0" Then
                    dt = ds.Tables("Offers")
                    If dt.Rows.Count > 0 Then
                        ReDim oOfferList.Offers(dt.Rows.Count - 1)
                        Dim oOffer As Offer
                        i = 0
                        For Each dr In dt.Rows
                            oOffer = New Offer
                            oOffer.OfferID = dr.Item("OfferID")
                            oOffer.Name = dr.Item("Name")
                            oOffer.Description = dr.Item("Description")
                            oOffer.OfferCategory = dr.Item("OfferCategory")
                            oOffer.StartDate = dr.Item("StartDate")
                            oOffer.EndDate = dr.Item("EndDate")
                            oOffer.CustomerGroupID = dr.Item("CustomerGroupID")
                            oOffer.EmployeesOnly = dr.Item("EmployeesOnly").ToString.ToLower
                            oOffer.EmployeesExcluded = dr.Item("EmployeesExcluded").ToString.ToLower
                            oOfferList.Offers(i) = oOffer
                            i += 1
                        Next
                    End If
                    dt = ds.Tables("Programs")
                    If dt.Rows.Count > 0 Then
                        ReDim oOfferList.Programs(dt.Rows.Count - 1)
                        Dim p As Program
                        i = 0
                        For Each dr In dt.Rows
                            p = New Program
                            p.OfferID = dr.Item("OfferID")
                            p.ConditionOrReward = dr.Item("ConditionOrReward")
                            p.ProgramType = dr.Item("ProgramType")
                            p.ProgramID = dr.Item("ProgramID")
                            p.TierLevel = dr.Item("TierLevel")
                            p.Amount = dr.Item("Amount")
                            p.Name = dr.Item("Name")
                            oOfferList.Programs(i) = p
                            i += 1
                        Next
                    End If
                End If
            End If
        Catch ex As Exception
            oStatus.StatusCode = StatusCodes.APPLICATION_EXCEPTION
            oStatus.Description = "Failure: Application " & ex.ToString
            oOfferList.Status = oStatus

        Finally

        End Try
        Return oOfferList

    End Function

    Private Function _ExtendedExternalOfferListCM(ByVal GUID As String, ByVal ExtCardID As String, ByVal CardTypeID As Integer, ByVal bIncludeRedeemed As Boolean, _
                                      ByVal bIncludeUnRedeemed As Boolean, ByVal bIncludeExpired As Boolean, ByVal bIncludeActive As Boolean) As System.Data.DataSet
        Dim dt As System.Data.DataTable
        Dim dtOffers As System.Data.DataTable
        Dim dtPrograms As System.Data.DataTable
        Dim dtStatus As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim lCustomerPK As Long
        Dim lHouseholdPK As Long
        Dim bStatus As Boolean
        Dim ResultSet As New System.Data.DataSet("OfferListCM")

        'Initialize the status table, which will report the success or failure of the CustomerDetails operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))
        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            If Not IsValidGUID(GUID) Then
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
            ElseIf (ExtCardID.Length < 1) Then
                'Bad customer ID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERID
                row.Item("Description") = "Failure: Invalid Card ID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
            Else
                'Pad the customer ID
                ExtCardID = MyCommon.Pad_ExtCardID(ExtCardID, CardTypeID)
            
                'Find the Customer PK
                MyCommon.QueryStr = "select CID.CustomerPK, C.HHPK from CardIDs as CID with (NoLock) " & _
                                    "inner join Customers as C with (NoLock) on C.CustomerPK=CID.CustomerPK " & _
                                    "where CID.ExtCardID='" & MyCryptLib.SQL_StringEncrypt(ExtCardID) & "' and CID.CardTypeID = " & CardTypeID & " ;"

                dt = MyCommon.LXS_Select
                If dt.Rows.Count = 0 Then
                    'Customer not found
                    If CardTypeID = 0 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_CUSTOMER
                        row.Item("Description") = "Failure: Card " & ExtCardID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    ElseIf CardTypeID = 1 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_HOUSEHOLD
                        row.Item("Description") = "Failure: Household " & ExtCardID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    ElseIf CardTypeID = 2 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_CAM
                        row.Item("Description") = "Failure: CAM " & ExtCardID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    Else
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERTYPEID
                        row.Item("Description") = "Failure: Invalid card type."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    End If
                Else
                    lCustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)
                    lHouseholdPK = MyCommon.NZ(dt.Rows(0).Item("HHPK"), 0)

                    'Put a success into the Status table
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = 0
                    row.Item("Description") = "Success."
                    dtStatus.Rows.Add(row)


                    dtOffers = New DataTable
                    dtOffers.TableName = "Offers"
                    dtOffers.Columns.Add("OfferID", System.Type.GetType("System.Int64"))
                    dtOffers.Columns.Add("Name", System.Type.GetType("System.String"))
                    dtOffers.Columns.Add("Description", System.Type.GetType("System.String"))
                    dtOffers.Columns.Add("OfferCategory", System.Type.GetType("System.String"))
                    dtOffers.Columns.Add("StartDate", System.Type.GetType("System.DateTime"))
                    dtOffers.Columns.Add("EndDate", System.Type.GetType("System.DateTime"))
                    dtOffers.Columns.Add("CustomerGroupID", System.Type.GetType("System.Int64"))
                    dtOffers.Columns.Add("EmployeesOnly", System.Type.GetType("System.Boolean"))
                    dtOffers.Columns.Add("EmployeesExcluded", System.Type.GetType("System.Boolean"))

                    dtPrograms = New DataTable
                    dtPrograms.TableName = "Programs"
                    dtPrograms.Columns.Add("OfferID", System.Type.GetType("System.Int64"))
                    dtPrograms.Columns.Add("ConditionOrReward", System.Type.GetType("System.String"))
                    dtPrograms.Columns.Add("ProgramType", System.Type.GetType("System.String"))
                    dtPrograms.Columns.Add("ProgramID", System.Type.GetType("System.Int64"))
                    dtPrograms.Columns.Add("TierLevel", System.Type.GetType("System.Int32"))
                    dtPrograms.Columns.Add("Amount", System.Type.GetType("System.Decimal"))
                    dtPrograms.Columns.Add("Name", System.Type.GetType("System.String"))
                    dtPrograms.Columns.Add("Description", System.Type.GetType("System.String"))
                    dtPrograms.Columns.Add("Category", System.Type.GetType("System.String"))
                    dtPrograms.Columns.Add("WebMessage", System.Type.GetType("System.String"))

                    bStatus = GetCustomerExtendedExternalCmOffers(lCustomerPK, lHouseholdPK, dtOffers, dtPrograms, bIncludeRedeemed, _
                                                          bIncludeUnRedeemed, bIncludeExpired, bIncludeActive)

                    'Populate and return the DataSet
                    ResultSet.Tables.Add(dtStatus.Copy())
                    ResultSet.Tables.Add(dtOffers.Copy())
                    ResultSet.Tables.Add(dtPrograms.Copy())

                End If
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

    Private Function GetCustomerExtendedExternalCmOffers(ByVal lCustomerPK As Long, ByVal lHouseholdPK As Long, ByRef dtOffers As DataTable, _
                                                          ByRef dtPrograms As DataTable, ByVal bIncludeRedeemed As Boolean, ByVal bIncludeUnRedeemed As Boolean, _
                                                          ByVal bIncludeExpired As Boolean, ByVal bIncludeActive As Boolean) As Boolean

        Const sDateFormatBusStart As String = "yyyy-MM-ddT23:59:59"
        Const sDateFormatBusEnd As String = "yyyy-MM-ddT00:00:00"
        Const sDateFormat As String = "yyyy-MM-ddTHH:mm:ss"

        Dim dt As DataTable = Nothing
        Dim dr As DataRow
        Dim dtOffer As DataTable
        Dim dtCustLimit As DataTable
        Dim drOffer As DataRow
        Dim sGroupCustomerList As String
        Dim sBusDate As String
        Dim sBusDateStart As String
        Dim sBusDateEnd As String
        Dim dateBusinessDate As Date = Now
        Dim sOfferId As String
        Dim i As Integer
        Dim iFilterEmp As Integer
        Dim iNonEmpOnly As Integer
        Dim bstatus As Boolean = True
        Dim iCategoryId As Integer
        Dim sCategory As String
        Dim iAdvancedLimitID As Integer
        Dim bOfferLimitReached As Boolean = False
        Dim iPromoVarID As Integer
        Dim dLimitValue As Decimal
        Dim dCustAmount As Decimal = 0
        Dim iPeriod As Integer
        Dim iDefaultLanguage As Integer = 0
        Dim iOfferId As Integer = 0

        sBusDate = Format(dateBusinessDate, sDateFormat)
        sBusDateStart = "'" & Format(dateBusinessDate, sDateFormatBusStart) & "'"
        sBusDateEnd = "'" & Format(dateBusinessDate, sDateFormatBusEnd) & "'"

        MyCommon.QueryStr = "dbo.pa_LogixServ_FetchCustGroups"
        MyCommon.Open_LXSsp()
        MyCommon.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = lCustomerPK
        MyCommon.LXSsp.Parameters.Add("@HHPK", SqlDbType.BigInt).Value = lHouseholdPK
        dt = MyCommon.LXSsp_select
        MyCommon.Close_LXSsp()

        If dt.Rows.Count > 0 Then
            sGroupCustomerList = "(" & dt.Rows(0).Item(0).ToString
            If dt.Rows.Count > 1 Then
                For i = 1 To dt.Rows.Count - 1
                    sGroupCustomerList += "," & dt.Rows(i).Item(0)
                Next
            End If
            sGroupCustomerList += ")"
        Else
            ' Member, but no specific groups assigned
            sGroupCustomerList = "()"
        End If

        MyCommon.QueryStr = "select distinct OfferId from CM_ST_OfferCustLocView with (NoLock) where" & _
                            " ((AnyCustomer = 0 and AnyCardholder = 0 and InCustGroupId in " & sGroupCustomerList & _
                            " and (ExCustGroupId is null or ExCustGroupId not in " & sGroupCustomerList & "))" & _
                            " or (AnyCardholder = 1 and (ExCustGroupId is null or ExCustGroupId not in " & sGroupCustomerList & "))) order by OfferId;"
        dt = MyCommon.LRT_Select
        If dt.Rows.Count > 0 Then
            For Each dr In dt.Rows
                sOfferId = dr.Item(0)
                iOfferId = Integer.Parse(MyCommon.NZ(sOfferId, "-1"))

                MyCommon.QueryStr = "select O.OfferId, O.Name,O.Description,O.OfferCategoryID,O.ProdStartDate,O.ProdEndDate," & _
                                    "O.DistPeriodLimit,O.DistPeriod,O.DistPeriodVarID,O.AdvancedLimitID," & _
                                    "O.NumTiers,O.EmployeeFiltering,O.NonEmployeesOnly,O.DisplayOnWebKiosk,OC.LinkID " & _
                                    "from CM_ST_Offers as O with (NoLock) " & _
                                    "inner join CM_ST_OfferConditions as OC with (NoLock) on OC.OfferID=O.OfferId and OC.ConditionTypeID=1 " & _
                                    "where O.OfferId=" & sOfferId & " AND O.ExtOfferID is not null;"
                dtOffer = MyCommon.LRT_Select
                If dtOffer.Rows.Count > 0 Then
                    'Read data and keep it for later comparison
                    If Not IsDBNull(dtOffer.Rows(0).Item("AdvancedLimitID")) Then
                        iAdvancedLimitID = dtOffer.Rows(0).Item("AdvancedLimitID")
                    Else
                        iAdvancedLimitID = 0
                    End If
                    If Not IsDBNull(dtOffer.Rows(0).Item("DistPeriodVarID")) Then
                        iPromoVarID = dtOffer.Rows(0).Item("DistPeriodVarID")
                    Else
                        iPromoVarID = 0
                    End If
                    If Not IsDBNull(dtOffer.Rows(0).Item("DistPeriod")) Then
                        iPeriod = dtOffer.Rows(0).Item("DistPeriod")
                    Else
                        iPeriod = 0
                    End If
                    If Not IsDBNull(dtOffer.Rows(0).Item("DistPeriodLimit")) Then
                        dLimitValue = dtOffer.Rows(0).Item("DistPeriodLimit")
                    Else
                        dLimitValue = 0
                    End If

                    If (Date.Compare(MyCommon.NZ(dtOffer.Rows(0).Item("ProdEndDate"), Now.Date), Now.Date) < 0 AndAlso bIncludeExpired = True) OrElse _
                    (Date.Compare(MyCommon.NZ(dtOffer.Rows(0).Item("ProdEndDate"), Now.Date), Now.Date) > 0 AndAlso bIncludeActive = True) OrElse _
                    (bIncludeExpired = True AndAlso bIncludeActive = True AndAlso bIncludeRedeemed = True) Then
                        If (iPeriod <> 0) Then ' Limits need to be checked only for Period <>0
                            'We will ignore transaction limits and only consider day and customer based limits. 
                            'Hence O.DistPeriod <> 0 check
                            'Handle Advance Limits
                            If (iAdvancedLimitID > 0) Then 'Get promoVar and limit for the advanced limit

                                MyCommon.QueryStr = "Select Amount FROM  CM_AdvancedLimitVariables WHERE PromoVarID = " & iPromoVarID & " AND " & _
                                                    " CustomerPK = " & lCustomerPK & " "
                                dtCustLimit = MyCommon.LXS_Select
                                If dtCustLimit.Rows.Count > 0 Then 'If entry exists  in CM_AdvancedLimitVariables for promovarid and customer
                                    dCustAmount = dtCustLimit.Rows(0).Item("Amount")
                                Else
                                    dCustAmount = 0
                                End If
                            Else  '(iAdvancedLimitID <= 0)
                                ' If not advance limit, it is a regular limit.
                                MyCommon.QueryStr = "Select Amount FROM  DistributionVariables WHERE PromoVarID = " & iPromoVarID & " AND " & _
                                                    " CustomerPK = " & lCustomerPK & " "
                                dtCustLimit = MyCommon.LXS_Select
                                If dtCustLimit.Rows.Count > 0 Then 'If entry exists  in DistributionVariables for promovarid and customer
                                    dCustAmount = dtCustLimit.Rows(0).Item("Amount")
                                Else
                                    dCustAmount = 0
                                End If
                            End If

                            If (dCustAmount > 0) And (dCustAmount >= dLimitValue) Then
                                bOfferLimitReached = True
                                If (bIncludeRedeemed = True) Then
                                    drOffer = dtOffers.NewRow()
                                    drOffer.Item("OfferID") = sOfferId
                                    If Not IsDBNull(dtOffer.Rows(0).Item("Name")) Then
                                        drOffer.Item("Name") = dtOffer.Rows(0).Item("Name")
                                    End If
                                    If Not IsDBNull(dtOffer.Rows(0).Item("ProdEndDate")) Then
                                        drOffer.Item("EndDate") = dtOffer.Rows(0).Item("ProdEndDate")
                                    End If
                                    If Not IsDBNull(dtOffer.Rows(0).Item("Description")) Then
                                        drOffer.Item("Description") = dtOffer.Rows(0).Item("Description")
                                    End If
                                    iCategoryId = MyCommon.NZ(dtOffer.Rows(0).Item("OfferCategoryID"), 0)
                                    sCategory = GetCategoryFromId(iCategoryId)
                                    If Not IsDBNull(sCategory) Then
                                        drOffer.Item("OfferCategory") = sCategory
                                    Else : drOffer.Item("OfferCategory") = ""
                                    End If
                                    If Not IsDBNull(dtOffer.Rows(0).Item("ProdStartDate")) Then
                                        drOffer.Item("StartDate") = dtOffer.Rows(0).Item("ProdStartDate")
                                    End If
                                    If Not IsDBNull(dtOffer.Rows(0).Item("LinkID")) Then
                                        drOffer.Item("CustomerGroupID") = dtOffer.Rows(0).Item("LinkID")
                                    End If
                                    If Not IsDBNull(dtOffer.Rows(0).Item("EmployeeFiltering")) Then
                                        iFilterEmp = dtOffer.Rows(0).Item("EmployeeFiltering")
                                    End If
                                    If Not IsDBNull(dtOffer.Rows(0).Item("NonEmployeesOnly")) Then
                                        iNonEmpOnly = dtOffer.Rows(0).Item("NonEmployeesOnly")
                                    End If
                                    If iFilterEmp Then
                                        If iNonEmpOnly Then
                                            drOffer.Item("EmployeesOnly") = 0
                                            drOffer.Item("EmployeesExcluded") = 1
                                        Else
                                            drOffer.Item("EmployeesOnly") = 1
                                            drOffer.Item("EmployeesExcluded") = 0
                                        End If
                                    Else
                                        drOffer.Item("EmployeesOnly") = 0
                                        drOffer.Item("EmployeesExcluded") = 0
                                    End If
                                    dtOffers.Rows.Add(drOffer)
                                    'Get program information for redeemed offers
                                    bstatus = GetProgramsForCMOffer(sOfferId, dtPrograms)
                                End If  'IncludeReemed

                            Else
                                bOfferLimitReached = False
                                If (bIncludeUnRedeemed = True) Then
                                    drOffer = dtOffers.NewRow()
                                    drOffer.Item("OfferID") = sOfferId
                                    If Not IsDBNull(dtOffer.Rows(0).Item("Name")) Then
                                        drOffer.Item("Name") = dtOffer.Rows(0).Item("Name")
                                    End If
                                    If Not IsDBNull(dtOffer.Rows(0).Item("ProdEndDate")) Then
                                        drOffer.Item("EndDate") = dtOffer.Rows(0).Item("ProdEndDate")
                                    End If
                                    If Not IsDBNull(dtOffer.Rows(0).Item("Description")) Then
                                        drOffer.Item("Description") = dtOffer.Rows(0).Item("Description")
                                    End If
                                    iCategoryId = MyCommon.NZ(dtOffer.Rows(0).Item("OfferCategoryID"), 0)
                                    sCategory = GetCategoryFromId(iCategoryId)
                                    If Not IsDBNull(sCategory) Then
                                        drOffer.Item("OfferCategory") = sCategory
                                    Else : drOffer.Item("OfferCategory") = ""
                                    End If
                                    If Not IsDBNull(dtOffer.Rows(0).Item("ProdStartDate")) Then
                                        drOffer.Item("StartDate") = dtOffer.Rows(0).Item("ProdStartDate")
                                    End If
                                    If Not IsDBNull(dtOffer.Rows(0).Item("LinkID")) Then
                                        drOffer.Item("CustomerGroupID") = dtOffer.Rows(0).Item("LinkID")
                                    End If
                                    If Not IsDBNull(dtOffer.Rows(0).Item("EmployeeFiltering")) Then
                                        iFilterEmp = dtOffer.Rows(0).Item("EmployeeFiltering")
                                    End If
                                    If Not IsDBNull(dtOffer.Rows(0).Item("NonEmployeesOnly")) Then
                                        iNonEmpOnly = dtOffer.Rows(0).Item("NonEmployeesOnly")
                                    End If
                                    If iFilterEmp Then
                                        If iNonEmpOnly Then
                                            drOffer.Item("EmployeesOnly") = 0
                                            drOffer.Item("EmployeesExcluded") = 1
                                        Else
                                            drOffer.Item("EmployeesOnly") = 1
                                            drOffer.Item("EmployeesExcluded") = 0
                                        End If
                                    Else
                                        drOffer.Item("EmployeesOnly") = 0
                                        drOffer.Item("EmployeesExcluded") = 0
                                    End If
                                    dtOffers.Rows.Add(drOffer)
                                    bstatus = GetProgramsForCMOffer(sOfferId, dtPrograms)
                                End If ' bIncludeUnRedeemed = true

                            End If 'dcustamount
                        Else
                            drOffer = dtOffers.NewRow()
                            drOffer.Item("OfferID") = sOfferId
                            If Not IsDBNull(dtOffer.Rows(0).Item("Name")) Then
                                drOffer.Item("Name") = dtOffer.Rows(0).Item("Name")
                            End If
                            If Not IsDBNull(dtOffer.Rows(0).Item("Description")) Then
                                drOffer.Item("Description") = dtOffer.Rows(0).Item("Description")
                            End If
                            iCategoryId = MyCommon.NZ(dtOffer.Rows(0).Item("OfferCategoryID"), 0)
                            sCategory = GetCategoryFromId(iCategoryId)
                            If Not IsDBNull(sCategory) Then
                                drOffer.Item("OfferCategory") = sCategory
                            Else : drOffer.Item("OfferCategory") = ""
                            End If
                            If Not IsDBNull(dtOffer.Rows(0).Item("ProdStartDate")) Then
                                drOffer.Item("StartDate") = dtOffer.Rows(0).Item("ProdStartDate")
                            End If
                            If Not IsDBNull(dtOffer.Rows(0).Item("ProdEndDate")) Then
                                drOffer.Item("EndDate") = dtOffer.Rows(0).Item("ProdEndDate")
                            End If
                            If Not IsDBNull(dtOffer.Rows(0).Item("LinkID")) Then
                                drOffer.Item("CustomerGroupID") = dtOffer.Rows(0).Item("LinkID")
                            End If
                            If Not IsDBNull(dtOffer.Rows(0).Item("EmployeeFiltering")) Then
                                iFilterEmp = dtOffer.Rows(0).Item("EmployeeFiltering")
                            End If
                            If Not IsDBNull(dtOffer.Rows(0).Item("NonEmployeesOnly")) Then
                                iNonEmpOnly = dtOffer.Rows(0).Item("NonEmployeesOnly")
                            End If
                            If iFilterEmp Then
                                If iNonEmpOnly Then
                                    drOffer.Item("EmployeesOnly") = 0
                                    drOffer.Item("EmployeesExcluded") = 1
                                Else
                                    drOffer.Item("EmployeesOnly") = 1
                                    drOffer.Item("EmployeesExcluded") = 0
                                End If
                            Else
                                drOffer.Item("EmployeesOnly") = 0
                                drOffer.Item("EmployeesExcluded") = 0
                            End If
                            dtOffers.Rows.Add(drOffer)
                            bstatus = GetProgramsForCMOffer(sOfferId, dtPrograms)
                        End If '(iPeriod <> 0)
                    End If 'Date.compare
                End If 'dtOffer.Rows.Count > 0

            Next

            If dtOffers.Rows.Count > 0 Then dtOffers.AcceptChanges()
            If dtPrograms.Rows.Count > 0 Then dtPrograms.AcceptChanges()

        Else
            bstatus = False
        End If

        Return bstatus


    End Function

    <WebMethod()> _
    Public Function TransactionHistoryCM(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As String, ByVal StartDate As Date, ByVal EndDate As Date) As System.Data.DataSet
        InitApp()
        'added
        'Dim iCustomerTypeID As Integer = -1
        'If Not String.IsNullOrEmpty(CustomerTypeID) AndAlso Not CustomerTypeID.Trim = String.Empty AndAlso IsNumeric(CustomerTypeID) Then iCustomerTypeID = Convert.ToInt32(CustomerTypeID)
        Return _TransactionHistoryCM(GUID, CustomerID, CustomerTypeID, StartDate, EndDate)
    End Function

    Public Class Transaction
        Public LogixTransactionNumber As String = ""
        Public TransactionDate As String = ""
        Public LocationCode As String = ""
        Public LocationName As String = ""
        Public CardID As String = ""
        Public HouseholdID As String = ""
    End Class

    Public Class TransactionPointsProgram
        Public LogixTransactionNumber As String = ""
        Public ID As String = ""
        Public Amount As String = ""
    End Class

    Public Class TransactionStoredValueProgram
        Public LogixTransactionNumber As String = ""
        Public ID As String = ""
        Public Amount As String = ""
        Public Action As String = ""
        Public Status As String = ""
        Public ExpirationDate As String = ""
    End Class

    Public Class TransactionHistoryCMClass
        Public Status As CustWebStatus
        Public Transactions() As Transaction
        Public PointsPrograms() As TransactionPointsProgram
        Public StoredValuePrograms() As TransactionStoredValueProgram
    End Class

    <WebMethod()> _
    Public Function TransactionHistoryCM_Class(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As String, ByVal StartDate As Date, ByVal EndDate As Date) As TransactionHistoryCMClass
        Dim ds As System.Data.DataSet
        Dim oTransactionHistoryCM As New TransactionHistoryCMClass
        Dim oStatus As New CustWebStatus
        Dim dt As DataTable
        Dim dr As DataRow
        Dim i As Integer
        'added
        'Dim iCustomerTypeID As Integer = -1
        'If Not String.IsNullOrEmpty(CustomerTypeID) AndAlso Not CustomerTypeID.Trim = String.Empty AndAlso IsNumeric(CustomerTypeID) Then iCustomerTypeID = Convert.ToInt32(CustomerTypeID)
       

        InitApp()

        ds = _TransactionHistoryCM(GUID, CustomerID, CustomerTypeID, StartDate, EndDate, False)

        dt = ds.Tables("Status")
        oStatus.StatusCode = dt.Rows(0).Item("StatusCode")
        oStatus.Description = dt.Rows(0).Item("Description")
        oTransactionHistoryCM.Status = oStatus
        If oStatus.StatusCode = "0" Then
            dt = ds.Tables("Transactions")
            If dt.Rows.Count > 0 Then
                ReDim oTransactionHistoryCM.Transactions(dt.Rows.Count - 1)
                Dim oTransaction As Transaction
                i = 0
                For Each dr In dt.Rows
                    oTransaction = New Transaction
                    oTransaction.LogixTransactionNumber = dr.Item("LogixTransactionNumber")
                    oTransaction.TransactionDate = dr.Item("Date")
                    oTransaction.LocationCode = dr.Item("LocationCode")
                    oTransaction.LocationName = dr.Item("LocationName")
                    oTransaction.CardID = dr.Item("CardID")
                    oTransaction.HouseholdID = dr.Item("HouseHoldID")
                    oTransactionHistoryCM.Transactions(i) = oTransaction
                    i += 1
                Next
            End If
            dt = ds.Tables("StoredValues")
            If dt.Rows.Count > 0 Then
                ReDim oTransactionHistoryCM.StoredValuePrograms(dt.Rows.Count - 1)
                Dim sv As TransactionStoredValueProgram
                i = 0
                For Each dr In dt.Rows
                    sv = New TransactionStoredValueProgram
                    sv.LogixTransactionNumber = dr.Item("LogixTransactionNumber")
                    sv.ID = dr.Item("ProgramID")
                    sv.Amount = dr.Item("Amount")
                    sv.Action = dr.Item("Action")
                    sv.Status = dr.Item("Status")
                    sv.ExpirationDate = dr.Item("ExpirationDate")
                    oTransactionHistoryCM.StoredValuePrograms(i) = sv
                    i += 1
                Next
            End If
            dt = ds.Tables("Points")
            If dt.Rows.Count > 0 Then
                ReDim oTransactionHistoryCM.PointsPrograms(dt.Rows.Count - 1)
                Dim p As TransactionPointsProgram
                i = 0
                For Each dr In dt.Rows
                    p = New TransactionPointsProgram
                    p.LogixTransactionNumber = dr.Item("LogixTransactionNumber")
                    p.ID = dr.Item("ProgramID")
                    p.Amount = dr.Item("Amount")
                    oTransactionHistoryCM.PointsPrograms(i) = p
                    i += 1
                Next
            End If
        End If

        Return oTransactionHistoryCM
    End Function
    <WebMethod()> _
    Public Function TransactionHistoryCM_ClassByCardID(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As String, ByVal CardTypeID As Integer, ByVal StartDate As Date, ByVal EndDate As Date) As TransactionHistoryCMClass
        Dim ds As System.Data.DataSet
        Dim oTransactionHistoryCM As New TransactionHistoryCMClass
        Dim oStatus As New CustWebStatus
        Dim dt As DataTable
        Dim dr As DataRow
        Dim i As Integer
        'added
        ' Dim iCustomerTypeID As Integer = -1
        'If Not String.IsNullOrEmpty(CustomerTypeID) AndAlso Not CustomerTypeID.Trim = String.Empty AndAlso IsNumeric(CustomerTypeID) Then iCustomerTypeID = Convert.ToInt32(CustomerTypeID)
       

        InitApp()

        ds = _TransactionHistoryCM(GUID, CustomerID, CustomerTypeID, StartDate, EndDate, CardTypeID)

        dt = ds.Tables("Status")
        oStatus.StatusCode = dt.Rows(0).Item("StatusCode")
        oStatus.Description = dt.Rows(0).Item("Description")
        oTransactionHistoryCM.Status = oStatus
        If oStatus.StatusCode = "0" Then
            dt = ds.Tables("Transactions")
            If dt.Rows.Count > 0 Then
                ReDim oTransactionHistoryCM.Transactions(dt.Rows.Count - 1)
                Dim oTransaction As Transaction
                i = 0
                For Each dr In dt.Rows
                    oTransaction = New Transaction
                    oTransaction.LogixTransactionNumber = dr.Item("LogixTransactionNumber")
                    oTransaction.TransactionDate = dr.Item("Date")
                    oTransaction.LocationCode = dr.Item("LocationCode")
                    oTransaction.LocationName = dr.Item("LocationName")
                    oTransaction.CardID = dr.Item("CardID")
                    oTransaction.HouseholdID = dr.Item("HouseHoldID")
                    oTransactionHistoryCM.Transactions(i) = oTransaction
                    i += 1
                Next
            End If
            dt = ds.Tables("StoredValues")
            If dt.Rows.Count > 0 Then
                ReDim oTransactionHistoryCM.StoredValuePrograms(dt.Rows.Count - 1)
                Dim sv As TransactionStoredValueProgram
                i = 0
                For Each dr In dt.Rows
                    sv = New TransactionStoredValueProgram
                    sv.LogixTransactionNumber = dr.Item("LogixTransactionNumber")
                    sv.ID = dr.Item("ProgramID")
                    sv.Amount = dr.Item("Amount")
                    sv.Action = dr.Item("Action")
                    sv.Status = dr.Item("Status")
                    sv.ExpirationDate = dr.Item("ExpirationDate")
                    oTransactionHistoryCM.StoredValuePrograms(i) = sv
                    i += 1
                Next
            End If
            dt = ds.Tables("Points")
            If dt.Rows.Count > 0 Then
                ReDim oTransactionHistoryCM.PointsPrograms(dt.Rows.Count - 1)
                Dim p As TransactionPointsProgram
                i = 0
                For Each dr In dt.Rows
                    p = New TransactionPointsProgram
                    p.LogixTransactionNumber = dr.Item("LogixTransactionNumber")
                    p.ID = dr.Item("ProgramID")
                    p.Amount = dr.Item("Amount")
                    oTransactionHistoryCM.PointsPrograms(i) = p
                    i += 1
                Next
            End If
        End If

        Return oTransactionHistoryCM
    End Function

    Private Function _TransactionHistoryCM(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As String, ByVal StartDate As Date, ByVal EndDate As Date, Optional ByVal CardTypeID As Integer = -1) As System.Data.DataSet
        Dim dt As System.Data.DataTable
        Dim dtTransactions As System.Data.DataTable
        Dim dtPoints As System.Data.DataTable
        Dim dtStoredValues As System.Data.DataTable
        Dim dtStatus As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim lCustomerPK As Long
        Dim lHouseholdPK As Long
        Dim bStatus As Boolean
        Dim ResultSet As New System.Data.DataSet("TransactionHistoryCM")
        'added
        Dim RetCode As StatusCodes
        Dim RetMsg As String = ""

        'Initialize the status table, which will report the success or failure of the CustomerDetails operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
            If MyCommon.LWHadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixWH()

            If Not IsValidGUID(GUID) Then
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                'ElseIf (CustomerID.Length < 1) Then
            ElseIf Not IsValidCustomerCard(CustomerID, CustomerTypeID, RetCode, RetMsg) Then
                
                'Bad customer ID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = RetCode
                row.Item("Description") = RetMsg
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
         
            Else
                Dim iCustomerTypeID As Integer = -1
                If Not String.IsNullOrEmpty(CustomerTypeID) AndAlso Not CustomerTypeID.Trim = String.Empty AndAlso IsNumeric(CustomerTypeID) Then iCustomerTypeID = Convert.ToInt32(CustomerTypeID)
                If CardTypeID <> -1 Then
                    'Pad the customer ID
                    CustomerID = MyCommon.Pad_ExtCardID(CustomerID, CardTypeID)
                Else
                    'Pad the customer ID
                    CustomerID = MyCommon.Pad_ExtCardID(CustomerID, CardTypes.CUSTOMER)
                End If
                'Find the Customer PK
                MyCommon.QueryStr = "select CID.CustomerPK, C.HHPK from CardIDs as CID with (NoLock) " & _
                                    "inner join Customers as C with (NoLock) on C.CustomerPK=CID.CustomerPK " & _
                                    "INNER JOIN CardTypes AS CT with (NoLock) ON CID.CardTypeID = CT.CardTypeID " & _
                                            "WHERE CT.CustTypeID=" & iCustomerTypeID & " AND CID.ExtCardID='" & MyCryptLib.SQL_StringEncrypt(CustomerID, True) & "';"
                dt = MyCommon.LXS_Select
                If dt.Rows.Count = 0 Then
                    'Customer not found
                    If iCustomerTypeID = 0 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_CUSTOMER
                        row.Item("Description") = "Failure: Customer " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    ElseIf iCustomerTypeID = 1 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_HOUSEHOLD
                        row.Item("Description") = "Failure: Household " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    ElseIf iCustomerTypeID = 2 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_CAM
                        row.Item("Description") = "Failure: CAM " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    Else
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERTYPEID
                        row.Item("Description") = "Failure: Invalid customer type."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    End If
                Else
                    lCustomerPK = dt.Rows(0).Item("CustomerPK")
                    lHouseholdPK = dt.Rows(0).Item("HHPK")

                    'Put a success into the Status table
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = 0
                    row.Item("Description") = "Success."
                    dtStatus.Rows.Add(row)

                    dtTransactions = New DataTable
                    dtTransactions.TableName = "Transactions"
                    dtTransactions.Columns.Add("LogixTransactionNumber", System.Type.GetType("System.String"))
                    dtTransactions.Columns.Add("Date", System.Type.GetType("System.DateTime"))
                    dtTransactions.Columns.Add("LocationCode", System.Type.GetType("System.String"))
                    dtTransactions.Columns.Add("LocationName", System.Type.GetType("System.String"))
                    dtTransactions.Columns.Add("CardID", System.Type.GetType("System.String"))
                    dtTransactions.Columns.Add("HouseHoldID", System.Type.GetType("System.String"))

                    dtStoredValues = New DataTable
                    dtStoredValues.TableName = "StoredValues"
                    dtStoredValues.Columns.Add("LogixTransactionNumber", System.Type.GetType("System.String"))
                    dtStoredValues.Columns.Add("ProgramID", System.Type.GetType("System.Int64"))
                    dtStoredValues.Columns.Add("Amount", System.Type.GetType("System.Decimal"))
                    dtStoredValues.Columns.Add("Action", System.Type.GetType("System.String"))
                    dtStoredValues.Columns.Add("Status", System.Type.GetType("System.String"))
                    dtStoredValues.Columns.Add("ExpirationDate", System.Type.GetType("System.DateTime"))

                    dtPoints = New DataTable
                    dtPoints.TableName = "Points"
                    dtPoints.Columns.Add("LogixTransactionNumber", System.Type.GetType("System.String"))
                    dtPoints.Columns.Add("ProgramID", System.Type.GetType("System.Int64"))
                    dtPoints.Columns.Add("Amount", System.Type.GetType("System.Decimal"))

                    bStatus = GetCustomerTransactions(lCustomerPK, lHouseholdPK, CustomerID, iCustomerTypeID, StartDate, EndDate, dtTransactions, dtPoints, dtStoredValues)

                    'Populate and return the DataSet
                    ResultSet.Tables.Add(dtStatus.Copy())
                    ResultSet.Tables.Add(dtTransactions.Copy())
                    ResultSet.Tables.Add(dtStoredValues.Copy())
                    ResultSet.Tables.Add(dtPoints.Copy())

                End If
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
            If MyCommon.LWHadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixWH()
        End Try

        Return ResultSet
    End Function

    <WebMethod()> _
    Public Function PointsBalancesCM(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As String) As DataSet
        InitApp()
        'added
        'Dim iCustomerTypeID As Integer = -1
        'If Not String.IsNullOrEmpty(CustomerTypeID) AndAlso Not CustomerTypeID.Trim = String.Empty AndAlso IsNumeric(CustomerTypeID) Then iCustomerTypeID = Convert.ToInt32(CustomerTypeID)
       
        Return _PointsBalancesCM(GUID, CustomerID, CustomerTypeID)
    End Function

    Public Class PointsProgram
        Public ID As String = 0
        Public Name As String = ""
        Public Category As String = ""
        Public Balance As Long = 0
    End Class

    Public Class PointsBalancesCMClass
        Public Status As CustWebStatus
        Public PointsPrograms() As PointsProgram
    End Class

    <WebMethod()> _
    Public Function PointsBalancesCM_Class(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As String) As PointsBalancesCMClass
        Dim ds As System.Data.DataSet
        Dim oPointsBalancesCM As New PointsBalancesCMClass
        Dim oStatus As New CustWebStatus
        Dim dt As DataTable
        Dim dr As DataRow
        Dim i As Integer
        'added
        'Dim iCustomerTypeID As Integer = -1
        'If Not String.IsNullOrEmpty(CustomerTypeID) AndAlso Not CustomerTypeID.Trim = String.Empty AndAlso IsNumeric(CustomerTypeID) Then iCustomerTypeID = Convert.ToInt32(CustomerTypeID)
       

        InitApp()
        ds = _PointsBalancesCM(GUID, CustomerID, CustomerTypeID)

        dt = ds.Tables("Status")
        oStatus.StatusCode = dt.Rows(0).Item("StatusCode")
        oStatus.Description = dt.Rows(0).Item("Description")
        oPointsBalancesCM.Status = oStatus
        If oStatus.StatusCode = "0" Then
            dt = ds.Tables("PointsProgram")
            If dt.Rows.Count > 0 Then
                ReDim oPointsBalancesCM.PointsPrograms(dt.Rows.Count - 1)
                Dim pp As PointsProgram
                i = 0
                For Each dr In dt.Rows
                    pp = New PointsProgram
                    pp.ID = dr.Item("ProgramID")
                    pp.Name = dr.Item("ProgramName")
                    pp.Category = dr.Item("Category")
                    pp.Balance = dr.Item("Balance")
                    oPointsBalancesCM.PointsPrograms(i) = pp
                    i += 1
                Next
            End If
        End If

        Return oPointsBalancesCM

    End Function
    
    <WebMethod()> _
    Public Function PointsBalancesCM_Class_ByCardID(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As String, ByVal CardTypeID As Integer) As PointsBalancesCMClass
        Dim ds As System.Data.DataSet
        Dim oPointsBalancesCM As New PointsBalancesCMClass
        Dim oStatus As New CustWebStatus
        Dim dt As DataTable
        Dim dr As DataRow
        Dim i As Integer
        'added
        'Dim iCustomerTypeID As Integer = -1
        'If Not String.IsNullOrEmpty(CustomerTypeID) AndAlso Not CustomerTypeID.Trim = String.Empty AndAlso IsNumeric(CustomerTypeID) Then iCustomerTypeID = Convert.ToInt32(CustomerTypeID)
       


        InitApp()

        ds = _PointsBalancesCM(GUID, CustomerID, CustomerTypeID, CardTypeID)

        dt = ds.Tables("Status")
        oStatus.StatusCode = dt.Rows(0).Item("StatusCode")
        oStatus.Description = dt.Rows(0).Item("Description")
        oPointsBalancesCM.Status = oStatus
        If oStatus.StatusCode = "0" Then
            dt = ds.Tables("PointsProgram")
            If dt.Rows.Count > 0 Then
                ReDim oPointsBalancesCM.PointsPrograms(dt.Rows.Count - 1)
                Dim pp As PointsProgram
                i = 0
                For Each dr In dt.Rows
                    pp = New PointsProgram
                    pp.ID = dr.Item("ProgramID")
                    pp.Name = dr.Item("ProgramName")
                    pp.Category = dr.Item("Category")
                    pp.Balance = dr.Item("Balance")
                    oPointsBalancesCM.PointsPrograms(i) = pp
                    i += 1
                Next
            End If
        End If

        Return oPointsBalancesCM

    End Function

    Private Function _PointsBalancesCM(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As String, Optional ByVal CardTypeID As Integer = -1) As DataSet
        Dim dt As System.Data.DataTable
        Dim dtBalances As System.Data.DataTable
        Dim dtStatus As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim lCustomerPK As Long
        Dim lHouseholdPK As Long
        Dim bStatus As Boolean
        Dim ResultSet As New System.Data.DataSet("PointsBalancesCM")
        'added
        Dim RetCode As StatusCodes
        Dim RetMsg As String = ""

        'Initialize the status table, which will report the success or failure of the CustomerDetails operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            If Not IsValidGUID(GUID) Then
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                'ElseIf (CustomerID.Length < 1) Then
            ElseIf Not IsValidCustomerCard(CustomerID, CustomerTypeID, RetCode, RetMsg) Then
                
                'Bad customer ID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = RetCode
                row.Item("Description") = RetMsg
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
            
            Else
                Dim iCustomerTypeID As Integer = -1
                If Not String.IsNullOrEmpty(CustomerTypeID) AndAlso Not CustomerTypeID.Trim = String.Empty AndAlso IsNumeric(CustomerTypeID) Then iCustomerTypeID = Convert.ToInt32(CustomerTypeID)
                'Pad the customer ID
                If CardTypeID <> -1 Then
                    CustomerID = MyCommon.Pad_ExtCardID(CustomerID, CardTypeID)
                Else
                    CustomerID = MyCommon.Pad_ExtCardID(CustomerID, CardTypes.CUSTOMER)
                End If
                'Find the Customer PK
                MyCommon.QueryStr = "select CID.CustomerPK, C.HHPK from CardIDs as CID with (NoLock) " & _
                                    "inner join Customers as C with (NoLock) on C.CustomerPK=CID.CustomerPK " & _
                                    "INNER JOIN CardTypes AS CT with (NoLock) ON CID.CardTypeID = CT.CardTypeID " & _
                                            "WHERE CT.CustTypeID=" & iCustomerTypeID & " AND CID.ExtCardID='" & MyCryptLib.SQL_StringEncrypt(CustomerID, True) & "';"
                dt = MyCommon.LXS_Select
                If dt.Rows.Count = 0 Then
                    'Customer not found
                    If iCustomerTypeID = 0 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_CUSTOMER
                        row.Item("Description") = "Failure: Customer " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    ElseIf iCustomerTypeID = 1 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_HOUSEHOLD
                        row.Item("Description") = "Failure: Household " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    ElseIf iCustomerTypeID = 2 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_CAM
                        row.Item("Description") = "Failure: CAM " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    Else
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERTYPEID
                        row.Item("Description") = "Failure: Invalid customer type."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    End If
                Else
                    lCustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)
                    lHouseholdPK = MyCommon.NZ(dt.Rows(0).Item("HHPK"), 0)

                    'Put a success into the Status table
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = 0
                    row.Item("Description") = "Success."
                    dtStatus.Rows.Add(row)


                    dtBalances = New DataTable("PointsProgram")
                    dtBalances.Columns.Add("ProgramID", System.Type.GetType("System.Int64"))
                    dtBalances.Columns.Add("ProgramName", System.Type.GetType("System.String"))
                    dtBalances.Columns.Add("Category", System.Type.GetType("System.String"))
                    dtBalances.Columns.Add("Balance", System.Type.GetType("System.Decimal"))

                    bStatus = GetPointsBalances(CustomerID, lCustomerPK, lHouseholdPK, dtBalances)

                    'Populate and return the DataSet
                    ResultSet.Tables.Add(dtStatus.Copy())
                    ResultSet.Tables.Add(dtBalances.Copy())

                End If
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

    <WebMethod()> _
    Public Function StoredValueBalancesCM(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As String, ByVal AboutToExpireDays As Integer) As DataSet
        InitApp()
        'added
        'Dim iCustomerTypeID As Integer = -1
        'If Not String.IsNullOrEmpty(CustomerTypeID) AndAlso Not CustomerTypeID.Trim = String.Empty AndAlso IsNumeric(CustomerTypeID) Then iCustomerTypeID = Convert.ToInt32(CustomerTypeID)
        Return _StoredValueBalancesCM(GUID, CustomerID, CustomerTypeID, AboutToExpireDays)
    End Function

    Public Class StoredValueProgram
        Public ID As String = ""
        Public Name As String = ""
        Public Value As Decimal = 0.0
        Public UnitOfMeasureLimit As Integer = 0
        Public Balance As Long = 0
        Public BalanceExpireDate As String = ""
        Public AboutToExpireQuantity As Integer = 0
        Public AboutToExpireDays As Integer = 0
    End Class

    Public Class StoredValueBalancesCMClass
        Public Status As CustWebStatus
        Public StoredValuePrograms() As StoredValueProgram
    End Class

    <WebMethod()> _
    Public Function StoredValueBalancesCM_Class(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As String, ByVal AboutToExpireDays As Integer) As StoredValueBalancesCMClass
        Dim ds As System.Data.DataSet
        Dim oStoredValueBalancesCM As New StoredValueBalancesCMClass
        Dim oStatus As New CustWebStatus
        Dim dt As DataTable
        Dim dr As DataRow
        Dim i As Integer
        'added
        'Dim iCustomerTypeID As Integer = -1
        'If Not String.IsNullOrEmpty(CustomerTypeID) AndAlso Not CustomerTypeID.Trim = String.Empty AndAlso IsNumeric(CustomerTypeID) Then iCustomerTypeID = Convert.ToInt32(CustomerTypeID)
       


        InitApp()

        ds = _StoredValueBalancesCM(GUID, CustomerID, CustomerTypeID, AboutToExpireDays)

        dt = ds.Tables("Status")
        oStatus.StatusCode = dt.Rows(0).Item("StatusCode")
        oStatus.Description = dt.Rows(0).Item("Description")
        oStoredValueBalancesCM.Status = oStatus
        If oStatus.StatusCode = "0" Then
            dt = ds.Tables("StoredValueProgram")
            If dt.Rows.Count > 0 Then
                ReDim oStoredValueBalancesCM.StoredValuePrograms(dt.Rows.Count - 1)
                Dim sv As StoredValueProgram
                i = 0
                For Each dr In dt.Rows
                    sv = New StoredValueProgram
                    sv.ID = dr.Item("ProgramID")
                    sv.Name = dr.Item("ProgramName")
                    sv.Value = dr.Item("Value")
                    sv.UnitOfMeasureLimit = dr.Item("UnitOfMeasureLimit")
                    sv.Balance = dr.Item("Balance")
                    sv.BalanceExpireDate = dr.Item("BalanceExpireDate")
                    sv.AboutToExpireQuantity = dr.Item("AboutToExpireQuantity")
                    sv.AboutToExpireDays = dr.Item("AboutToExpireDays")
                    oStoredValueBalancesCM.StoredValuePrograms(i) = sv
                    i += 1
                Next
            End If
        End If

        Return oStoredValueBalancesCM

    End Function
    <WebMethod()> _
    Public Function StoredValueBalancesCM_Class_ByCardID(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As String, ByVal CardTypeID As Integer, ByVal AboutToExpireDays As Integer) As StoredValueBalancesCMClass
        Dim ds As System.Data.DataSet
        Dim oStoredValueBalancesCM As New StoredValueBalancesCMClass
        Dim oStatus As New CustWebStatus
        Dim dt As DataTable
        Dim dr As DataRow
        Dim i As Integer
        'added
        'Dim iCustomerTypeID As Integer = -1
        'If Not String.IsNullOrEmpty(CustomerTypeID) AndAlso Not CustomerTypeID.Trim = String.Empty AndAlso IsNumeric(CustomerTypeID) Then iCustomerTypeID = Convert.ToInt32(CustomerTypeID)
        InitApp()

        ds = _StoredValueBalancesCM(GUID, CustomerID, CustomerTypeID, AboutToExpireDays, CardTypeID)

        dt = ds.Tables("Status")
        oStatus.StatusCode = dt.Rows(0).Item("StatusCode")
        oStatus.Description = dt.Rows(0).Item("Description")
        oStoredValueBalancesCM.Status = oStatus
        If oStatus.StatusCode = "0" Then
            dt = ds.Tables("StoredValueProgram")
            If dt.Rows.Count > 0 Then
                ReDim oStoredValueBalancesCM.StoredValuePrograms(dt.Rows.Count - 1)
                Dim sv As StoredValueProgram
                i = 0
                For Each dr In dt.Rows
                    sv = New StoredValueProgram
                    sv.ID = dr.Item("ProgramID")
                    sv.Name = dr.Item("ProgramName")
                    sv.Value = dr.Item("Value")
                    sv.UnitOfMeasureLimit = dr.Item("UnitOfMeasureLimit")
                    sv.Balance = dr.Item("Balance")
                    sv.BalanceExpireDate = dr.Item("BalanceExpireDate")
                    sv.AboutToExpireQuantity = dr.Item("AboutToExpireQuantity")
                    sv.AboutToExpireDays = dr.Item("AboutToExpireDays")
                    oStoredValueBalancesCM.StoredValuePrograms(i) = sv
                    i += 1
                Next
            End If
        End If

        Return oStoredValueBalancesCM

    End Function

    Private Function _StoredValueBalancesCM(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As String, ByVal AboutToExpireDays As Integer, Optional ByVal CardTypeID As Integer = -1) As DataSet
        Dim dt As System.Data.DataTable
        Dim dtBalances As System.Data.DataTable
        Dim dtStatus As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim lCustomerPK As Long
        Dim lHouseholdPK As Long
        Dim bStatus As Boolean
        Dim ResultSet As New System.Data.DataSet("StoredValueBalancesCM")
        'added
        Dim RetCode As StatusCodes
        Dim RetMsg As String = ""

        'Initialize the status table, which will report the success or failure of the CustomerDetails operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            If Not IsValidGUID(GUID) Then
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                'ElseIf (CustomerID.Length < 1) Then
            ElseIf Not IsValidCustomerCard(CustomerID, CustomerTypeID, RetCode, RetMsg) Then
                
                'Bad customer ID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = RetCode
                row.Item("Description") = RetMsg
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())

            Else
                Dim iCustomerTypeID As Integer = -1
                If Not String.IsNullOrEmpty(CustomerTypeID) AndAlso Not CustomerTypeID.Trim = String.Empty AndAlso IsNumeric(CustomerTypeID) Then iCustomerTypeID = Convert.ToInt32(CustomerTypeID)
                'Pad the customer ID
                If CardTypeID <> -1 Then
                    CustomerID = MyCommon.Pad_ExtCardID(CustomerID, CardTypeID)
                Else
                    CustomerID = MyCommon.Pad_ExtCardID(CustomerID, CardTypes.CUSTOMER)
                End If
                'Find the Customer PK
                MyCommon.QueryStr = "select CID.CustomerPK, C.HHPK from CardIDs as CID with (NoLock) " & _
                                    "inner join Customers as C with (NoLock) on C.CustomerPK=CID.CustomerPK " & _
                                    "INNER JOIN CardTypes AS CT with (NoLock) ON CID.CardTypeID = CT.CardTypeID " & _
                                            "WHERE CT.CustTypeID=" & iCustomerTypeID & " AND CID.ExtCardID='" & MyCryptLib.SQL_StringEncrypt(CustomerID, True) & "';"
                dt = MyCommon.LXS_Select
                If dt.Rows.Count = 0 Then
                    'Customer not found
                    If iCustomerTypeID = 0 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_CUSTOMER
                        row.Item("Description") = "Failure: Customer " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    ElseIf iCustomerTypeID = 1 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_HOUSEHOLD
                        row.Item("Description") = "Failure: Household " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    ElseIf iCustomerTypeID = 2 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_CAM
                        row.Item("Description") = "Failure: CAM " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    Else
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERTYPEID
                        row.Item("Description") = "Failure: Invalid customer type."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    End If
                Else
                    lCustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)
                    lHouseholdPK = MyCommon.NZ(dt.Rows(0).Item("HHPK"), 0)

                    'Put a success into the Status table
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = 0
                    row.Item("Description") = "Success."
                    dtStatus.Rows.Add(row)


                    dtBalances = New DataTable("StoredValueProgram")
                    dtBalances.Columns.Add("ProgramID", System.Type.GetType("System.Int64"))
                    dtBalances.Columns.Add("ProgramName", System.Type.GetType("System.String"))
                    dtBalances.Columns.Add("Value", System.Type.GetType("System.Decimal"))
                    dtBalances.Columns.Add("UnitOfMeasureLimit", System.Type.GetType("System.Int32"))
                    dtBalances.Columns.Add("Balance", System.Type.GetType("System.Int32"))
                    dtBalances.Columns.Add("BalanceExpireDate", System.Type.GetType("System.DateTime"))
                    dtBalances.Columns.Add("AboutToExpireQuantity", System.Type.GetType("System.Int32"))
                    dtBalances.Columns.Add("AboutToExpireDays", System.Type.GetType("System.Int32"))

                    bStatus = GetStoredValueBalances(lCustomerPK, lHouseholdPK, AboutToExpireDays, dtBalances)

                    'Populate and return the DataSet
                    ResultSet.Tables.Add(dtStatus.Copy())
                    ResultSet.Tables.Add(dtBalances.Copy())

                End If
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


    <WebMethod()> _
    Public Function GetStoredValuePoints(ByVal GUID As String, ByVal CardID As String, ByVal CardTypeID As String) As DataSet
        Dim dt As System.Data.DataTable
        Dim dtBalances As System.Data.DataTable
        Dim dtStatus As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim lCustomerPK As Long
        Dim lHouseholdPK As Long
        Dim bStatus As Boolean
        Dim ResultSet As New System.Data.DataSet("StoredValuePoints")
        'added
        'Dim iCardTypeID As Integer = -1
        'If Not String.IsNullOrEmpty(CardTypeID) AndAlso Not CardTypeID.Trim = String.Empty AndAlso IsNumeric(CardTypeID) Then iCardTypeID = Convert.ToInt32(CardTypeID)
       
        Dim RetCode As StatusCodes
        Dim RetMsg As String = ""
        

        InitApp()

        'Initialize the status table, which will report the success or failure of the CustomerDetails operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            If Not IsValidGUID(GUID) Then
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                ' ElseIf (CardID.Length < 1) Then
                'Bad Card ID
            ElseIf Not IsValidCustomerCard(CardID, CardTypeID, RetCode, RetMsg) Then
                
                'Bad customer ID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = RetCode
                row.Item("Description") = RetMsg
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
            Else
                Dim iCardTypeID As Integer = -1
                If Not String.IsNullOrEmpty(CardTypeID) AndAlso Not CardTypeID.Trim = String.Empty AndAlso IsNumeric(CardTypeID) Then iCardTypeID = Convert.ToInt32(CardTypeID)
                'Pad the customer ID
                CardID = MyCommon.Pad_ExtCardID(CardID, iCardTypeID)

                'Find the Customer PK
                MyCommon.QueryStr = "select CID.CustomerPK, C.HHPK from CardIDs as CID with (NoLock) " & _
                                    "inner join Customers as C with (NoLock) on C.CustomerPK=CID.CustomerPK " & _
                                    "INNER JOIN CardTypes AS CT with (NoLock) ON CID.CardTypeID = CT.CardTypeID " & _
                                    "WHERE CT.CustTypeID=" & CardTypeID & " AND CID.ExtCardID='" & MyCryptLib.SQL_StringEncrypt(CardID, True) & "';"
                dt = MyCommon.LXS_Select
                If dt.Rows.Count = 0 Then
                    'Customer not found
                    If CardTypeID = 0 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_CUSTOMER
                        row.Item("Description") = "Failure: Card " & CardID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    ElseIf CardTypeID = 1 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_HOUSEHOLD
                        row.Item("Description") = "Failure: Household " & CardID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    ElseIf CardTypeID = 2 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_CAM
                        row.Item("Description") = "Failure: CAM " & CardID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    Else
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERTYPEID
                        row.Item("Description") = "Failure: Invalid Card type."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    End If
                Else
                    lCustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)
                    lHouseholdPK = MyCommon.NZ(dt.Rows(0).Item("HHPK"), 0)

                    'Put a success into the Status table
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = 0
                    row.Item("Description") = "Success."
                    dtStatus.Rows.Add(row)


                    dtBalances = New DataTable("StoredValuePointsProgram")
                    dtBalances.Columns.Add("ProgramID", System.Type.GetType("System.Int64"))
                    dtBalances.Columns.Add("ExpirationDate", System.Type.GetType("System.DateTime"))
                    dtBalances.Columns.Add("Balance", System.Type.GetType("System.Int32"))

                    bStatus = GetAllSVBalancesForCustomer(lCustomerPK, lHouseholdPK, dtBalances)

                    'Populate and return the DataSet
                    ResultSet.Tables.Add(dtStatus.Copy())
                    ResultSet.Tables.Add(dtBalances.Copy())

                End If
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



    Private Function GetCustomerCmOffers(ByVal lCustomerPK As Long, ByVal lHouseholdPK As Long, ByRef dtOffers As DataTable, ByRef dtPrograms As DataTable) As Boolean
        Const sDateFormatBusStart As String = "yyyy-MM-ddT23:59:59"
        Const sDateFormatBusEnd As String = "yyyy-MM-ddT00:00:00"
        Const sDateFormat As String = "yyyy-MM-ddTHH:mm:ss"

        Dim dt As New DataTable()
        Dim dr As DataRow
        Dim CGdt As DataTable = Nothing
        Dim CGdr As DataRow
        Dim dtOffer As DataTable
        Dim drOffer As DataRow
        Dim cgXml As String
        Dim sBusDate As String
        Dim sBusDateStart As String
        Dim sBusDateEnd As String
        Dim dateBusinessDate As Date = Now
        Dim sOfferId As String
        Dim i As Integer
        Dim iFilterEmp As Integer
        Dim iNonEmpOnly As Integer
        Dim bstatus As Boolean = True
        Dim iCategoryId As Integer
        Dim sCategory As String
        Dim reader As SqlDataReader = Nothing
        Dim dtDisp As DataTable

        sBusDate = Format(dateBusinessDate, sDateFormat)
        sBusDateStart = Format(dateBusinessDate, sDateFormatBusStart)
        sBusDateEnd = Format(dateBusinessDate, sDateFormatBusEnd)

        MyCommon.QueryStr = "dbo.pa_LogixServ_FetchCustGroups"
        MyCommon.Open_LXSsp()
        MyCommon.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = lCustomerPK
        MyCommon.LXSsp.Parameters.Add("@HHPK", SqlDbType.BigInt).Value = lHouseholdPK
        WriteDebug("Begin Execution of pa_LogixServ_FetchCustGroups", DebugState.CurrentTime)
        CGdt = MyCommon.LXSsp_select
        WriteDebug("Completed Execution of pa_LogixServ_FetchCustGroups", DebugState.CurrentTime)
        MyCommon.Close_LXSsp()

        cgXml = "<customergroups><id>1</id><id>2</id>"
        If CGdt.Rows.Count > 0 Then
            For Each CGdr In CGdt.Rows
                cgXml &= "<id>" & MyCommon.NZ(CGdr.Item(0), "") & "</id>"
            Next
        End If
        cgXml &= "</customergroups>"
    
        WriteDebug("Customer Groups XML: " & cgXml, DebugState.CurrentTime)
    
        MyCommon.QueryStr = "dbo.pa_CMCustomerWebOffersList"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@cgXml", SqlDbType.Xml).Value = cgXml
        MyCommon.LRTsp.Parameters.Add("@BusDateStart", SqlDbType.NVarChar).Value = sBusDateStart
        MyCommon.LRTsp.Parameters.Add("@BusDateEnd", SqlDbType.NVarChar).Value = sBusDateEnd
    
        WriteDebug("Begin Execution of OffersID query", DebugState.CurrentTime)
        reader = MyCommon.LRTsp.ExecuteReader
        WriteDebug("Completed Execution of OffersID Query", DebugState.CurrentTime)
    
        Try
            dt.Load(reader)
        Catch ex As Exception
            WriteDebug("Exception " & ex.GetType.Name & ":" & ex.Message(), DebugState.CurrentTime)
        End Try
    

        WriteDebug("OfferIDs loaded into data table with row count=" & dt.Rows.Count, DebugState.CurrentTime)
          
        MyCommon.Close_LRTsp()
        reader.Close()
    
        If dt.Rows.Count > 0 Then
            WriteDebug("Begin Processing OfferIDs", DebugState.CurrentTime)
            For Each dr In dt.Rows
                sOfferId = dr.Item(0)
                MyCommon.QueryStr = "select O.OfferId, O.Name,O.Description,O.OfferCategoryID,O.ProdStartDate,O.ProdEndDate," & _
                                    "O.NumTiers,O.EmployeeFiltering,O.NonEmployeesOnly,O.DisplayOnWebKiosk,OC.LinkID " & _
                                    "from CM_ST_Offers as O with (NoLock) " & _
                                    "inner join CM_ST_OfferConditions as OC with (NoLock) on OC.OfferID=O.OfferId and OC.ConditionTypeID=1 " & _
                                    "where O.OfferId=" & sOfferId & ";"

                dtOffer = MyCommon.LRT_Select
                If dtOffer.Rows.Count > 0 Then
                    drOffer = dtOffers.NewRow()
                    drOffer.Item("OfferID") = sOfferId
                    drOffer.Item("Name") = dtOffer.Rows(0).Item("Name")
                    drOffer.Item("Description") = dtOffer.Rows(0).Item("Description")
                    iCategoryId = MyCommon.NZ(dtOffer.Rows(0).Item("OfferCategoryID"), 0)
                    sCategory = GetCategoryFromId(iCategoryId)
                    drOffer.Item("OfferCategory") = sCategory
                    drOffer.Item("StartDate") = dtOffer.Rows(0).Item("ProdStartDate")
                    drOffer.Item("EndDate") = dtOffer.Rows(0).Item("ProdEndDate")
                    drOffer.Item("CustomerGroupID") = dtOffer.Rows(0).Item("LinkID")

                    If (MyCommon.Fetch_CM_SystemOption(85) = "1") Then
                        MyCommon.QueryStr = "SELECT DisplayStartDate,DisplayEndDate FROM OfferAccessoryFields with (NoLock) WHERE OfferID = " & sOfferId & " "
                        dtDisp = MyCommon.LRT_Select
                        If dtDisp.Rows.Count > 0 Then
                            drOffer.Item("DisplayStartDate") = MyCommon.NZ(dtDisp.Rows(0).Item("DisplayStartDate"), DBNull.Value)
                            drOffer.Item("DisplayEndDate") = MyCommon.NZ(dtDisp.Rows(0).Item("DisplayEndDate"), DBNull.Value)
                        Else
                            drOffer.Item("DisplayStartDate") = DBNull.Value
                            drOffer.Item("DisplayEndDate") = DBNull.Value
                        End If
                    End If
                    
                    iFilterEmp = dtOffer.Rows(0).Item("EmployeeFiltering")
                    iNonEmpOnly = dtOffer.Rows(0).Item("NonEmployeesOnly")
                    If iFilterEmp Then
                        If iNonEmpOnly Then
                            drOffer.Item("EmployeesOnly") = 0
                            drOffer.Item("EmployeesExcluded") = 1
                        Else
                            drOffer.Item("EmployeesOnly") = 1
                            drOffer.Item("EmployeesExcluded") = 0
                        End If
                    Else
                        drOffer.Item("EmployeesOnly") = 0
                        drOffer.Item("EmployeesExcluded") = 0
                    End If
                    dtOffers.Rows.Add(drOffer)
                    bstatus = GetProgramsForCMOffer(sOfferId, dtPrograms)
                End If
            Next
            WriteDebug("End Processing OfferIDs", DebugState.CurrentTime)
            If dtOffers.Rows.Count > 0 Then dtOffers.AcceptChanges()
            If dtPrograms.Rows.Count > 0 Then dtPrograms.AcceptChanges()
        Else
            bstatus = False
        End If

        Return bstatus
    End Function

    Private Function GetCustomerCmOffersLimitsBased(ByVal lCustomerPK As Long, ByVal lHouseholdPK As Long, ByRef dtOffers As DataTable, ByVal bGetHHAndMemberOffers As Boolean) As Boolean
        Const sDateFormatBusStart As String = "yyyy-MM-ddT23:59:59"
        Const sDateFormatBusEnd As String = "yyyy-MM-ddT00:00:00"
        Const sDateFormat As String = "yyyy-MM-ddTHH:mm:ss"

        Dim dt As New DataTable()
        Dim dr As DataRow
        Dim CGdt As DataTable = Nothing
        Dim CGdr As DataRow
        Dim dtOffer As DataTable
        Dim dtCustLimit As DataTable
        Dim drOffer As DataRow
        Dim cgXML As String
        Dim sBusDate As String
        Dim sBusDateStart As String
        Dim sBusDateEnd As String
        Dim dateBusinessDate As Date = Now
        Dim sOfferId As String
        Dim i As Integer
        Dim iFilterEmp As Integer
        Dim iNonEmpOnly As Integer
        Dim bstatus As Boolean = True
        Dim iCategoryId As Integer
        Dim sCategory As String
        Dim iAdvancedLimitID As Integer
        Dim bOfferLimitReached As Boolean = False
        Dim iPromoVarID As Integer
        Dim dLimitValue As Decimal
        Dim dCustAmount As Decimal = 0
        Dim iPeriod As Integer
        Dim rstDispDates As DataTable
        Dim reader As SqlDataReader = Nothing


        sBusDate = Format(dateBusinessDate, sDateFormat)
        sBusDateStart = Format(dateBusinessDate, sDateFormatBusStart)
        sBusDateEnd = Format(dateBusinessDate, sDateFormatBusEnd)
        If (bGetHHAndMemberOffers = True) Then
            MyCommon.QueryStr = "pa_LogixServ_FetchCustGroups_MemberOrHousehold"
        Else
            MyCommon.QueryStr = "dbo.pa_LogixServ_FetchCustGroups"
        End If
        MyCommon.Open_LXSsp()
        MyCommon.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = lCustomerPK
        MyCommon.LXSsp.Parameters.Add("@HHPK", SqlDbType.BigInt).Value = lHouseholdPK
        WriteDebug("Begin Execution of pa_LogixServ_FetchCustGroups", DebugState.CurrentTime)
        CGdt = MyCommon.LXSsp_select
        WriteDebug("Completed Execution of pa_LogixServ_FetchCustGroups", DebugState.CurrentTime)
        MyCommon.Close_LXSsp()

        cgXml = "<customergroups><id>1</id><id>2</id>"
        If CGdt.Rows.Count > 0 Then
            For Each CGdr In CGdt.Rows
                cgXml &= "<id>" & MyCommon.NZ(CGdr.Item(0), "") & "</id>"
            Next
        End If
        cgXml &= "</customergroups>"
    
        WriteDebug("Customer Groups XML: " & cgXml, DebugState.CurrentTime)
    
        MyCommon.QueryStr = "dbo.pa_CMCustomerWebOffersList"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@cgXml", SqlDbType.Xml).Value = cgXml
        MyCommon.LRTsp.Parameters.Add("@BusDateStart", SqlDbType.NVarChar).Value = sBusDateStart
        MyCommon.LRTsp.Parameters.Add("@BusDateEnd", SqlDbType.NVarChar).Value = sBusDateEnd
    
        WriteDebug("Begin Execution of OffersID query", DebugState.CurrentTime)
        reader = MyCommon.LRTsp.ExecuteReader
        WriteDebug("Completed Execution of OffersID Query", DebugState.CurrentTime)
    
        Try
            dt.Load(reader)
        Catch ex As Exception
            WriteDebug("Exception " & ex.GetType.Name & ":" & ex.Message(), DebugState.CurrentTime)
        End Try
    

        WriteDebug("OfferIDs loaded into data table with row count=" & dt.Rows.Count, DebugState.CurrentTime)
          
        MyCommon.Close_LRTsp()
        reader.Close()
                        
        If dt.Rows.Count > 0 Then
            WriteDebug("Begin Processing OffersIDs", DebugState.CurrentTime)
            For Each dr In dt.Rows
                sOfferId = dr.Item(0)
                bOfferLimitReached = False
                dCustAmount = 0
                MyCommon.QueryStr = "select O.OfferId, O.Name,O.Description,O.OfferCategoryID,O.ProdStartDate,O.ProdEndDate," & _
                                    "O.DistPeriodLimit,O.DistPeriod,O.DistPeriodVarID,O.AdvancedLimitID," & _
                                    "O.NumTiers,O.EmployeeFiltering,O.NonEmployeesOnly,O.DisplayOnWebKiosk,OC.LinkID " & _
                                    "from CM_ST_Offers as O with (NoLock) " & _
                                    "inner join CM_ST_OfferConditions as OC with (NoLock) on OC.OfferID=O.OfferId and OC.ConditionTypeID=1 " & _
                                    "where O.OfferId=" & sOfferId & " ;"

                dtOffer = MyCommon.LRT_Select
                If dtOffer.Rows.Count > 0 Then
                    'Read data and keep it for later comparison
                    If Not IsDBNull(dtOffer.Rows(0).Item("AdvancedLimitID")) Then
                        iAdvancedLimitID = dtOffer.Rows(0).Item("AdvancedLimitID")
                    Else
                        iAdvancedLimitID = 0
                    End If
                    If Not IsDBNull(dtOffer.Rows(0).Item("DistPeriodVarID")) Then
                        iPromoVarID = dtOffer.Rows(0).Item("DistPeriodVarID")
                    Else
                        iPromoVarID = 0
                    End If
                    If Not IsDBNull(dtOffer.Rows(0).Item("DistPeriod")) Then
                        iPeriod = dtOffer.Rows(0).Item("DistPeriod")
                    Else
                        iPeriod = 0
                    End If
                    If Not IsDBNull(dtOffer.Rows(0).Item("DistPeriodLimit")) Then
                        dLimitValue = dtOffer.Rows(0).Item("DistPeriodLimit")
                    Else
                        dLimitValue = 0
                    End If

                    If (iPeriod <> 0) Then ' Limits need to be checked only for Period <>0
                        'We will ignore transaction limits and only consider day and customer based limits. 
                        'Hence O.DistPeriod <> 0 check
                        'Handle Advance Limits
                        If (iAdvancedLimitID > 0) Then 'Get promoVar and limit for the advanced limit

                            MyCommon.QueryStr = "Select Amount FROM  CM_AdvancedLimitVariables WHERE PromoVarID = " & iPromoVarID & " AND " & _
                                                " CustomerPK = " & lCustomerPK & " "
                            dtCustLimit = MyCommon.LXS_Select
                            If dtCustLimit.Rows.Count > 0 Then 'If entry exists  in CM_AdvancedLimitVariables for promovarid and customer
                                dCustAmount = dtCustLimit.Rows(0).Item("Amount")
                            End If
                        Else  '(iAdvancedLimitID < 0)
                            ' If not advance limit, check if it is a regular limit.
                            MyCommon.QueryStr = "Select Amount FROM  DistributionVariables WHERE PromoVarID = " & iPromoVarID & " AND " & _
                                                " CustomerPK = " & lCustomerPK & " "
                            dtCustLimit = MyCommon.LXS_Select
                            If dtCustLimit.Rows.Count > 0 Then 'If entry exists  in DistributionVariables for promovarid and customer
                                dCustAmount = dtCustLimit.Rows(0).Item("Amount")
                            Else : dCustAmount = 0
                            End If
                        End If
                        If (dCustAmount > 0) And (dCustAmount >= dLimitValue) Then
                            bOfferLimitReached = True
                        End If
                    End If '(iPeriod <> 0)

         
                    If (bOfferLimitReached = False) Then
                        drOffer = dtOffers.NewRow()
                        drOffer.Item("OfferID") = sOfferId
                        If Not IsDBNull(dtOffer.Rows(0).Item("Name")) Then
                            drOffer.Item("Name") = dtOffer.Rows(0).Item("Name")
                        End If
                        If Not IsDBNull(dtOffer.Rows(0).Item("Description")) Then
                            drOffer.Item("Description") = dtOffer.Rows(0).Item("Description")
                        End If
                        iCategoryId = MyCommon.NZ(dtOffer.Rows(0).Item("OfferCategoryID"), 0)
                        sCategory = GetCategoryFromId(iCategoryId)
                        If Not IsDBNull(sCategory) Then
                            drOffer.Item("OfferCategory") = sCategory
                        Else : drOffer.Item("OfferCategory") = ""
                        End If
                        If Not IsDBNull(dtOffer.Rows(0).Item("ProdStartDate")) Then
                            drOffer.Item("StartDate") = dtOffer.Rows(0).Item("ProdStartDate")
                        End If
                        If Not IsDBNull(dtOffer.Rows(0).Item("ProdEndDate")) Then
                            drOffer.Item("EndDate") = dtOffer.Rows(0).Item("ProdEndDate")
                        End If
                        If (MyCommon.Fetch_CM_SystemOption(85) = "1") Then
                            MyCommon.QueryStr = "SELECT DisplayStartDate,DisplayEndDate FROM OfferAccessoryFields with (NoLock) WHERE OfferID = " & sOfferId & " "
                            rstDispDates = MyCommon.LRT_Select
                            If rstDispDates.Rows.Count > 0 Then
                                If Not IsDBNull(rstDispDates.Rows(0).Item("DisplayStartDate")) Then
                                    drOffer.Item("DisplayStartDate") = rstDispDates.Rows(0).Item("DisplayStartDate")
                                End If
                                If Not IsDBNull(rstDispDates.Rows(0).Item("DisplayEndDate")) Then
                                    drOffer.Item("DisplayEndDate") = rstDispDates.Rows(0).Item("DisplayEndDate")
                                End If
                            End If
                        End If
                        If Not IsDBNull(dtOffer.Rows(0).Item("LinkID")) Then
                            drOffer.Item("CustomerGroupID") = dtOffer.Rows(0).Item("LinkID")
                        End If
                        If Not IsDBNull(dtOffer.Rows(0).Item("EmployeeFiltering")) Then
                            iFilterEmp = dtOffer.Rows(0).Item("EmployeeFiltering")
                        End If
                        If Not IsDBNull(dtOffer.Rows(0).Item("NonEmployeesOnly")) Then
                            iNonEmpOnly = dtOffer.Rows(0).Item("NonEmployeesOnly")
                        End If
                        If iFilterEmp Then
                            If iNonEmpOnly Then
                                drOffer.Item("EmployeesOnly") = 0
                                drOffer.Item("EmployeesExcluded") = 1
                            Else
                                drOffer.Item("EmployeesOnly") = 1
                                drOffer.Item("EmployeesExcluded") = 0
                            End If
                        Else
                            drOffer.Item("EmployeesOnly") = 0
                            drOffer.Item("EmployeesExcluded") = 0
                        End If
                        dtOffers.Rows.Add(drOffer)
                    End If 'Offer limit reached = false

                End If 'dtOffer.Rows.Count > 0
            Next
            WriteDebug("End Processing OffersIDs", DebugState.CurrentTime)
            If dtOffers.Rows.Count > 0 Then dtOffers.AcceptChanges()

        Else
            bstatus = False
        End If

        Return bstatus
    End Function

    Private Function GetCustomerCpeOffersLimitsBased(ByVal lCustomerPK As Long, ByVal lHouseholdPK As Long, ByRef dtOffers As DataTable, ByVal bGetHHAndMemberOffers As Boolean, ByRef dtudfs As DataTable) As Boolean
        Const sDateFormatBusStart As String = "yyyy-MM-ddT23:59:59"
        Const sDateFormatBusEnd As String = "yyyy-MM-ddT00:00:00"
        Const sDateFormat As String = "yyyy-MM-ddTHH:mm:ss"

        Dim dt As DataTable = Nothing
        Dim dt2 As DataTable = Nothing
        Dim dr As DataRow
        Dim dtOffer As DataTable
        Dim dtUDField As DataTable
        Dim dtCustLimit As DataTable
        Dim drOffer As DataRow
        Dim drOfferudf As DataRow
        Dim sGroupCustomerList As String
        Dim sBusDate As String
        Dim sBusDateStart As String
        Dim sBusDateEnd As String
        Dim dateBusinessDate As Date = Now
        Dim sIncentiveId As String
        Dim i As Integer
        Dim iFilterEmp As Integer
        Dim iNonEmpOnly As Integer
        Dim bstatus As Boolean = True
        Dim iCategoryId As Integer
        Dim sCategory As String
        Dim iAdvancedLimitID As Integer
        Dim bIncentiveLimitReached As Boolean = False
        Dim iPromoVarID As Integer
        Dim dLimitValue As Decimal
        Dim dCustAmount As Decimal = 0
        Dim iPeriod As Integer



        sBusDate = Format(dateBusinessDate, sDateFormat)
        sBusDateStart = "'" & Format(dateBusinessDate, sDateFormatBusStart) & "'"
        sBusDateEnd = "'" & Format(dateBusinessDate, sDateFormatBusEnd) & "'"
        bGetHHAndMemberOffers = False
        If (bGetHHAndMemberOffers = True) Then
            MyCommon.QueryStr = "pa_LogixServ_FetchCustGroups_MemberOrHousehold"
        Else
            MyCommon.QueryStr = "dbo.pa_LogixServ_FetchCustGroups"
        End If
        MyCommon.Open_LXSsp()
        MyCommon.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = lCustomerPK
        MyCommon.LXSsp.Parameters.Add("@HHPK", SqlDbType.BigInt).Value = lHouseholdPK
        dt = MyCommon.LXSsp_select
        MyCommon.Close_LXSsp()

        If dt.Rows.Count > 0 Then
            sGroupCustomerList = "(" & dt.Rows(0).Item(0).ToString
            If dt.Rows.Count > 1 Then
                For i = 1 To dt.Rows.Count - 1
                    sGroupCustomerList += "," & dt.Rows(i).Item(0)
                Next
            End If
            sGroupCustomerList += ",1,2)"
        Else
            ' Member, but no specific groups assigned
            sGroupCustomerList = "(1,2)"
        End If

        'Select all offers that have a customer group that this customer belongs to
        'Do not include offers that exclude a customer group that this customer belongs to
        MyCommon.QueryStr = "select  distinct i.incentiveid " & _
                            "from CPE_ST_Incentives as i with (nolock) " & _
                            "inner join CPE_ST_RewardOptions as r with (nolock) on i.incentiveid=r.incentiveid " & _
                            "inner join CPE_ST_IncentiveCustomerGroups CG1 with (nolock) on CG1.rewardoptionid=r.rewardoptionid " & _
                            "where i.EngineId = 2 and r.Deleted = 0 and CG1.Deleted = 0 " & _
                            "and i.StartDate <= " & sBusDateStart & " and i.EndDate >=  " & sBusDateEnd & _
                            "and ((CG1.CustomerGroupID in " & sGroupCustomerList & " or CG1.CustomerGroupID in (1,2)) and i.IncentiveID not IN ( " & _
                            "select i.IncentiveID from CPE_ST_Incentives i with (nolock) " & _
                            "inner join CPE_ST_RewardOptions as r with (nolock) on i.incentiveid=r.incentiveid " & _
                            "inner join CPE_ST_IncentiveCustomerGroups CG1 with (nolock) on CG1.rewardoptionid=r.rewardoptionid " & _
                            "where i.EngineId = 2 and r.Deleted = 0 and CG1.Deleted = 0 " & _
                            "and i.StartDate <= " & sBusDateEnd & " and i.EndDate >= " & sBusDateEnd & " and " & _
                            "(CG1.CustomerGroupID in " & sGroupCustomerList & " or CG1.CustomerGroupID in (1,2)) and ExcludedUsers = 1));"

        dt = MyCommon.LRT_Select
        If dt.Rows.Count > 0 Then
            For Each dr In dt.Rows
                dtudfs.Clear()
                sIncentiveId = dr.Item(0)
                bIncentiveLimitReached = False
                'gather the necessary information to determine if the customer has exceeded their limit
                MyCommon.QueryStr = "select i.incentiveid, i.StartDate, i.EndDate, i.IncentiveName, i.Description, i.employeesonly, " & _
                                    "i.employeesexcluded, i.P3DistPeriod, i.P3DistTimeType, i.P3DistQtyLimit, CG1.CustomerGroupID " & _
                                    "from CPE_ST_Incentives I with (nolock) " & _
                                    "inner join CPE_ST_RewardOptions as r with (nolock) on i.incentiveid=r.incentiveid " & _
                                    "inner join CPE_ST_IncentiveCustomerGroups CG1 with (nolock) on CG1.rewardoptionid=r.rewardoptionid " & _
                                    "where i.incentiveid = " & sIncentiveId & " and CG1.ExcludedUsers <> 1;"

                dtOffer = MyCommon.LRT_Select
                If dtOffer.Rows.Count > 0 Then
                    'see if a customer has exceeded their limit for this Incentive
                    ' Limits need to be checked only if Period and limit <> 0
                    If ((dtOffer.Rows(0).Item("P3DistPeriod") <> 0) Or (dtOffer.Rows(0).Item("P3DistQtyLimit") <> 0)) Then
                        Dim NumberRedemptions As Integer = 0
                        Dim NumberHours As Integer = 0
                        MyCommon.QueryStr = "select count(RD.DistributionID) as NumberRedemptions" & _
                                               " from CPE_RewardDistribution rd" & _
                                               " where rd.DistributionDate >= DATEADD(day, -" & dtOffer.Rows(0).Item("P3DistPeriod") & ", convert(datetime, GETDATE()))" & _
                                               " and rd.DistributionDate < convert(datetime, GETDATE())" & _
                                               " and rd.IncentiveID=" & dtOffer.Rows(0).Item("IncentiveID") & _
                                               " and rd.CustomerPK=" & lCustomerPK & ";"
                        dt2 = MyCommon.LXS_Select
                        NumberRedemptions = dt2.Rows(0).Item("NumberRedemptions")

                        MyCommon.QueryStr = "select DATEDIFF(day, MAX(rd.DistributionDate), getdate()) as NumberHours" & _
                                               " from CPE_RewardDistribution RD" & _
                                               " where rd.IncentiveID=" & dtOffer.Rows(0).Item("IncentiveID") & _
                                               " and rd.CustomerPK=" & lCustomerPK & ";"
                        dt2 = MyCommon.LXS_Select
                        NumberHours =  Copient.commonShared.NZ(dt2.Rows(0).Item("NumberHours"), -1)
              
                        MyCommon.QueryStr = "dbo.pa_CPE_TestLimits"
                        MyCommon.Open_LRTsp()
                        MyCommon.LRTsp.Parameters.Add("@incentiveid", SqlDbType.BigInt).Value = dtOffer.Rows(0).Item("IncentiveID")
                        MyCommon.LRTsp.Parameters.Add("@NumberDays", SqlDbType.Int).Value = dtOffer.Rows(0).Item("P3DistPeriod")
                        MyCommon.LRTsp.Parameters.Add("@NumberHours", SqlDbType.Int).Value = NumberHours
                        MyCommon.LRTsp.Parameters.Add("@LimitNumber", SqlDbType.Int).Value = dtOffer.Rows(0).Item("P3DistQtyLimit")
                        MyCommon.LRTsp.Parameters.Add("@TimeType", SqlDbType.Int).Value = dtOffer.Rows(0).Item("P3DistTimeType")
                        MyCommon.LRTsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = lCustomerPK
                        MyCommon.LRTsp.Parameters.Add("@NumberRedemptions", SqlDbType.Int).Value = NumberRedemptions
                        MyCommon.LRTsp.Parameters.Add("@limitexceeded", SqlDbType.Bit).Direction = ParameterDirection.Output
                        MyCommon.LRTsp.ExecuteNonQuery()
                        If MyCommon.LRTsp.Parameters("@limitexceeded").Value = True Then
                            bIncentiveLimitReached = True
                        End If
                        MyCommon.Close_LRTsp()
                    End If '(iPeriod <> 0)

                    If (bIncentiveLimitReached = False) Then
                        'send back information for each incentive where the customer has not exceeded their limit
                       
                        Dim Tempdtudfs As New DataTable
                        'Enable the userdefinedfields for an offer '
                        If MyCommon.Fetch_SystemOption(156) = "1" AndAlso MyCommon.Fetch_SystemOption(293) = "1" Then
                            'Retrieving  userdefinedfields associated  to the offer
                            MyCommon.QueryStr = "select udf.ExternalID, udf.Description, UDFV.StringValue, UDFV.IntValue, " & _
                                  " UDFV.DateValue, UDFV.BooleanValue from UserDefinedFieldsValues UDFV inner join userdefinedfields UDF on udfv.UDFPK=udf.UDFPK " & _
                                  " inner join UserDefinedFieldsTypes UDFT on  udf.DataType=udft.UDFTypeID where OfferID =" & sIncentiveId & ""
                            dtUDField = MyCommon.LRT_Select()
                            If dtUDField.Rows.Count > 0 Then
                           
                                For Each dr1 As DataRow In dtUDField.Rows
                                
                                    drOfferudf = dtudfs.NewRow()
                                    If Not IsDBNull(dr1.Item("StringValue")) Then
                                        drOfferudf.Item("ExternalID") = dr1.Item("ExternalID")
                                        drOfferudf.Item("UserDescription") = dr1.Item("Description")
                                        drOfferudf.Item("Value") = dr1.Item("StringValue")
                                    End If
                        
                                    If Not IsDBNull(dr1.Item("IntValue")) Then
                                        drOfferudf.Item("ExternalID") = dr1.Item("ExternalID")
                                        drOfferudf.Item("UserDescription") = dr1.Item("Description")
                                        drOfferudf.Item("Value") = dr1.Item("IntValue")
                                    End If
                        
                                    If Not IsDBNull(dr1.Item("DateValue")) Then
                                        drOfferudf.Item("ExternalID") = dr1.Item("ExternalID")
                                        drOfferudf.Item("UserDescription") = dr1.Item("Description")
                                        drOfferudf.Item("Value") = dr1.Item("DateValue")
                                    End If
                        
                                    If Not IsDBNull(dr1.Item("BooleanValue")) Then
                                        drOfferudf.Item("ExternalID") = dr1.Item("ExternalID")
                                        drOfferudf.Item("UserDescription") = dr1.Item("Description")
                                        drOfferudf.Item("Value") = dr1.Item("BooleanValue")
                                    End If
                                    'Added userdefined fields to dtudfs'
                                    dtudfs.Rows.Add(drOfferudf)
                                    Tempdtudfs = dtudfs.Copy()
                                Next
                                'Userdefinedfields is empty
                            ElseIf dtUDField.Rows.Count = 0 Then
                                drOfferudf = dtudfs.NewRow()
                                drOfferudf.Item("ExternalID") = ""
                                drOfferudf.Item("UserDescription") = ""
                                drOfferudf.Item("Value") = ""
                                dtudfs.Rows.Add(drOfferudf)
                                Tempdtudfs = dtudfs.Copy()
                            End If
           
                        End If

                        
                        drOffer = dtOffers.NewRow()
                        drOffer.Item("IncentiveID") = sIncentiveId
                        If Not IsDBNull(dtOffer.Rows(0).Item("IncentiveName")) Then
                            drOffer.Item("IncentiveName") = dtOffer.Rows(0).Item("IncentiveName")
                        End If
                        If Not IsDBNull(dtOffer.Rows(0).Item("Description")) Then
                            drOffer.Item("Description") = dtOffer.Rows(0).Item("Description")
                        End If
                        If Not IsDBNull(dtOffer.Rows(0).Item("StartDate")) Then
                            drOffer.Item("StartDate") = dtOffer.Rows(0).Item("StartDate")
                        End If
                        If Not IsDBNull(dtOffer.Rows(0).Item("EndDate")) Then
                            drOffer.Item("EndDate") = dtOffer.Rows(0).Item("EndDate")
                        End If
                        If Not IsDBNull(dtOffer.Rows(0).Item("CustomerGroupID")) Then
                            drOffer.Item("CustomerGroupID") = dtOffer.Rows(0).Item("CustomerGroupID")
                        End If
                        If Not IsDBNull(dtOffer.Rows(0).Item("EmployeesOnly")) Then
                            drOffer.Item("EmployeesOnly") = dtOffer.Rows(0).Item("EmployeesOnly")
                        End If
                        If Not IsDBNull(dtOffer.Rows(0).Item("EmployeesExcluded")) Then
                            drOffer.Item("EmployeesExcluded") = dtOffer.Rows(0).Item("EmployeesExcluded")
                        End If
                        
                        If Tempdtudfs.Rows.Count > 0 Then
                            
                            drOffer.Item("UserDefinedFields") = Tempdtudfs
                        End If
                        
                        dtOffers.Rows.Add(drOffer)
                    End If 'Offer limit reached = false

                End If 'dtOffer.Rows.Count > 0
            Next
            If dtOffers.Rows.Count > 0 Then dtOffers.AcceptChanges()
        Else
            bstatus = False
        End If

        Return bstatus

    End Function
  
    Private Function GetProgramsForCMOffer(ByVal sOfferId As String, ByRef dtPrograms As DataTable) As Boolean
        Dim dt As DataTable = Nothing
        Dim dr As DataRow
        Dim drProgram As DataRow
        Dim bstatus As Boolean = True
        Dim iCategoryId As Integer
        Dim sCategory As String

        MyCommon.QueryStr = "dbo.pa_CustWebCM_OfferProgramsGet"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = sOfferId
        dt = MyCommon.LRTsp_select
        MyCommon.Close_LRTsp()

        If dt.Rows.Count > 0 Then
            For Each dr In dt.Rows
                drProgram = dtPrograms.NewRow()
                drProgram.Item("OfferID") = dr.Item("OfferId")
                drProgram.Item("ConditionOrReward") = dr.Item("ConditionOrReward")
                drProgram.Item("ProgramType") = dr.Item("ProgramType")
                drProgram.Item("ProgramID") = dr.Item("ProgramID")
                drProgram.Item("TierLevel") = dr.Item("TierLevel")
                drProgram.Item("Amount") = dr.Item("Amount")
                drProgram.Item("Name") = dr.Item("Name")
                drProgram.Item("Description") = dr.Item("Description")
                iCategoryId = MyCommon.NZ(dr.Item("CategoryID"), 0)
                sCategory = GetCategoryFromId(iCategoryId)
                drProgram.Item("Category") = sCategory
                drProgram.Item("WebMessage") = dr.Item("WebMessage")

                dtPrograms.Rows.Add(drProgram)
            Next
        End If

        Return bstatus
    End Function

    Private Function GetCustomerExternalCmOffers(ByVal lCustomerPK As Long, ByVal lHouseholdPK As Long, ByRef dtOffers As DataTable, ByRef dtPrograms As DataTable, ByVal bIncludeRedeemed As Boolean, ByVal bIncludeUnRedeemed As Boolean) As Boolean

        Const sDateFormatBusStart As String = "yyyy-MM-ddT23:59:59"
        Const sDateFormatBusEnd As String = "yyyy-MM-ddT00:00:00"
        Const sDateFormat As String = "yyyy-MM-ddTHH:mm:ss"

        Dim dt As DataTable = Nothing
        Dim dr As DataRow
        Dim dtOffer As DataTable
        Dim dtCustLimit As DataTable
        Dim drOffer As DataRow
        Dim sGroupCustomerList As String
        Dim sBusDate As String
        Dim sBusDateStart As String
        Dim sBusDateEnd As String
        Dim dateBusinessDate As Date = Now
        Dim sOfferId As String
        Dim i As Integer
        Dim iFilterEmp As Integer
        Dim iNonEmpOnly As Integer
        Dim bstatus As Boolean = True
        Dim iCategoryId As Integer
        Dim sCategory As String
        Dim iAdvancedLimitID As Integer
        Dim bOfferLimitReached As Boolean = False
        Dim iPromoVarID As Integer
        Dim dLimitValue As Decimal
        Dim dCustAmount As Decimal = 0
        Dim iPeriod As Integer

        sBusDate = Format(dateBusinessDate, sDateFormat)
        sBusDateStart = "'" & Format(dateBusinessDate, sDateFormatBusStart) & "'"
        sBusDateEnd = "'" & Format(dateBusinessDate, sDateFormatBusEnd) & "'"

        MyCommon.QueryStr = "dbo.pa_LogixServ_FetchCustGroups"
        MyCommon.Open_LXSsp()
        MyCommon.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = lCustomerPK
        MyCommon.LXSsp.Parameters.Add("@HHPK", SqlDbType.BigInt).Value = lHouseholdPK
        dt = MyCommon.LXSsp_select
        MyCommon.Close_LXSsp()

        If dt.Rows.Count > 0 Then
            sGroupCustomerList = "(" & dt.Rows(0).Item(0).ToString
            If dt.Rows.Count > 1 Then
                For i = 1 To dt.Rows.Count - 1
                    sGroupCustomerList += "," & dt.Rows(i).Item(0)
                Next
            End If
            sGroupCustomerList += ")"
        Else
            ' Member, but no specific groups assigned
            sGroupCustomerList = "()"
        End If

        MyCommon.QueryStr = "select distinct OfferId from CM_ST_OfferCustLocView with (NoLock) where" & _
                            " ProdStartDate <= " & sBusDateStart & " and ProdEndDate >= " & sBusDateEnd & _
                            " and ((AnyCustomer = 0 and AnyCardholder = 0 and InCustGroupId in " & sGroupCustomerList & _
                            " and (ExCustGroupId is null or ExCustGroupId not in " & sGroupCustomerList & "))" & _
                            " or (AnyCardholder = 1 and (ExCustGroupId is null or ExCustGroupId not in " & sGroupCustomerList & "))) order by OfferId;"
        dt = MyCommon.LRT_Select
        If dt.Rows.Count > 0 Then
            For Each dr In dt.Rows
                sOfferId = dr.Item(0)


                MyCommon.QueryStr = "select O.OfferId, O.Name,O.Description,O.OfferCategoryID,O.ProdStartDate,O.ProdEndDate," & _
                                    "O.DistPeriodLimit,O.DistPeriod,O.DistPeriodVarID,O.AdvancedLimitID," & _
                                    "O.NumTiers,O.EmployeeFiltering,O.NonEmployeesOnly,O.DisplayOnWebKiosk,OC.LinkID " & _
                                    "from CM_ST_Offers as O with (NoLock) " & _
                                    "inner join CM_ST_OfferConditions as OC with (NoLock) on OC.OfferID=O.OfferId and OC.ConditionTypeID=1 " & _
                                    "where O.OfferId=" & sOfferId & " AND O.ExtOfferID is not null;"

                dtOffer = MyCommon.LRT_Select
                If dtOffer.Rows.Count > 0 Then
                    'Read data and keep it for later comparison
                    If Not IsDBNull(dtOffer.Rows(0).Item("AdvancedLimitID")) Then
                        iAdvancedLimitID = dtOffer.Rows(0).Item("AdvancedLimitID")
                    Else
                        iAdvancedLimitID = 0
                    End If
                    If Not IsDBNull(dtOffer.Rows(0).Item("DistPeriodVarID")) Then
                        iPromoVarID = dtOffer.Rows(0).Item("DistPeriodVarID")
                    Else
                        iPromoVarID = 0
                    End If
                    If Not IsDBNull(dtOffer.Rows(0).Item("DistPeriod")) Then
                        iPeriod = dtOffer.Rows(0).Item("DistPeriod")
                    Else
                        iPeriod = 0
                    End If
                    If Not IsDBNull(dtOffer.Rows(0).Item("DistPeriodLimit")) Then
                        dLimitValue = dtOffer.Rows(0).Item("DistPeriodLimit")
                    Else
                        dLimitValue = 0
                    End If

                    If (iPeriod <> 0) Then ' Limits need to be checked only for Period <>0
                        'We will ignore transaction limits and only consider day and customer based limits. 
                        'Hence O.DistPeriod <> 0 check
                        'Handle Advance Limits
                        If (iAdvancedLimitID > 0) Then 'Get promoVar and limit for the advanced limit

                            MyCommon.QueryStr = "Select Amount FROM  CM_AdvancedLimitVariables WHERE PromoVarID = " & iPromoVarID & " AND " & _
                                                " CustomerPK = " & lCustomerPK & " "
                            dtCustLimit = MyCommon.LXS_Select
                            If dtCustLimit.Rows.Count > 0 Then 'If entry exists  in CM_AdvancedLimitVariables for promovarid and customer
                                dCustAmount = dtCustLimit.Rows(0).Item("Amount")
                            Else
                                dCustAmount = 0
                            End If
                        Else  '(iAdvancedLimitID <= 0)
                            ' If not advance limit, it is a regular limit.
                            MyCommon.QueryStr = "Select Amount FROM  DistributionVariables WHERE PromoVarID = " & iPromoVarID & " AND " & _
                                                " CustomerPK = " & lCustomerPK & " "
                            dtCustLimit = MyCommon.LXS_Select
                            If dtCustLimit.Rows.Count > 0 Then 'If entry exists  in DistributionVariables for promovarid and customer
                                dCustAmount = dtCustLimit.Rows(0).Item("Amount")
                            Else
                                dCustAmount = 0
                            End If
                        End If


                        If (dCustAmount > 0) And (dCustAmount >= dLimitValue) Then
                            bOfferLimitReached = True
                            If (bIncludeRedeemed = True) Then
                                drOffer = dtOffers.NewRow()
                                drOffer.Item("OfferID") = sOfferId
                                If Not IsDBNull(dtOffer.Rows(0).Item("Name")) Then
                                    drOffer.Item("Name") = dtOffer.Rows(0).Item("Name")
                                End If
                                If Not IsDBNull(dtOffer.Rows(0).Item("Description")) Then
                                    drOffer.Item("Description") = dtOffer.Rows(0).Item("Description")
                                End If
                                iCategoryId = MyCommon.NZ(dtOffer.Rows(0).Item("OfferCategoryID"), 0)
                                sCategory = GetCategoryFromId(iCategoryId)
                                If Not IsDBNull(sCategory) Then
                                    drOffer.Item("OfferCategory") = sCategory
                                Else : drOffer.Item("OfferCategory") = ""
                                End If
                                If Not IsDBNull(dtOffer.Rows(0).Item("ProdStartDate")) Then
                                    drOffer.Item("StartDate") = dtOffer.Rows(0).Item("ProdStartDate")
                                End If
                                If Not IsDBNull(dtOffer.Rows(0).Item("ProdEndDate")) Then
                                    drOffer.Item("EndDate") = dtOffer.Rows(0).Item("ProdEndDate")
                                End If
                                If Not IsDBNull(dtOffer.Rows(0).Item("LinkID")) Then
                                    drOffer.Item("CustomerGroupID") = dtOffer.Rows(0).Item("LinkID")
                                End If
                                If Not IsDBNull(dtOffer.Rows(0).Item("EmployeeFiltering")) Then
                                    iFilterEmp = dtOffer.Rows(0).Item("EmployeeFiltering")
                                End If
                                If Not IsDBNull(dtOffer.Rows(0).Item("NonEmployeesOnly")) Then
                                    iNonEmpOnly = dtOffer.Rows(0).Item("NonEmployeesOnly")
                                End If
                                If iFilterEmp Then
                                    If iNonEmpOnly Then
                                        drOffer.Item("EmployeesOnly") = 0
                                        drOffer.Item("EmployeesExcluded") = 1
                                    Else
                                        drOffer.Item("EmployeesOnly") = 1
                                        drOffer.Item("EmployeesExcluded") = 0
                                    End If
                                Else
                                    drOffer.Item("EmployeesOnly") = 0
                                    drOffer.Item("EmployeesExcluded") = 0
                                End If
                                dtOffers.Rows.Add(drOffer)
                                'Get program information for redeemed offers
                                bstatus = GetProgramsForCMOffer(sOfferId, dtPrograms)
                            End If  'IncludeReemed

                        Else
                            bOfferLimitReached = False
                            If (bIncludeUnRedeemed = True) Then
                                drOffer = dtOffers.NewRow()
                                drOffer.Item("OfferID") = sOfferId
                                If Not IsDBNull(dtOffer.Rows(0).Item("Name")) Then
                                    drOffer.Item("Name") = dtOffer.Rows(0).Item("Name")
                                End If
                                If Not IsDBNull(dtOffer.Rows(0).Item("Description")) Then
                                    drOffer.Item("Description") = dtOffer.Rows(0).Item("Description")
                                End If
                                iCategoryId = MyCommon.NZ(dtOffer.Rows(0).Item("OfferCategoryID"), 0)
                                sCategory = GetCategoryFromId(iCategoryId)
                                If Not IsDBNull(sCategory) Then
                                    drOffer.Item("OfferCategory") = sCategory
                                Else : drOffer.Item("OfferCategory") = ""
                                End If
                                If Not IsDBNull(dtOffer.Rows(0).Item("ProdStartDate")) Then
                                    drOffer.Item("StartDate") = dtOffer.Rows(0).Item("ProdStartDate")
                                End If
                                If Not IsDBNull(dtOffer.Rows(0).Item("ProdEndDate")) Then
                                    drOffer.Item("EndDate") = dtOffer.Rows(0).Item("ProdEndDate")
                                End If
                                If Not IsDBNull(dtOffer.Rows(0).Item("LinkID")) Then
                                    drOffer.Item("CustomerGroupID") = dtOffer.Rows(0).Item("LinkID")
                                End If
                                If Not IsDBNull(dtOffer.Rows(0).Item("EmployeeFiltering")) Then
                                    iFilterEmp = dtOffer.Rows(0).Item("EmployeeFiltering")
                                End If
                                If Not IsDBNull(dtOffer.Rows(0).Item("NonEmployeesOnly")) Then
                                    iNonEmpOnly = dtOffer.Rows(0).Item("NonEmployeesOnly")
                                End If
                                If iFilterEmp Then
                                    If iNonEmpOnly Then
                                        drOffer.Item("EmployeesOnly") = 0
                                        drOffer.Item("EmployeesExcluded") = 1
                                    Else
                                        drOffer.Item("EmployeesOnly") = 1
                                        drOffer.Item("EmployeesExcluded") = 0
                                    End If
                                Else
                                    drOffer.Item("EmployeesOnly") = 0
                                    drOffer.Item("EmployeesExcluded") = 0
                                End If
                                dtOffers.Rows.Add(drOffer)
                                bstatus = GetProgramsForCMOffer(sOfferId, dtPrograms)
                            End If ' bIncludeUnRedeemed = true

                        End If
                    Else
                        drOffer = dtOffers.NewRow()
                        drOffer.Item("OfferID") = sOfferId
                        If Not IsDBNull(dtOffer.Rows(0).Item("Name")) Then
                            drOffer.Item("Name") = dtOffer.Rows(0).Item("Name")
                        End If
                        If Not IsDBNull(dtOffer.Rows(0).Item("Description")) Then
                            drOffer.Item("Description") = dtOffer.Rows(0).Item("Description")
                        End If
                        iCategoryId = MyCommon.NZ(dtOffer.Rows(0).Item("OfferCategoryID"), 0)
                        sCategory = GetCategoryFromId(iCategoryId)
                        If Not IsDBNull(sCategory) Then
                            drOffer.Item("OfferCategory") = sCategory
                        Else : drOffer.Item("OfferCategory") = ""
                        End If
                        If Not IsDBNull(dtOffer.Rows(0).Item("ProdStartDate")) Then
                            drOffer.Item("StartDate") = dtOffer.Rows(0).Item("ProdStartDate")
                        End If
                        If Not IsDBNull(dtOffer.Rows(0).Item("ProdEndDate")) Then
                            drOffer.Item("EndDate") = dtOffer.Rows(0).Item("ProdEndDate")
                        End If
                        If Not IsDBNull(dtOffer.Rows(0).Item("LinkID")) Then
                            drOffer.Item("CustomerGroupID") = dtOffer.Rows(0).Item("LinkID")
                        End If
                        If Not IsDBNull(dtOffer.Rows(0).Item("EmployeeFiltering")) Then
                            iFilterEmp = dtOffer.Rows(0).Item("EmployeeFiltering")
                        End If
                        If Not IsDBNull(dtOffer.Rows(0).Item("NonEmployeesOnly")) Then
                            iNonEmpOnly = dtOffer.Rows(0).Item("NonEmployeesOnly")
                        End If
                        If iFilterEmp Then
                            If iNonEmpOnly Then
                                drOffer.Item("EmployeesOnly") = 0
                                drOffer.Item("EmployeesExcluded") = 1
                            Else
                                drOffer.Item("EmployeesOnly") = 1
                                drOffer.Item("EmployeesExcluded") = 0
                            End If
                        Else
                            drOffer.Item("EmployeesOnly") = 0
                            drOffer.Item("EmployeesExcluded") = 0
                        End If
                        dtOffers.Rows.Add(drOffer)
                        bstatus = GetProgramsForCMOffer(sOfferId, dtPrograms)
                    End If '(iPeriod <> 0)

                End If 'dtOffer.Rows.Count > 0
            Next

            If dtOffers.Rows.Count > 0 Then dtOffers.AcceptChanges()
            If dtPrograms.Rows.Count > 0 Then dtPrograms.AcceptChanges()

        Else
            bstatus = False
        End If

        Return bstatus


    End Function

    Private Function GetCustomerExternalExpiredCmOffers(ByVal GUID As String, ByVal ExtCardID As String, ByVal CardTypeID As Integer) As System.Data.DataSet
        Const sDateFormatBusStart As String = "yyyy-MM-ddT23:59:59"
        Const sDateFormatBusEnd As String = "yyyy-MM-ddT00:00:00"
        Const sDateFormat As String = "yyyy-MM-ddTHH:mm:ss"

        Dim dt As System.Data.DataTable
        Dim dr As DataRow
        Dim dtOffer As DataTable
        Dim drOffer As DataRow
        Dim sGroupCustomerList As String
        Dim sBusDate As String
        Dim sBusDateStart As String
        Dim sBusDateEnd As String
        Dim dateBusinessDate As Date = Now
        Dim sOfferId As String
        Dim i As Integer
        Dim iFilterEmp As Integer
        Dim iNonEmpOnly As Integer
        Dim bstatus As Boolean = True
        Dim iCategoryId As Integer
        Dim sCategory As String
        Dim dtOffers As System.Data.DataTable
        Dim dtPrograms As System.Data.DataTable
        Dim ResultSet As New System.Data.DataSet("OfferListCM")
        Dim dtStatus As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim lCustomerPK As Long
        Dim lHouseholdPK As Long
        'added
        Dim RetCode As StatusCodes
        Dim RetMsg As String = ""
        'Dim MyCryptLib As New Copient.CryptLib()


        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            If Not IsValidGUID(GUID) Then
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
              
            ElseIf Not IsValidCustomerCard(ExtCardID, CardTypeID, RetCode, RetMsg) Then
                
                'Bad customer ID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = RetCode
                row.Item("Description") = RetMsg
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
          
                
            Else 'Valid card id length and GUID
                'Pad the card ID
                ExtCardID = MyCommon.Pad_ExtCardID(ExtCardID, CardTypeID)

                'Find the Customer PK



                MyCommon.QueryStr = "select CID.CustomerPK, C.HHPK from CardIDs as CID with (NoLock) " & _
                                    "inner join Customers as C with (NoLock) on C.CustomerPK=CID.CustomerPK " & _
                                    "where CID.ExtCardID='" & MyCryptLib.SQL_StringEncrypt(ExtCardID, True) & "' and CID.CardTypeID = " & CardTypeID & "; "

                dt = MyCommon.LXS_Select
                If dt.Rows.Count = 0 Then
                    'Card not found
                    If CardTypeID = 0 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_CUSTOMER
                        row.Item("Description") = "Failure: Customer " & ExtCardID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    ElseIf CardTypeID = 1 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_HOUSEHOLD
                        row.Item("Description") = "Failure: Household " & ExtCardID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    ElseIf CardTypeID = 2 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_CAM
                        row.Item("Description") = "Failure: CAM " & ExtCardID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    End If
                Else 'Card Record Found
                    lCustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)
                    lHouseholdPK = MyCommon.NZ(dt.Rows(0).Item("HHPK"), 0)
                    'Put a success into the Status table
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = 0
                    row.Item("Description") = "Success."
                    dtStatus.Rows.Add(row)

                    'sBusDate = Format(dateBusinessDate, sDateFormat)
                    'sBusDateStart = "'" & Format(dateBusinessDate, sDateFormatBusStart) & "'"
                    'sBusDateEnd = "'" & Format(dateBusinessDate, sDateFormatBusEnd) & "'"
                    sBusDate = dateBusinessDate.ToString(sDateFormat)
                    sBusDateStart = "'" & dateBusinessDate.ToString(sDateFormatBusStart) & "'"
                    sBusDateEnd = "'" & dateBusinessDate.ToString(sDateFormatBusEnd) & "'"

                    dtOffers = New DataTable
                    dtOffers.TableName = "Offers"
                    dtOffers.Columns.Add("OfferID", System.Type.GetType("System.Int64"))
                    dtOffers.Columns.Add("Name", System.Type.GetType("System.String"))
                    dtOffers.Columns.Add("Description", System.Type.GetType("System.String"))
                    dtOffers.Columns.Add("OfferCategory", System.Type.GetType("System.String"))
                    dtOffers.Columns.Add("StartDate", System.Type.GetType("System.DateTime"))
                    dtOffers.Columns.Add("EndDate", System.Type.GetType("System.DateTime"))
                    dtOffers.Columns.Add("CustomerGroupID", System.Type.GetType("System.Int64"))
                    dtOffers.Columns.Add("EmployeesOnly", System.Type.GetType("System.Boolean"))
                    dtOffers.Columns.Add("EmployeesExcluded", System.Type.GetType("System.Boolean"))

                    dtPrograms = New DataTable
                    dtPrograms.TableName = "Programs"
                    dtPrograms.Columns.Add("OfferID", System.Type.GetType("System.Int64"))
                    dtPrograms.Columns.Add("ConditionOrReward", System.Type.GetType("System.String"))
                    dtPrograms.Columns.Add("ProgramType", System.Type.GetType("System.String"))
                    dtPrograms.Columns.Add("ProgramID", System.Type.GetType("System.Int64"))
                    dtPrograms.Columns.Add("TierLevel", System.Type.GetType("System.Int32"))
                    dtPrograms.Columns.Add("Amount", System.Type.GetType("System.Decimal"))
                    dtPrograms.Columns.Add("Name", System.Type.GetType("System.String"))
                    dtPrograms.Columns.Add("Description", System.Type.GetType("System.String"))
                    dtPrograms.Columns.Add("Category", System.Type.GetType("System.String"))
                    dtPrograms.Columns.Add("WebMessage", System.Type.GetType("System.String"))


                    MyCommon.QueryStr = "dbo.pa_LogixServ_FetchCustGroups"
                    MyCommon.Open_LXSsp()
                    MyCommon.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = lCustomerPK
                    MyCommon.LXSsp.Parameters.Add("@HHPK", SqlDbType.BigInt).Value = lHouseholdPK
                    dt = MyCommon.LXSsp_select
                    MyCommon.Close_LXSsp()
                    If dt.Rows.Count > 0 Then
                        sGroupCustomerList = "(" & dt.Rows(0).Item(0).ToString
                        If dt.Rows.Count > 1 Then
                            For i = 1 To dt.Rows.Count - 1
                                sGroupCustomerList += "," & dt.Rows(i).Item(0)
                            Next
                        End If
                        sGroupCustomerList += ")"
                    Else
                        ' Member, but no specific groups assigned
                        sGroupCustomerList = "()"
                    End If

                    MyCommon.QueryStr = "select distinct OfferId from CM_ST_OfferCustLocView with (NoLock) where" & _
                                  " ProdEndDate < " & sBusDateEnd & _
                                  " and ((AnyCustomer = 0 and AnyCardholder = 0 and InCustGroupId in " & sGroupCustomerList & _
                                  " and (ExCustGroupId is null or ExCustGroupId not in " & sGroupCustomerList & "))" & _
                                  " or (AnyCardholder = 1 and (ExCustGroupId is null or ExCustGroupId not in " & sGroupCustomerList & "))) order by OfferId;"
                    dt = MyCommon.LRT_Select
                    If dt.Rows.Count > 0 Then 'Offer found
                        For Each dr In dt.Rows
                            sOfferId = dr.Item(0)
                            MyCommon.QueryStr = "select O.OfferId, O.Name,O.Description,O.OfferCategoryID,O.ProdStartDate,O.ProdEndDate," & _
                                         "O.NumTiers,O.EmployeeFiltering,O.NonEmployeesOnly,O.DisplayOnWebKiosk,OC.LinkID " & _
                                         "from CM_ST_Offers as O with (NoLock) " & _
                                         "inner join CM_ST_OfferConditions as OC with (NoLock) on OC.OfferID=O.OfferId and OC.ConditionTypeID=1 " & _
                                         "where O.OfferId=" & sOfferId & " AND O.ExtOfferID is not null ;"


                            dtOffer = MyCommon.LRT_Select
                            If dtOffer.Rows.Count > 0 Then
                                drOffer = dtOffers.NewRow()
                                drOffer.Item("OfferID") = sOfferId
                                drOffer.Item("Name") = dtOffer.Rows(0).Item("Name")
                                drOffer.Item("Description") = dtOffer.Rows(0).Item("Description")
                                iCategoryId = MyCommon.NZ(dtOffer.Rows(0).Item("OfferCategoryID"), 0)
                                sCategory = GetCategoryFromId(iCategoryId)
                                drOffer.Item("OfferCategory") = sCategory
                                drOffer.Item("StartDate") = dtOffer.Rows(0).Item("ProdStartDate")
                                drOffer.Item("EndDate") = dtOffer.Rows(0).Item("ProdEndDate")
                                drOffer.Item("CustomerGroupID") = dtOffer.Rows(0).Item("LinkID")

                                iFilterEmp = dtOffer.Rows(0).Item("EmployeeFiltering")
                                iNonEmpOnly = dtOffer.Rows(0).Item("NonEmployeesOnly")
                                If iFilterEmp Then
                                    If iNonEmpOnly Then
                                        drOffer.Item("EmployeesOnly") = 0
                                        drOffer.Item("EmployeesExcluded") = 1
                                    Else
                                        drOffer.Item("EmployeesOnly") = 1
                                        drOffer.Item("EmployeesExcluded") = 0
                                    End If
                                Else
                                    drOffer.Item("EmployeesOnly") = 0
                                    drOffer.Item("EmployeesExcluded") = 0
                                End If
                                dtOffers.Rows.Add(drOffer)
                                bstatus = GetProgramsForCMOffer(sOfferId, dtPrograms)
                            End If 'dtOffer.Rows.count >0
                        Next

                    Else 'Offer not found
                        bstatus = False
                    End If
                    If dtOffers.Rows.Count > 0 Then dtOffers.AcceptChanges()
                    If dtPrograms.Rows.Count > 0 Then dtPrograms.AcceptChanges()
                    ResultSet.Tables.Add(dtStatus.Copy())
                    ResultSet.Tables.Add(dtOffers.Copy())
                    ResultSet.Tables.Add(dtPrograms.Copy())
                End If 'Card record found
            End If 'Valid GUI and card length

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

    Private Function GetCustomerTransactions(ByVal lCustomerPK As Long, ByVal lHouseholdPK As Long, ByVal sCustomerID As String, ByVal CustomerTypeID As Integer, ByVal StartDate As Date, ByVal EndDate As Date, ByRef dtTransactions As DataTable, ByRef dtPoints As DataTable, ByRef dtStoredValues As DataTable) As Boolean
        Const sDateFormat As String = "yyyy-MM-dd"

        Dim dtWH As DataTable = Nothing
        Dim dtRT As DataTable = Nothing
        Dim dtXS As DataTable = Nothing
        Dim dt As DataTable = Nothing
        Dim drWH As DataRow
        Dim drXS As DataRow
        Dim drTransaction As DataRow
        Dim drStoredValues As DataRow
        Dim drPoints As DataRow
        Dim sLocationCode As String
        Dim sLocationName As String
        Dim sHouseHoldID As String
        Dim sLogixTransNum As String
        Dim iStatusSV As Integer
        Dim bstatus As Boolean = True
        Dim lUseCustomerPK As Long
        Dim lLocalId As Long
        Dim iServerSerial As Integer
        Dim LastSVUpdate As Date = Date.Parse("01/01/1900")
        Dim LastPointsUpdate As Date = Date.Parse("01/01/1900")
        Dim LastUpdate As Date
        Dim drTranRows() As DataRow

        If lHouseholdPK = 0 Then
            lUseCustomerPK = lCustomerPK
        Else
            lUseCustomerPK = lHouseholdPK
        End If
        If CustomerTypeID = 1 Then
            sHouseHoldID = sCustomerID
        Else
            If lHouseholdPK > 0 Then
                MyCommon.QueryStr = "select ExtCardIDOriginal as ExtCardID from CardIDs with (NoLock) where CustomerPK = " & lHouseholdPK & ";"
                dtXS = MyCommon.LXS_Select
                If dtXS.Rows.Count > 0 Then
                    sHouseHoldID = MyCryptLib.SQL_StringDecrypt(dtXS.Rows(0).Item(0).ToString())
                Else
                    sHouseHoldID = ""
                End If
            Else
                sHouseHoldID = ""
            End If
        End If


        MyCommon.QueryStr = "select LogixTransNum, ExtLocationCode, TransDate, TerminalNum,POSTransNum" & _
                            " from TransHistory with (NoLock)" & _
                            " where CustomerPrimaryExtId='" & sCustomerID & "' and CustomerTypeID=" & CustomerTypeID & _
                            " and TransDate >= '" & StartDate & "' and TransDate <= '" & EndDate & "';"

        dtWH = MyCommon.LWH_Select
        If dtWH.Rows.Count > 0 Then
            For Each drWH In dtWH.Rows
                sLocationCode = drWH.Item("ExtLocationCode")
                sLogixTransNum = drWH.Item("LogixTransNum")
                MyCommon.QueryStr = "select LocationName from Locations with (NoLock) where ExtLocationCode = '" & sLocationCode & "';"
                dtRT = MyCommon.LRT_Select
                If dtRT.Rows.Count > 0 Then
                    sLocationName = MyCommon.NZ(dtRT.Rows(0).Item(0), "")
                Else
                    sLocationName = ""
                End If
                drTransaction = dtTransactions.NewRow()
                drTransaction.Item("LogixTransactionNumber") = sLogixTransNum
                drTransaction.Item("Date") = drWH.Item("TransDate")
                drTransaction.Item("LocationCode") = drWH.Item("ExtLocationCode")
                drTransaction.Item("LocationName") = sLocationName
                drTransaction.Item("CardID") = sCustomerID
                drTransaction.Item("HouseHoldID") = sHouseHoldID
                dtTransactions.Rows.Add(drTransaction)

                MyCommon.QueryStr = "select SVProgramID,LocalID,ServerSerial,QtyEarned,QtyUsed,StatusFlag,ExpireDate from SVHistory with (NoLock) " & _
                                    "where Deleted=0 and CustomerPK=" & lUseCustomerPK & " and LogixTransNum='" & sLogixTransNum & "';"
                dtXS = MyCommon.LXS_Select
                If dtXS.Rows.Count > 0 Then
                    For Each drXS In dtXS.Rows
                        drStoredValues = dtStoredValues.NewRow()
                        drStoredValues.Item("LogixTransactionNumber") = sLogixTransNum
                        drStoredValues.Item("ProgramID") = drXS.Item("SVProgramID")
                        iStatusSV = drXS.Item("StatusFlag")
                        Select Case iStatusSV
                            Case 1
                                drStoredValues.Item("Amount") = drXS.Item("QtyEarned")
                                drStoredValues.Item("Action") = "Earned"
                            Case 2
                                drStoredValues.Item("Amount") = "0"
                                drStoredValues.Item("Action") = "Revoked"
                            Case 3
                                drStoredValues.Item("Amount") = "0"
                                drStoredValues.Item("Action") = "Expired"
                            Case 4
                                drStoredValues.Item("Amount") = drXS.Item("QtyUsed")
                                drStoredValues.Item("Action") = "Redeemed"
                                drStoredValues.Item("Status") = "Redeemed"
                            Case Else
                                drStoredValues.Item("Amount") = "0"
                                drStoredValues.Item("Action") = "Unknown"
                        End Select

                        iServerSerial = drXS.Item("ServerSerial")
                        lLocalId = drXS.Item("LocalID")
                        MyCommon.QueryStr = "select SVProgramID,StatusFlag,ExpireDate from StoredValue with (NoLock) " & _
                                            "where Deleted=0 and LocalID=" & lLocalId & " and ServerSerial=" & iServerSerial & ";"
                        dt = MyCommon.LXS_Select
                        If dt.Rows.Count > 0 Then
                            drStoredValues.Item("ExpirationDate") = dt.Rows(0).Item("ExpireDate")
                            iStatusSV = dt.Rows(0).Item("StatusFlag")
                            Select Case iStatusSV
                                Case 1
                                    drStoredValues.Item("Status") = "Earned"
                                Case 2
                                    drStoredValues.Item("Status") = "Revoked"
                                Case 3
                                    drStoredValues.Item("Status") = "Expired"
                                Case 4
                                    drStoredValues.Item("Status") = "Redeemed"
                                Case Else
                                    drStoredValues.Item("Status") = "Unknown"
                            End Select
                        Else
                            drStoredValues.Item("ExpirationDate") = drXS.Item("ExpireDate")
                            drStoredValues.Item("Status") = "Unknown"
                        End If

                        dtStoredValues.Rows.Add(drStoredValues)
                    Next
                End If

                MyCommon.QueryStr = "select ProgramID,AdjAmount from PointsHistory with (NoLock) " & _
                                    "where CustomerPK=" & lUseCustomerPK & " and LogixTransNum='" & sLogixTransNum & "';"
                dtXS = MyCommon.LXS_Select
                If dtXS.Rows.Count > 0 Then
                    For Each drXS In dtXS.Rows
                        drPoints = dtPoints.NewRow()
                        drPoints.Item("LogixTransactionNumber") = sLogixTransNum
                        drPoints.Item("ProgramID") = drXS.Item("ProgramID")
                        drPoints.Item("Amount") = drXS.Item("AdjAmount")
                        dtPoints.Rows.Add(drPoints)
                    Next
                End If

            Next
        End If

        ' Get adjustments made during this period
        sLogixTransNum = "Adjustments:  " & Format(Now, sDateFormat)

        MyCommon.QueryStr = "select SVProgramID,LocalID,ServerSerial,QtyEarned,QtyUsed,StatusFlag,ExpireDate,LastUpdate from SVHistory with (NoLock) " & _
                            "where Deleted=0 and CustomerPK = " & lUseCustomerPK & " and isnull(LogixTransNum,'0')='0' " & _
                            "and LastUpdate >= '" & StartDate & "' and LastUpdate <= '" & EndDate & "' order by LastUpdate;"
        dtXS = MyCommon.LXS_Select
        If dtXS.Rows.Count > 0 Then
            For Each drXS In dtXS.Rows
                LastUpdate = drXS.Item("LastUpdate")
                LastUpdate = Date.Parse(Format(LastUpdate, sDateFormat))
                If LastUpdate <> LastSVUpdate Then
                    LastSVUpdate = LastUpdate
                    sLogixTransNum = "Adjustments:  " & Format(LastUpdate, sDateFormat)
                    drTransaction = dtTransactions.NewRow()
                    drTransaction.Item("LogixTransactionNumber") = sLogixTransNum
                    drTransaction.Item("Date") = LastUpdate
                    drTransaction.Item("LocationCode") = "Central"
                    drTransaction.Item("LocationName") = "Adjustments"
                    drTransaction.Item("CardID") = sCustomerID
                    drTransaction.Item("HouseHoldID") = sHouseHoldID
                    dtTransactions.Rows.Add(drTransaction)
                End If
                drStoredValues = dtStoredValues.NewRow()
                drStoredValues.Item("LogixTransactionNumber") = sLogixTransNum
                drStoredValues.Item("ProgramID") = drXS.Item("SVProgramID")
                iStatusSV = drXS.Item("StatusFlag")
                Select Case iStatusSV
                    Case 1
                        drStoredValues.Item("Amount") = drXS.Item("QtyEarned")
                        drStoredValues.Item("Action") = "Earned"
                    Case 2
                        drStoredValues.Item("Amount") = "0"
                        drStoredValues.Item("Action") = "Revoked"
                    Case 3
                        drStoredValues.Item("Amount") = "0"
                        drStoredValues.Item("Action") = "Expired"
                    Case 4
                        drStoredValues.Item("Amount") = drXS.Item("QtyUsed")
                        drStoredValues.Item("Action") = "Redeemed"
                        drStoredValues.Item("Status") = "Redeemed"
                    Case Else
                        drStoredValues.Item("Amount") = "0"
                        drStoredValues.Item("Action") = "Unknown"
                End Select

                iServerSerial = drXS.Item("ServerSerial")
                lLocalId = drXS.Item("LocalID")
                MyCommon.QueryStr = "select SVProgramID,StatusFlag,ExpireDate from StoredValue with (NoLock) " & _
                                    "where Deleted=0 and LocalID=" & lLocalId & " and ServerSerial=" & iServerSerial & ";"
                dt = MyCommon.LXS_Select
                If dt.Rows.Count > 0 Then
                    drStoredValues.Item("ExpirationDate") = dt.Rows(0).Item("ExpireDate")
                    iStatusSV = dt.Rows(0).Item("StatusFlag")
                    Select Case iStatusSV
                        Case 1
                            drStoredValues.Item("Status") = "Earned"
                        Case 2
                            drStoredValues.Item("Status") = "Revoked"
                        Case 3
                            drStoredValues.Item("Status") = "Expired"
                        Case 4
                            drStoredValues.Item("Status") = "Redeemed"
                        Case Else
                            drStoredValues.Item("Status") = "Unknown"
                    End Select
                Else
                    drStoredValues.Item("ExpirationDate") = drXS.Item("ExpireDate")
                    drStoredValues.Item("Status") = "Unknown"
                End If

                dtStoredValues.Rows.Add(drStoredValues)
            Next
        End If

        MyCommon.QueryStr = "select ProgramID,AdjAmount,LastUpdate from PointsHistory with (NoLock) " & _
                            "where CustomerPK=" & lUseCustomerPK & "and isnull(LogixTransNum,'0')='0' " & _
                            "and LastUpdate >= '" & StartDate & "' and LastUpdate <= '" & EndDate & "' order by LastUpdate;"
        dtXS = MyCommon.LXS_Select
        If dtXS.Rows.Count > 0 Then
            For Each drXS In dtXS.Rows
                LastUpdate = drXS.Item("LastUpdate")
                LastUpdate = Date.Parse(Format(LastUpdate, sDateFormat))
                If LastUpdate <> LastPointsUpdate Then
                    LastPointsUpdate = LastUpdate
                    drTranRows = dtTransactions.Select("Date = '" & Format(LastUpdate, sDateFormat) & "' and Locationcode='Central'")
                    If drTranRows.Length > 0 Then
                        sLogixTransNum = drTranRows(0).Item("LogixTransactionNumber")
                    Else
                        sLogixTransNum = "Adjustments:  " & Format(LastUpdate, sDateFormat)
                        drTransaction = dtTransactions.NewRow()
                        drTransaction.Item("LogixTransactionNumber") = sLogixTransNum
                        drTransaction.Item("Date") = LastUpdate
                        drTransaction.Item("LocationCode") = "Central"
                        drTransaction.Item("LocationName") = "Adjustments"
                        drTransaction.Item("CardID") = sCustomerID
                        drTransaction.Item("HouseHoldID") = sHouseHoldID
                        dtTransactions.Rows.Add(drTransaction)
                    End If
                End If
                drPoints = dtPoints.NewRow()
                drPoints.Item("LogixTransactionNumber") = sLogixTransNum
                drPoints.Item("ProgramID") = drXS.Item("ProgramID")
                drPoints.Item("Amount") = drXS.Item("AdjAmount")
                dtPoints.Rows.Add(drPoints)
            Next
        End If

        If dtTransactions.Rows.Count > 0 Then dtTransactions.AcceptChanges()
        If dtStoredValues.Rows.Count > 0 Then dtStoredValues.AcceptChanges()
        If dtPoints.Rows.Count > 0 Then dtPoints.AcceptChanges()

        Return bstatus
    End Function

    Private Function GetPointsBalances(ByVal sCustomerID As String, ByVal lCustomerPK As Long, ByVal lHouseholdPK As Long, ByRef dtBalances As DataTable) As Boolean
        Dim dt As DataTable = Nothing
        Dim dr As DataRow
        Dim dtPoints As DataTable
        Dim drBalances As DataRow
        Dim sProgramId As String
        Dim sAmount As String
        Dim bstatus As Boolean = True
        Dim lUseCustomerPK As Long
        Dim decAmount As Decimal
        Dim iCategoryId As Integer
        Dim sCategory As String

        If lHouseholdPK = 0 Then
            lUseCustomerPK = lCustomerPK
        Else
            lUseCustomerPK = lHouseholdPK
        End If

        MyCommon.QueryStr = "select P.ProgramID, P.Amount, PV.ExternalID" & _
                            " from Points as P with (NoLock)" & _
                            " inner join PromoVariables as PV with (NoLock) on PV.PromoVarID = P.PromoVarID" & _
                            " where P.CustomerPK = " & lUseCustomerPK & " and PV.ExternalID is null" & _
                            " order by ProgramID;"
        dt = MyCommon.LXS_Select
        If dt.Rows.Count > 0 Then
            ' Get all External Points Programs balances
            If MyCommon.Fetch_SystemOption(80) = "1" Then
                Try
                    If sCustomerID.Length > 0 Then
                        Dim ExternalPP As Copient.ExternalRewards
                        Dim sDb2Connection As String
                        Dim iDb2Connection As Integer = 5

                        sDb2Connection = MyCommon.Fetch_CM_SystemOption(iDb2Connection)
                        ExternalPP = New Copient.ExternalRewards("", "", "", sDb2Connection)
                        ' add balances for external points programs for which the customer has current balance.
                        ExternalPP.appendExtProgramBalances(sCustomerID, False, dt, MyCommon)
                    End If
                Catch
                    Throw
                End Try
            End If

            For Each dr In dt.Rows
                decAmount = Decimal.Parse(dr.Item(1))
                If decAmount > 0.0 Then
                    sProgramId = dr.Item(0)
                    sAmount = decAmount.ToString("#.##")
                    MyCommon.QueryStr = "select ProgramID, ProgramName, CategoryID from PointsPrograms with (NoLock) " & _
                                        "where ProgramID=" & sProgramId & ";"
                    dtPoints = MyCommon.LRT_Select
                    If dtPoints.Rows.Count > 0 Then
                        drBalances = dtBalances.NewRow()
                        drBalances.Item("ProgramID") = sProgramId
                        drBalances.Item("ProgramName") = dtPoints.Rows(0).Item("ProgramName")
                        iCategoryId = MyCommon.NZ(dtPoints.Rows(0).Item("CategoryID"), 0)
                        sCategory = GetCategoryFromId(iCategoryId)
                        drBalances.Item("Category") = sCategory
                        drBalances.Item("Balance") = sAmount
                        dtBalances.Rows.Add(drBalances)
                    End If
                End If
            Next
            If dtBalances.Rows.Count > 0 Then
                dtBalances.AcceptChanges()
            End If
        Else
            bstatus = False
        End If

        Return bstatus
    End Function

    Private Function GetCategoryFromId(ByVal iCategoryId As Integer) As String
        Dim dt As DataTable = Nothing
        Dim sCategory As String

        MyCommon.QueryStr = "select Description from OfferCategories with (NoLock) where Deleted=0 and OfferCategoryId = " & iCategoryId & ";"
        dt = MyCommon.LRT_Select
        If dt.Rows.Count > 0 Then
            sCategory = MyCommon.NZ(dt.Rows(0).Item(0), "")
        Else
            sCategory = ""
        End If

        Return sCategory
    End Function


    Private Function GetStoredValueBalances(ByVal lCustomerPK As Long, ByVal lHouseholdPK As Long, ByVal iNumDays As Integer, ByRef dtBalances As DataTable) As Boolean
        Dim dt1 As DataTable = Nothing
        Dim dt2 As DataTable = Nothing
        Dim dr As DataRow
        Dim dtSV As DataTable
        Dim drBalances As DataRow
        Dim sProgramId As String
        Dim sAboutToExpireQuantity As String
        Dim bstatus As Boolean = True
        Dim lUseCustomerPK As Long
        Dim iQuantity As Decimal

        If lHouseholdPK = 0 Then
            lUseCustomerPK = lCustomerPK
        Else
            lUseCustomerPK = lHouseholdPK
        End If

        MyCommon.QueryStr = "select SVProgramID as ProgramID, Sum(QtyEarned) - Sum(QtyUsed) as Quantity, Min(ExpireDate) as ExpireDate " & _
                            "from StoredValue with (NoLock) " & _
                            "where CustomerPK=" & lUseCustomerPK & " and Deleted=0 and StatusFlag=1 and ExpireDate >= getdate() " & _
                            "group by SVProgramID order by SVProgramID;"
        dt1 = MyCommon.LXS_Select
        If dt1.Rows.Count > 0 Then
            For Each dr In dt1.Rows
                iQuantity = Integer.Parse(MyCommon.NZ(dr.Item("Quantity"), "0"))
                If iQuantity > 0 Then
                    sProgramId = dr.Item("ProgramID")

                    MyCommon.QueryStr = "select Sum(QtyEarned) - Sum(QtyUsed) as Quantity " & _
                                        "from StoredValue with (NoLock) " & _
                                        "where CustomerPK=" & lUseCustomerPK & " and Deleted=0 and StatusFlag=1 and ExpireDate >= getdate() " & _
                                        "and ExpireDate <= dateadd(d," & iNumDays & ", getdate()) and SVProgramID=" & sProgramId & ";"
                    dt2 = MyCommon.LXS_Select
                    If dt2.Rows.Count > 0 Then
                        sAboutToExpireQuantity = MyCommon.NZ(dt2.Rows(0).Item(0), "0")
                    Else
                        sAboutToExpireQuantity = "0"
                    End If

                    MyCommon.QueryStr = "select Name, Value, UnitOfMeasureLimit from StoredValuePrograms with (NoLock) " & _
                                        "where SVProgramID=" & sProgramId & ";"
                    dtSV = MyCommon.LRT_Select
                    If dtSV.Rows.Count > 0 Then
                        drBalances = dtBalances.NewRow()
                        drBalances.Item("ProgramID") = sProgramId
                        drBalances.Item("ProgramName") = dtSV.Rows(0).Item("Name")
                        drBalances.Item("Value") = dtSV.Rows(0).Item("Value")
                        drBalances.Item("UnitOfMeasureLimit") = dtSV.Rows(0).Item("UnitOfMeasureLimit")
                        drBalances.Item("Balance") = iQuantity
                        drBalances.Item("BalanceExpireDate") = dr.Item("ExpireDate")
                        drBalances.Item("AboutToExpireQuantity") = sAboutToExpireQuantity
                        drBalances.Item("AboutToExpireDays") = iNumDays
                        dtBalances.Rows.Add(drBalances)
                    End If
                End If
            Next
            If dtBalances.Rows.Count > 0 Then
                dtBalances.AcceptChanges()
            End If
        Else
            bstatus = False
        End If

        Return bstatus
    End Function


    Private Function GetAllSVBalancesForCustomer(ByVal lCustomerPK As Long, ByVal lHouseholdPK As Long, ByRef dtBalances As DataTable) As Boolean
        Dim dt1 As DataTable = Nothing
        Dim dt2 As DataTable = Nothing
        Dim dr As DataRow
        Dim drBalances As DataRow
        Dim bstatus As Boolean = True
        Dim lUseCustomerPK As Long


        If lHouseholdPK = 0 Then
            lUseCustomerPK = lCustomerPK
        Else
            lUseCustomerPK = lHouseholdPK
        End If

        MyCommon.QueryStr = "select SVProgramID as ProgramID, Sum(QtyEarned) - Sum(QtyUsed) as Quantity, Min(ExpireDate) as ExpireDate " & _
                            "from StoredValue with (NoLock) " & _
                            "where CustomerPK=" & lUseCustomerPK & " and Deleted=0 and StatusFlag=1 and ExpireDate >= getdate() " & _
                            "group by SVProgramID, ExpireDate HAVING (Sum(QtyEarned) - Sum(QtyUsed)) > 0  order by SVProgramID;"

        dt1 = MyCommon.LXS_Select
        If dt1.Rows.Count > 0 Then
            For Each dr In dt1.Rows
                drBalances = dtBalances.NewRow()
                drBalances.Item("ProgramID") = dr.Item("ProgramID")
                drBalances.Item("ExpirationDate") = dr.Item("ExpireDate")
                drBalances.Item("Balance") = Integer.Parse(MyCommon.NZ(dr.Item("Quantity"), "0"))
                dtBalances.Rows.Add(drBalances)
            Next
            If dtBalances.Rows.Count > 0 Then
                dtBalances.AcceptChanges()
            End If
        Else
            bstatus = False
        End If

        Return bstatus
    End Function


    Private Function RetrieveAccumulationBalance(ByVal OfferID As Long, ByVal CustomerPK As Long) As Decimal
        'This function is used by Send_XMLCurrentOffers to get a customer's accumulation balance in an offer 
        Dim rst As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim EngineID As Integer = -1
        Dim AccumProgram As Boolean = False
        Dim RewardOptionID As Long = -1
        Dim HHEnable As Boolean = False
        Dim UnitType As Integer = 0
        Dim HouseholdPK As Integer = 0
        Dim TotalAccum As Double
        Dim AccumulationBalance As Decimal = 0

        MyCommon.QueryStr = "select EngineID from OfferIDs with (NoLock) where OfferID = " & OfferID
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
            EngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), -1)
        End If

        If EngineID = 2 Then
            'First find if there's accumulation
            MyCommon.QueryStr = "select IPG.AccumMin, RO.RewardOptionID, RO.HHEnable, IPG.QtyUnitType " & _
                                "from CPE_IncentiveProductGroups as IPG with (NoLock) " & _
                                "inner join CPE_RewardOptions as RO with (NoLock) on IPG.RewardOptionID=RO.RewardOptionID and IPG.Deleted=0 and IPG.ExcludedProducts=0 and RO.Deleted=0 " & _
                                "where RO.IncentiveID=" & OfferID & ";"
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count > 0) Then
                If MyCommon.NZ(rst.Rows(0).Item("AccumMin"), 0) > 0 Then
                    AccumProgram = True
                End If
                UnitType = MyCommon.NZ(rst.Rows(0).Item("QtyUnitType"), 2)
                RewardOptionID = rst.Rows(0).Item("RewardOptionID")
                HHEnable = MyCommon.NZ(rst.Rows(0).Item("HHEnable"), False)
                If HHEnable Then
                    MyCommon.QueryStr = "select HHPK from Customers with (NoLock) where CustomerPK='" & CustomerPK & "';"
                    rst = MyCommon.LXS_Select
                    If (rst.Rows.Count > 0) Then
                        HouseholdPK = MyCommon.NZ(rst.Rows(0).Item("HHPK"), 0)
                    End If
                End If
            End If

            If AccumProgram Then
                'There's accumulation, so get the data
                If HHEnable Then
                    MyCommon.QueryStr = "select RA.AccumulationDate, RA.LastServerID, RA.ServerSerial, RA.CustomerPK, " & _
                                        "RA.PurchCustomerPK, RA.TotalPrice, RA.QtyPurchased, RA.Deleted, RA.LocationID " & _
                                        "from CPE_RewardAccumulation as RA with (NoLock) " & _
                                        "where (RA.CustomerPK=" & CustomerPK & " or RA.CustomerPK=" & HouseholdPK & ") and RA.RewardOptionID=" & RewardOptionID & " order by AccumulationDate;"
                Else
                    MyCommon.QueryStr = "select RA.AccumulationDate, RA.LastServerID, RA.ServerSerial, RA.CustomerPK, " & _
                                        "RA.PurchCustomerPK, RA.TotalPrice, RA.QtyPurchased, RA.Deleted, RA.LocationID " & _
                                        "from CPE_RewardAccumulation as RA with (NoLock) " & _
                                        "where RA.CustomerPK=" & CustomerPK & " and RA.RewardOptionID=" & RewardOptionID & " order by AccumulationDate;"
                End If
                rst = MyCommon.LXS_Select
                If (rst.Rows.Count > 0) Then
                    TotalAccum = 0
                    For Each row In rst.Rows
                        If UnitType = 1 Then
                            TotalAccum = TotalAccum + Math.Round(MyCommon.NZ(row.Item("QtyPurchased"), 0), 0)
                        ElseIf UnitType = 2 Then
                            TotalAccum = TotalAccum + Math.Round(MyCommon.NZ(row.Item("TotalPrice"), 0), 2)
                        ElseIf UnitType = 3 Then
                            TotalAccum = TotalAccum + Math.Round(MyCommon.NZ(row.Item("QtyPurchased"), 0), 3)
                        End If
                    Next
                    AccumulationBalance = TotalAccum
                End If
            End If

        End If

        Return AccumulationBalance
    End Function

    Private Function RetrieveGraphicPath(ByVal OfferID As Long) As String
        'This function is used by Send_XMLCurrentOffers and Send_XMLGroupOffers to get the path of an offer's graphic
        Dim GraphicsFileName As String = ""
        Dim GraphicsFilePath As String = ""
        Dim GraphicsNewFilePath As String = ""
        Dim rst As System.Data.DataTable

        Try
            'Find if a graphic is assigned to this offer
            MyCommon.QueryStr = "select OnScreenAdID, Name, ImageType, Width, Height from OnScreenAds with (NoLock) where Deleted=0 and OnScreenAdID in " & _
                                " (select OutputID from CPE_Deliverables with (NoLock) where Deleted=0 and DeliverableTypeID=1 and RewardOptionPhase=1 and RewardOptionID in " & _
                                "  (select RewardOptionID from CPE_RewardOptions with (NoLock) where IncentiveID=" & OfferID & " and TouchResponse=0 and Deleted=0));"
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count > 0) Then
                GraphicsFilePath = MyCommon.Fetch_SystemOption(47)
                If (GraphicsFilePath.Length > 0) Then
                    If (Right(GraphicsFilePath, 1) <> "\") Then
                        GraphicsFilePath += "\"
                    End If
                End If
                GraphicsFileName = MyCommon.NZ(rst.Rows(0).Item("OnScreenAdID"), "") & "img_tn."
                GraphicsFileName += IIf(MyCommon.NZ(rst.Rows(0).Item("ImageType"), 1) = 2, "gif", "jpg")
                GraphicsFilePath += GraphicsFileName

        If (bDisableGraphicFileSearch) Then
          WriteDebug("RetrieveGraphicPath - Returning Graphic Bypassed", DebugState.CurrentTime)
          Return GraphicsFileName
        End If

                If (File.Exists(GraphicsFilePath)) Then
                    GraphicsNewFilePath = Server.MapPath("")
                    GraphicsNewFilePath = GraphicsNewFilePath.Substring(0, GraphicsNewFilePath.LastIndexOf("\"))
                    GraphicsNewFilePath += "\images\" & GraphicsFileName
                    If Not (File.Exists(GraphicsNewFilePath)) Then
                        File.Copy(GraphicsFilePath, GraphicsNewFilePath)
                        If Not (File.Exists(GraphicsNewFilePath)) Then
                            'GraphicsFileName = ""
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            'GraphicsFileName = ""
        End Try
         WriteDebug("RetrieveGraphicPath - Returning Graphic", DebugState.CurrentTime)
        Return GraphicsFileName
    End Function

    Private Function RetrievePrintedMessage(ByVal OfferID As Long) As String
        'This function is used by Send_XMLCurrentOffers and Send_XMLGroupOffers to get the text of an offer's printed message  
        Dim PMsgBuf As New StringBuilder()
        Dim rst As System.Data.DataTable

        
        MyCommon.QueryStr = "select PMTypes.Description, PMTiers.TierLevel, PMTiers.BodyText " & _
                            "from PrintedMessages as PM with (NoLock) " & _
                            "inner join PrintedMessageTiers as PMTiers with (NoLock) on PM.MessageID=PMTiers.MessageID " & _
                            "inner join PrintedMessageTypes as PMTypes with (NoLock) on PM.MessageTypeID=PMTypes.TypeID " & _
                            "inner join CPE_Deliverables as D with (NoLock) on D.OutputID=PM.MessageID " & _
                            "inner join CPE_RewardOptions as RO with (NoLock) on RO.RewardOptionID=D.RewardOptionID " & _
                            "where RO.IncentiveID=" & OfferID & " and RO.Deleted=0 and D.Deleted=0 and D.RewardOptionPhase=1 and DeliverableTypeID=4;"
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
            PMsgBuf.Append(MyCommon.NZ(rst.Rows(0).Item("BodyText"), ""))
        End If
        

        Return PMsgBuf.ToString()
    End Function

    Private Function Send_XMLCurrentOffers(ByVal CustomerPK As Long) As DataTable
        'This function is used by OfferList and returns a list of customer offers
        Dim rst As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim rst2 As System.Data.DataTable
        Dim row2 As System.Data.DataRow
        Dim rst3 As System.Data.DataTable
        Dim rst4 As System.Data.DataTable
        Dim row4 As System.Data.DataRow
        Dim rst5 As System.Data.DataTable
        Dim rstWeb As System.Data.DataTable
        Dim rstExcluded As System.Data.DataTable
        Dim rstDispDates As System.Data.DataTable
        Dim rowCount As Integer
        Dim Employee As Boolean
        Dim ExtCustomerID As String = ""
        Dim CustomerGroupID As Long
        Dim OfferID As Integer
        Dim OfferName As String
        Dim OfferDesc As String
        Dim OfferCategoryID As Integer
        Dim OfferStart As Date
        Dim OfferEnd As Date
        Dim DispStartDate As Date
        Dim DispEndDate As Date
        Dim OfferOdds As Integer
        Dim InstantWin As Integer
        Dim OfferDaysLeft As Integer
        Dim GraphicsFileName As String = ""
        Dim PrintedMessage As String = ""
        Dim AccumulationBalance As Decimal = 0
        Dim ProgramID As String = ""
        Dim ProgID As Integer
        Dim ProgramName As String = ""
        Dim Amount As Long
        Dim AllowOptOut As Integer
        Dim EmployeesOnly As Integer
        Dim EmployeesExcluded As Integer
        Dim dt As New DataTable
        Dim dtOffers As System.Data.DataTable
        Dim retString As New StringBuilder
        Dim CgBuf As New StringBuilder()
        Dim MyLookup As New Copient.CustomerLookup
        Dim Localization As Copient.Localization
        Dim BalRetCode As Copient.CustomerLookup.RETURN_CODE
        Dim Balances(-1) As Copient.CustomerLookup.PointsBalance
        Dim LanguageID As Integer
        Dim bc As SqlBulkCopy = Nothing
        Dim ClientOfferID As String = ""

        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
        If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

        'Create customer groups temp table    
        MyCommon.QueryStr = "BEGIN TRY DROP TABLE #CustomerGroups; END TRY BEGIN CATCH END CATCH;" & _
                               "CREATE TABLE #CustomerGroups (GroupID bigint);"
        MyCommon.LRT_Execute()

        'Create excluded offers temp table    
        MyCommon.QueryStr = "BEGIN TRY DROP TABLE #ExcludedOffers; END TRY BEGIN CATCH END CATCH;" & _
                               "CREATE TABLE #ExcludedOffers (OfferID bigint);"
        MyCommon.LRT_Execute()

        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
        If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
        If MyCommon.PMRTadoConn.State = ConnectionState.Closed AndAlso MyCommon.IsIntegrationInstalled(Integrations.PREFERENCE_MANAGER) Then
            MyCommon.Open_PrefManRT()
        End If
        Localization = New Copient.Localization(MyCommon)
        LanguageID = Localization.GetCustLanguageID(CustomerPK)
  
        'Create a new datatable to hold the results we'll be assembling
        dtOffers = New DataTable
        dtOffers.Columns.Add("OfferID", System.Type.GetType("System.Int64"))
        dtOffers.Columns.Add("Name", System.Type.GetType("System.String"))
        dtOffers.Columns.Add("Description", System.Type.GetType("System.String"))
        dtOffers.Columns.Add("OfferCategoryID", System.Type.GetType("System.Int32"))
        dtOffers.Columns.Add("StartDate", System.Type.GetType("System.DateTime"))
        dtOffers.Columns.Add("EndDate", System.Type.GetType("System.DateTime"))
        If (MyCommon.Fetch_CM_SystemOption(85) = "1" Or MyCommon.Fetch_UE_SystemOption(143) = "1") Then
            dtOffers.Columns.Add("DisplayStartDate", System.Type.GetType("System.DateTime"))
            dtOffers.Columns.Add("DisplayEndDate", System.Type.GetType("System.DateTime"))
        End If
        dtOffers.Columns.Add("CustomerGroupID", System.Type.GetType("System.Int64"))
        dtOffers.Columns.Add("AllowOptOut", System.Type.GetType("System.Boolean"))
        dtOffers.Columns.Add("EmployeesOnly", System.Type.GetType("System.Boolean"))
        dtOffers.Columns.Add("EmployeesExcluded", System.Type.GetType("System.Boolean"))
        dtOffers.Columns.Add("Points", System.Type.GetType("System.Int32"))
        dtOffers.Columns.Add("Accumulation", System.Type.GetType("System.Decimal"))
        dtOffers.Columns.Add("BodyText", System.Type.GetType("System.String"))
        dtOffers.Columns.Add("Graphic", System.Type.GetType("System.String"))
        dtOffers.Columns.Add("ClientOfferID", System.Type.GetType("System.String"))

        MyCommon.QueryStr = "select C.CustomerPK, C.Employee, C.CustomerStatusID as CardStatusID, CE.Email " & _
                        "from Customers as C with (NoLock) " & _
                        "left join CustomerExt as CE with (NoLock) on CE.CustomerPK=C.CustomerPK " & _
                        "where C.CustomerPK=" & CustomerPK & ";"
        rst = MyCommon.LXS_Select
        If (rst.Rows.Count > 0) Then
            'A customer was found, so assign values to variables
            Employee = rst.Rows(0).Item("Employee")

            'Next, get the associated customer groups
            MyCommon.QueryStr = "select distinct CustomerGroupID from GroupMembership with (NoLock) where CustomerPK=" & CustomerPK & " and Deleted=0;"
            rst = MyCommon.LXS_Select()
            rst.Rows.Add(New String() {"1"})
            rst.Rows.Add(New String() {"2"})
            rowCount = rst.Rows.Count

            'Build CustomerGroup temp table
            WriteDebug("Starting Bulk Copy - CustomerGroup", DebugState.CurrentTime)
            bc = New SqlBulkCopy(MyCommon.LRTadoConn)
            bc.BatchSize = rowCount
            bc.DestinationTableName = "#CustomerGroups"
            bc.WriteToServer(rst)
            bc.Close()
            WriteDebug("Ending Bulk Copy - CustomerGroup", DebugState.CurrentTime)
      
            'Return a list of offers that the customer is excluded from
            MyCommon.QueryStr = "select OfferID from OfferConditions as OC with (NoLock) inner join #CustomerGroups as CG on OC.ExcludedID=CG.GroupID " & _
                                "union " & _
                                "select RewardOptionID as OfferID from CPE_IncentiveCustomerGroups as ICG with (NoLock) inner join #CustomerGroups as CG on ICG.CustomerGroupID=CG.GroupID " & _
                                "and ExcludedUsers=1 and Deleted=0;"
            rst = MyCommon.LRT_Select()
        
            WriteDebug("Starting ExcludedOffers Bulk Copy", DebugState.CurrentTime)
            bc = New SqlBulkCopy(MyCommon.LRTadoConn)
            bc.BatchSize = rst.Rows.Count
            bc.DestinationTableName = "#ExcludedOffers"
            bc.WriteToServer(rst)
            bc.Close()
            WriteDebug("Finished ExcludedOffers Bulk Copy", DebugState.CurrentTime)

            'The customer's in at least one group, so for each one we'll grab the associated offer(s)
      MyCommon.QueryStr = "select distinct O.OfferID, O.ExtOfferID, O.IsTemplate, O.CMOADeployStatus, O.StatusFlag, O.OddsOfWinning, O.InstantWin, " & _
                          "O.Name, O.Description, O.OfferCategoryID, O.ProdStartDate, O.ProdEndDate, 0 as AllowOptOut, O.EmployeeFiltering as EmployeesOnly, O.NonEmployeesOnly as EmployeesExcluded, LinkID, OID.EngineID " & _
                          "from CM_ST_Offers as O with (NoLock) " & _
                          "left join CM_ST_OfferConditions as OC with (NoLock) on OC.OfferID=O.OfferID " & _
                        "left outer join #ExcludedOffers as EO on O.OfferID=EO.OfferID " & _
                          "inner join OfferIDs as OID with (NoLock) on OID.OfferID=O.OfferID " & _
                        "inner join #CustomerGroups as CG on OC.LinkID=CG.GroupID " & _
                          "where O.IsTemplate=0 and O.Deleted=0 and O.CMOADeployStatus=1 and OC.Deleted=0 and OC.ConditionTypeID=1 " & _
                        "and O.DisabledOnCFW=0 and ProdEndDate>'" & Today.AddDays(-1).ToString & "' and O.ProdStartDate<=GETDATE() and OID.EngineID != 3 " & _
                        "union " & _
                          "select distinct I.IncentiveID, I.ClientOfferID, I.IsTemplate, I.CPEOADeployStatus, I.StatusFlag, 0 as OddsOfWinning, 0 as InstantWin, " & _
                          "case when not(isnull(OT.OfferName, '')='') then isnull(OT.OfferName, '') else I.IncentiveName end as Name, " & _
                          "Convert(nvarchar(2000),I.Description) as Description, I.PromoClassID as OfferCategoryID, I.StartDate, I.EndDate, I.AllowOptOut, I.EmployeesOnly, I.EmployeesExcluded, ICG.CustomerGroupID, OID.EngineID " & _
                          "from CPE_ST_Incentives as I with (NoLock) " & _
                          "left join CPE_ST_RewardOptions as RO with (NoLock) on I.IncentiveID=RO.IncentiveID and RO.TouchResponse=0 and RO.Deleted=0 " & _
                          "left join CPE_ST_IncentiveCustomerGroups as ICG with (NoLock) on RO.RewardOptionID=ICG.RewardOptionID and ICG.ExcludedUsers=0 " & _
                          "left outer join #ExcludedOffers as EO on RO.RewardOptionID=EO.OfferID " & _
                          "inner join OfferIDs as OID with (NoLock) on OID.OfferID=I.IncentiveID " & _
                          "inner join #CustomerGroups as CG on ICG.CustomerGroupID=CG.GroupID " & _
                          "Left Join cpe_st_OfferTranslations as OT with (NoLock) on OT.OfferID=I.IncentiveID and OT.LanguageID=" & LanguageID & " " & _
                          "where (I.IsTemplate=0 and I.Deleted=0 and ICG.Deleted=0) " & _
                        "and I.DisabledOnCFW=0 and I.EndDate>'" & Today.AddDays(-1).ToString & "' and I.StartDate<=GETDATE() and OID.EngineID != 3 and EO.OfferID is null;"
      WriteDebug("Send_XMLCurrentOffers - Starting Query For Offers [LogixRT]", DebugState.CurrentTime)
      rst2 = MyCommon.LRT_Select
      WriteDebug("Send_XMLCurrentOffers - Completed Query For Offers with row count=" & rst2.Rows.Count, DebugState.CurrentTime)

      'Set the general info for each offer found
      If (rst2.Rows.Count > 0) Then
        For Each row2 In rst2.Rows
          OfferID = row2.Item("OfferID")
          OfferName = row2.Item("Name")
          OfferDesc = row2.Item("Description")
          If (OfferDesc.Trim() = "") Then OfferDesc = OfferName
          OfferCategoryID = MyCommon.NZ(row2.Item("OfferCategoryID"), 0)
          OfferStart = row2.Item("ProdStartDate")
          OfferEnd = row2.Item("ProdEndDate")
          OfferOdds = row2.Item("OddsOfWinning")
          InstantWin = MyCommon.NZ(row2.Item("InstantWin"), 0)
          OfferDaysLeft = DateDiff("d", Today, OfferEnd)
          AllowOptOut = IIf(MyCommon.NZ(row2.Item("AllowOptOut"), False), 1, 0)
          EmployeesOnly = IIf(MyCommon.NZ(row2.Item("EmployeesOnly"), False), 1, 0)
          EmployeesExcluded = IIf(MyCommon.NZ(row2.Item("EmployeesExcluded"), False), 1, 0)
          CustomerGroupID = row2.Item("LinkID")
      	  ClientOfferID = MyCommon.NZ(row2.Item("ExtOfferID"), "")

          'If enabled, test this OfferID for blown limits.  If blown, stop processing and continue to the next offer, consider it ineligible
          If(MyCommon.Fetch_CPE_SystemOption(198) = "1" And CPELimitsBlown(OfferID,CustomerPK))
            WriteDebug("Test Limit:Limit blown for..." & OfferID & "CustPK:" & CustomerPK, DebugState.CurrentTime)
            Continue For
          End If

          If (MyCommon.Fetch_CM_SystemOption(85) = "1" Or MyCommon.Fetch_UE_SystemOption(143) = "1") Then
            MyCommon.QueryStr = "SELECT DisplayStartDate,DisplayEndDate FROM OfferAccessoryFields with (NoLock) WHERE OfferID = " & OfferID & " "
            rstDispDates = MyCommon.LRT_Select
            If rstDispDates.Rows.Count > 0 Then
              DispStartDate = MyCommon.NZ(rstDispDates.Rows(0).Item("DisplayStartDate"), Nothing)
              DispEndDate = MyCommon.NZ(rstDispDates.Rows(0).Item("DisplayEndDate"), Nothing)
            Else
              DispStartDate = Nothing
              DispEndDate = Nothing
            End If
          End If
          
          If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
          If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
          
          'Create customer groups temp table    
          MyCommon.QueryStr = "BEGIN TRY DROP TABLE #CustomerGroups; END TRY BEGIN CATCH END CATCH;" & _
                                 "CREATE TABLE #CustomerGroups (GroupID bigint);"
          MyCommon.LRT_Execute()
          
          'Build CustomerGroup temp table
          WriteDebug("Starting Bulk Copy - CustomerGroup", DebugState.CurrentTime)
          bc = New SqlBulkCopy(MyCommon.LRTadoConn)
          bc.BatchSize = rowCount
          bc.DestinationTableName = "#CustomerGroups"
          bc.WriteToServer(rst)
          bc.Close()
          WriteDebug("Ending Bulk Copy - CustomerGroup", DebugState.CurrentTime)
                    
          'Filter out the website offers
          MyCommon.QueryStr = "select OfferID from OfferIDs with (NoLock) where OfferID=" & OfferID & " and EngineID=3;"
          rstWeb = MyCommon.LRT_Select

          'Filter out the offers where the customer is in the excluded customer group
          MyCommon.QueryStr = "select ExcludedID from OfferConditions as OC with (NoLock) " & _
                              "inner join #CustomerGroups as CG on OC.ExcludedID=CG.GroupID " & _
                              "where OfferID=" & OfferID & _
                              "union " & _
                              "select CustomerGroupID as ExcludedID from CPE_IncentiveCustomerGroups ICG with (NoLock) " & _
                              "inner join CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID=ICG.RewardOptionID " & _
                              "inner join #CustomerGroups as CG on ICG.CustomerGroupID=CG.GroupID " & _
                              "where ICG.Deleted=0 and RO.Deleted=0 and RO.IncentiveID=" & OfferID & " and ExcludedUsers=1; "
          rstExcluded = MyCommon.LRT_Select

          If (rstWeb.Rows.Count = 0 AndAlso rstExcluded.Rows.Count = 0) Then

            'Find the name of the associated (and non-excluding) location group
            MyCommon.QueryStr = "select OL.OfferID, OL.LocationGroupID, OL.Excluded, LG.Name from OfferLocations as OL with (NoLock) " & _
                                "left join LocationGroups as LG with (NoLock) on LG.LocationGroupID=OL.LocationGroupID " & _
                                "where OL.OfferID=" & OfferID & " and OL.Excluded=0;"
            rst3 = MyCommon.LRT_Select

            'Find any associated points programs
            MyCommon.QueryStr = "select O.Offerid, LinkID, ProgramName, PP.ProgramID, PromoVarID, isnull(PP.ExternalProgram, 0) as ExternalProgram, isnull(PP.ExtHostTypeID, 0) as ExtHostTypeID from OfferRewards as OFR with (NoLock) " & _
                                "left join RewardPoints as RP with (NoLock) on RP.RewardPointsID=OFR.LinkID " & _
                                "left join PointsPrograms as PP with (NoLock) on RP.ProgramID=PP.ProgramID " & _
                                "left join Offers as O with (NoLock) on O.OfferID=OFR.OfferID " & _
                                "where (RewardTypeID=2 and O.Deleted=0 and OFR.Deleted=0) " & _
                                "and RP.ProgramID is not null " & _
                                "and O.OfferID=" & OfferID & _
                                " union " & _
                                "select " & OfferID & " as OfferID, D.OutputID, PP.ProgramName, PP.ProgramID, PP.PromoVarID, isnull(PP.ExternalProgram, 0) as ExternalProgram, isnull(PP.ExtHostTypeID, 0) as ExtHostTypeID " & _
                                "from CPE_Deliverables D with (NoLock) inner join CPE_DeliverablePoints DP with (NoLock) on D.OutputID=DP.PKID " & _
                                "inner join PointsPrograms PP with (NoLock) on DP.ProgramID=PP.ProgramID " & _
                                "where D.RewardOptionID in (select RO.RewardOptionID from CPE_RewardOptions RO with (NoLock) where IncentiveID=" & OfferID & ") " & _
                                "and D.Deleted=0 and DP.Deleted=0 and PP.Deleted=0 and D.DeliverableTypeID=8;"
            rst4 = MyCommon.LRT_Select
            ProgramID = ""
            ProgramName = ""
            Amount = 0
            For Each row4 In rst4.Rows
              ProgramID = row4.Item("ProgramID")
              ProgramName = MyCommon.NZ(row4.Item("ProgramName"), "unknown").ToString.Replace(",", " ")
            Next

            If (ProgramName <> "" And ProgramID <> "") Then
              'Find the balance in the points program
              For Each row4 In rst4.Rows
                ProgID = MyCommon.NZ(row4.Item("ProgramID"), -1)
                If row4.Item("ExternalProgram") = True And row4.Item("ExtHostTypeID") = 2 Then
                  'This is an external points program ... get the balance via the web service
                  Balances = MyLookup.GetCustomerExternalPointsBalances(CustomerPK, False, BalRetCode)
                  WriteDebug("Send_XMLCurrentOffers - MyLookup.GetCustomerExternalPointsBalances CustomerPK=" & CustomerPK.ToString, DebugState.CurrentTime)
                  If ((BalRetCode = RETURN_CODE.OK) And (Balances.Length > 0)) Then
                    Amount = Balances(0).Balance
                  Else
                    Amount = 0
                  End If

                Else
                  MyCommon.QueryStr = "select Amount from Points with (NoLock) where CustomerPK=" & CustomerPK & " and ProgramID=" & ProgID
                  rst5 = MyCommon.LXS_Select
                  If (rst5.Rows.Count > 0) Then
                    Amount = MyCommon.NZ(rst5.Rows(0).Item("Amount"), 0)
                  Else
                    Amount = 0
                  End If
                End If
              Next
            End If

            PrintedMessage = RetrievePrintedMessage(OfferID)
            GraphicsFileName = RetrieveGraphicPath(OfferID)
            AccumulationBalance = RetrieveAccumulationBalance(OfferID, CustomerPK)

            'Put the resulting data into a table
            row = dtOffers.NewRow()
            row.Item("OfferID") = OfferID
            row.Item("Name") = OfferName
            row.Item("Description") = OfferDesc
            row.Item("OfferCategoryID") = OfferCategoryID
            row.Item("StartDate") = Date.Parse(OfferStart)
            row.Item("EndDate") = Date.Parse(OfferEnd)
            If (MyCommon.Fetch_CM_SystemOption(85) = "1") Then
              If DispStartDate <> Nothing Then
                row.Item("DisplayStartDate") = Date.Parse(DispStartDate)
              End If
              If DispEndDate <> Nothing Then
                row.Item("DisplayEndDate") = Date.Parse(DispEndDate)
              End If
            End If
            row.Item("CustomerGroupID") = CustomerGroupID
            row.Item("AllowOptOut") = AllowOptOut
            row.Item("EmployeesOnly") = EmployeesOnly
            row.Item("EmployeesExcluded") = EmployeesExcluded
            row.Item("Points") = Amount
            row.Item("Accumulation") = AccumulationBalance
            row.Item("BodyText") = PrintedMessage
            row.Item("Graphic") = GraphicsFileName
	          row.Item("ClientOfferID") = ClientOfferID
            dtOffers.Rows.Add(row)
            ProgramID = ""
            ProgramName = ""
          End If
        Next
      End If
      If dtOffers.Rows.Count > 0 Then dtOffers.AcceptChanges()
    End If

    If MyCommon.PMRTadoConn.State = ConnectionState.Open Then MyCommon.Close_PrefManRT()
    Return dtOffers
  End Function

    Private Function Send_XMLGroupOffers_FilteredResult(ByVal CustomerPK As Long) As DataTable
        Dim rst As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim rst2 As New DataTable()
        Dim row2 As System.Data.DataRow
        Dim rst3 As New DataTable()
        Dim rstTemp As New DataTable()
        Dim rstWeb As System.Data.DataTable
        Dim rstCG As System.Data.DataTable
        Dim cgXml As String
        Dim endDate As String
        Dim reader As SqlDataReader = Nothing
        Dim CurrentOffersClause As String = ""
        Dim CPECurrentOffers As String = ""
        Dim CustomerGroups As New StringBuilder()
        Dim CustomerGroupID As Long
        Dim RewardGroupID As Long
        Dim ROID As Long
        Dim Employee As Boolean
        Dim ExtCustomerID As String = ""
        Dim OfferID As Integer
        Dim CPEClientOfferID As String = ""
        Dim OfferName As String
        Dim OfferDesc As String
        Dim OfferCategoryID As Integer
        Dim OfferStart As Date
        Dim OfferEnd As Date
        Dim OfferDaysLeft As Integer
        Dim GraphicsFileName As String = ""
        Dim PrintedMessage As String = ""
        Dim AccumulationBalance As String = ""
        Dim AllowOptOut As Boolean = False
        Dim OptOutOffer As Boolean = False
        Dim EmployeesOnly As Integer
        Dim EmployeesExcluded As Integer
        Dim ExcludedFromOffer As Boolean = False
        Dim PointsConditionOK As Boolean = True
        Dim QtyRequired As Integer = 0
        Dim ProgramID As String = ""
        Dim ProgramName As String = ""
        Dim rowCount As Integer = 0
        Dim i As Integer
        Dim retstring As New StringBuilder
        Dim Handheld As Boolean = False
        Dim CustomerTypeID As Integer = 0
        Dim dtGroups As New System.Data.DataTable
        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
        If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
        
        'Create a new datatable to hold the results we'll be assembling
        dtGroups = New DataTable
       
        Dim GroupSelectQuery As New StringBuilder
        'Add the aliases with the column names.
        Dim GroupColList As New Dictionary(Of String, String)
        GroupColList.Add("OfferID", "I.IncentiveID")
        GroupColList.Add("CPEClientOfferID", "I.ClientOfferID as CPEClientOfferID")
        GroupColList.Add("Name", "I.IncentiveName")
        GroupColList.Add("Description", "I.Description")
        GroupColList.Add("OfferCategoryID", "I.PromoClassID as OfferCategoryID")
        GroupColList.Add("StartDate", "I.StartDate")
        GroupColList.Add("EndDate", "I.EndDate")
        GroupColList.Add("CustomerGroupID", "ICG.CustomerGroupID")
        GroupColList.Add("AllowOptOut", "I.AllowOptOut")
        GroupColList.Add("EmployeesOnly", "I.EmployeesOnly")
        GroupColList.Add("EmployeesExcluded", "I.EmployeesExcluded")
        
        'Constructing the select columns which should be sent to SP as per user filter
        Dim GroupDtColumns As String()
        MyCommon.QueryStr = "SELECT FilteredAttributes FROM FilterOutputColumns with (NoLock) WHERE ReferenceId = 3"
        Dim dt As DataTable = MyCommon.LRT_Select
        If (dt.Rows.Count > 0) Then
            For Each row2 In dt.Rows
                Dim filterstr As String = MyCommon.NZ(row2.Item("FilteredAttributes"), "")
                Dim sArr As String() = filterstr.Split("|")
                For Each s As String In sArr
                    Dim sArr1 As String() = s.Split("-")
                    Dim tableName As String = sArr1(0)
                    Dim selectedColumns As String() = sArr1(1).Split(",")
                    Select Case (tableName)
                        Case "Groups"
                            GroupDtColumns = selectedColumns
                            If (GroupDtColumns.Contains("OfferID")) Then
                                dtGroups.Columns.Add("OfferID", System.Type.GetType("System.Int64"))
                                If (GroupColList.ContainsKey("OfferID")) Then
                                    GroupSelectQuery.Append(GroupColList("OfferID")).Append(",")
                                End If
                            End If
                            If (GroupDtColumns.Contains("Name")) Then
                                dtGroups.Columns.Add("Name", System.Type.GetType("System.String"))
                                If (GroupColList.ContainsKey("Name")) Then
                                    GroupSelectQuery.Append(GroupColList("Name")).Append(",")
                                End If
                            End If
                            If (GroupDtColumns.Contains("Description")) Then
                                dtGroups.Columns.Add("Description", System.Type.GetType("System.String"))
                                If (GroupColList.ContainsKey("Description")) Then
                                    GroupSelectQuery.Append(GroupColList("Description")).Append(",")
                                End If
                            End If
                            If (GroupDtColumns.Contains("OfferCategoryID")) Then
                                dtGroups.Columns.Add("OfferCategoryID", System.Type.GetType("System.Int32"))
                                If (GroupColList.ContainsKey("OfferCategoryID")) Then
                                    GroupSelectQuery.Append(GroupColList("OfferCategoryID")).Append(",")
                                End If
                            End If
                            If (GroupDtColumns.Contains("StartDate")) Then
                                dtGroups.Columns.Add("StartDate", System.Type.GetType("System.DateTime"))
                                If (GroupColList.ContainsKey("StartDate")) Then
                                    GroupSelectQuery.Append(GroupColList("StartDate")).Append(",")
                                End If
                            End If
                            
                            If (GroupDtColumns.Contains("EndDate")) Then
                                dtGroups.Columns.Add("EndDate", System.Type.GetType("System.DateTime"))
                                If (GroupColList.ContainsKey("EndDate")) Then
                                    GroupSelectQuery.Append(GroupColList("EndDate")).Append(",")
                                End If
                            End If
                            If (GroupDtColumns.Contains("CustomerGroupID")) Then
                                dtGroups.Columns.Add("CustomerGroupID", System.Type.GetType("System.Int64"))
                                If (GroupColList.ContainsKey("CustomerGroupID")) Then
                                    GroupSelectQuery.Append(GroupColList("CustomerGroupID")).Append(",")
                                End If
                            End If
                            If (GroupDtColumns.Contains("AllowOptOut")) Then
                                dtGroups.Columns.Add("AllowOptOut", System.Type.GetType("System.Boolean"))
                                If (GroupColList.ContainsKey("AllowOptOut")) Then
                                    GroupSelectQuery.Append(GroupColList("AllowOptOut")).Append(",")
                                End If
                            End If
                            If (GroupDtColumns.Contains("EmployeesOnly")) Then
                                dtGroups.Columns.Add("EmployeesOnly", System.Type.GetType("System.Boolean"))
                                If (GroupColList.ContainsKey("EmployeesOnly")) Then
                                    GroupSelectQuery.Append(GroupColList("EmployeesOnly")).Append(",")
                                End If
                            End If
                            If (GroupDtColumns.Contains("EmployeesExcluded")) Then
                                dtGroups.Columns.Add("EmployeesExcluded", System.Type.GetType("System.Boolean"))
                                If (GroupColList.ContainsKey("EmployeesExcluded")) Then
                                    GroupSelectQuery.Append(GroupColList("EmployeesExcluded")).Append(",")
                                End If
                            End If
                            If (GroupDtColumns.Contains("BodyText")) Then
                                dtGroups.Columns.Add("BodyText", System.Type.GetType("System.String"))
                                If (GroupColList.ContainsKey("BodyText")) Then
                                    GroupSelectQuery.Append(GroupColList("BodyText")).Append(",")
                                End If
                            End If
                            If (GroupDtColumns.Contains("Graphic")) Then
                                dtGroups.Columns.Add("Graphic", System.Type.GetType("System.String"))
                                If (GroupColList.ContainsKey("Graphic")) Then
                                    GroupSelectQuery.Append(GroupColList("Graphic")).Append(",")
                                End If
                            End If
                            If (GroupDtColumns.Contains("OptType")) Then
                                dtGroups.Columns.Add("OptType", System.Type.GetType("System.String"))
                                If (GroupColList.ContainsKey("OptType")) Then
                                    GroupSelectQuery.Append(GroupColList("OptType")).Append(",")
                                End If
                            End If
                            If (GroupDtColumns.Contains("CPEClientOfferID")) Then
                                dtGroups.Columns.Add("CPEClientOfferID", System.Type.GetType("System.String"))
                                If (GroupColList.ContainsKey("CPEClientOfferID")) Then
                                    GroupSelectQuery.Append(GroupColList("CPEClientOfferID")).Append(",")
                                End If
                            End If
                        Case Else
                    End Select
                Next
            Next
        End If
       

        'First check to see if there's an identifier in the URL
        If (CustomerPK > 0) Then

            'There is, so find the customer's information
            MyCommon.QueryStr = "select C.CustomerPK, C.Employee, C.CustomerStatusID as CardStatusID, CE.Email, CustomerTypeID " & _
                                "from Customers as C with (NoLock) " & _
                                "left join CustomerExt as CE with (NoLock) on CE.CustomerPK=C.CustomerPK " & _
                                "where C.CustomerPK=" & CustomerPK & ";"
            rst = MyCommon.LXS_Select
            If (rst.Rows.Count > 0) Then
                'A customer was found, so assign values to variables
                CustomerPK = rst.Rows(0).Item("CustomerPK")
                CustomerTypeID = MyCommon.NZ(rst.Rows(0).Item("CustomerTypeID"), 0)
                Employee = rst.Rows(0).Item("Employee")

                'Next, get the associated customer groups
                MyCommon.QueryStr = "select CustomerGroupID from GroupMembership with (NoLock) where CustomerPK=" & CustomerPK & " and Deleted=0;"
                rstCG = MyCommon.LXS_Select()
    
                cgXml = "<customergroups><id>1</id><id>2</id>"
                If rstCG.Rows.Count > 0 Then
                    For Each row In rstCG.Rows
                        cgXml &= "<id>" & MyCommon.NZ(row.Item("CustomerGroupID"), "") & "</id>"
                    Next
                End If
                cgXml &= "</customergroups>"
        
                endDate = Today.AddDays(-1).ToString
        
                WriteDebug("Customer Groups FilteredResult: " & cgXml, DebugState.CurrentTime)
                WriteDebug("End Date: " & endDate, DebugState.CurrentTime)
        
                MyCommon.QueryStr = "dbo.pa_CustomerWebOffersOptIn_Filter"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@cgXML", SqlDbType.Xml).Value = cgXml
                MyCommon.LRTsp.Parameters.Add("@endDate", SqlDbType.DateTime).Value = endDate
                MyCommon.LRTsp.Parameters.Add("@filterQuery", SqlDbType.NVarChar).Value = GroupSelectQuery.ToString.Trim(",")
        
                WriteDebug("Executing OptIn Query [LogixRT]", DebugState.CurrentTime)
                reader = MyCommon.LRTsp.ExecuteReader
                WriteDebug("Completed OptIn Query [LogixRT]", DebugState.CurrentTime)
        
                Try
                    rst2.Load(reader)
                Catch ex As Exception
                    WriteDebug("Exception " & ex.GetType.Name & ":" & ex.Message(), DebugState.CurrentTime)
                End Try
                WriteDebug("Completed Load of OptIn Query with row count=" & rst2.Rows.Count, DebugState.CurrentTime)
 
                MyCommon.QueryStr = "dbo.pa_CustomerWebOffersOptOut_Filter"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@cgXML", SqlDbType.Xml).Value = cgXml
                MyCommon.LRTsp.Parameters.Add("@endDate", SqlDbType.DateTime).Value = endDate
                MyCommon.LRTsp.Parameters.Add("@filterQuery", SqlDbType.NVarChar).Value = GroupSelectQuery.ToString.Trim(",")
        
                WriteDebug("Executing OptOut Query [LogixRT]", DebugState.CurrentTime)
                reader = MyCommon.LRTsp.ExecuteReader
                WriteDebug("Completed OptOut Query [LogixRT]", DebugState.CurrentTime)
        
                Try
                    rst3.Load(reader)
                Catch ex As Exception
                    WriteDebug("Exception " & ex.GetType.Name & ":" & ex.Message(), DebugState.CurrentTime)
                End Try
                WriteDebug("Completed Load of OptOut Query with row count=" & rst3.Rows.Count, DebugState.CurrentTime)

                'Set the general info for each offer found
                If (rst2.Rows.Count > 0) Then
                    WriteDebug("Send_XMLGroupOffersFilteredResult - Processing OptIn Offers", DebugState.CurrentTime)
                    For Each row2 In rst2.Rows
                        If (GroupDtColumns.Contains("OfferID")) Then OfferID = row2.Item("IncentiveID")
                        If (GroupDtColumns.Contains("Name")) Then OfferName = row2.Item("IncentiveName")
                        If (GroupDtColumns.Contains("Description")) Then
                            OfferDesc = row2.Item("Description")
                            If (OfferDesc.Trim() = "") Then OfferDesc = OfferName
                        End If
                                    
                        If (GroupDtColumns.Contains("OfferCategoryID")) Then OfferCategoryID = MyCommon.NZ(row2.Item("OfferCategoryID"), 0)
                        If (GroupDtColumns.Contains("StartDate")) Then OfferStart = row2.Item("StartDate")
                        If (GroupDtColumns.Contains("EndDate")) Then
                            OfferEnd = row2.Item("EndDate")
                            OfferDaysLeft = DateDiff("d", Today, OfferEnd)
                        End If
                        If (GroupDtColumns.Contains("CustomerGroupID")) Then CustomerGroupID = row2.Item("CustomerGroupID")
                        RewardGroupID = MyCommon.NZ(row2.Item("RewardGroup"), -1)
                        If (GroupDtColumns.Contains("AllowOptOut")) Then AllowOptOut = MyCommon.NZ(row2.Item("AllowOptOut"), False)
                        If (GroupDtColumns.Contains("EmployeesOnly")) Then EmployeesOnly = IIf(MyCommon.NZ(row2.Item("EmployeesOnly"), False), 1, 0)
                        If (GroupDtColumns.Contains("EmployeesExcluded")) Then EmployeesExcluded = IIf(MyCommon.NZ(row2.Item("EmployeesExcluded"), False), 1, 0)

                        ROID = MyCommon.NZ(row2.Item("RewardOptionId"), -1)
            
                        OptOutOffer = False

                        'If enabled, test the first Incentive related to this Website Offer's Group Membership Customer Group ID for limits.  If blown, stop processing and continue to the next offer, consider thie Website Offer as ineligible
                        If (MyCommon.Fetch_CPE_SystemOption(198) = "1" And RewardGroupID <> -1) Then
                            MyCommon.QueryStr = "select i.incentiveid , i.ClientOfferID " & _
                                  "from CPE_ST_Incentives i with (nolock) " & _
                                  "inner join CPE_ST_RewardOptions as r with (nolock) on i.incentiveid=r.incentiveid " & _
                                  "inner join CPE_ST_IncentiveCustomerGroups CG1 with (nolock) on CG1.rewardoptionid=r.rewardoptionid " & _
                                  "where CG1.customergroupid = " & RewardGroupID & " and CG1.ExcludedUsers <> 1 order by i.incentiveid desc;"
                            rstTemp = MyCommon.LRT_Select()
                            'WriteDebug("RewardGroupID:"&RewardGroupID, DebugState.CurrentTime)
                            If (rstTemp.Rows.Count > 0) Then
                                If (GroupDtColumns.Contains("CPEClientOfferID")) Then CPEClientOfferID = MyCommon.NZ(rstTemp.Rows(0).Item("ClientOfferID"), "")
                                'WriteDebug("Test Limit:"&rstTemp.Rows(0).Item("incentiveID")&":"&CustomerPK, DebugState.CurrentTime)
                                If (CPELimitsBlown(rstTemp.Rows(0).Item("incentiveID"), CustomerPK)) Then
                                    WriteDebug("Test Limit:Limit Blown, skip...", DebugState.CurrentTime)
                                    Continue For
                                End If
                            End If
                        End If

                        'Check if the customer meets any points condition that may exist for this offer
                        MyCommon.QueryStr = "select QtyForIncentive, ProgramID from CPE_IncentivePointsGroups with (NoLock) where RewardOptionID=" & ROID & " and Deleted=0;"
                        rstWeb = MyCommon.LRT_Select
                        If (rstWeb.Rows.Count > 0) Then
                            ' check if the customer has enough points in the program
                            QtyRequired = MyCommon.NZ(rstWeb.Rows(0).Item("QtyForIncentive"), 0)
                            MyCommon.QueryStr = "select Amount from Points with (NoLock) where ProgramID=" & MyCommon.NZ(rstWeb.Rows(0).Item("ProgramID"), -1) & " and CustomerPK=" & CustomerPK & ";"
                            rstWeb = MyCommon.LXS_Select
                            If (rstWeb.Rows.Count > 0) Then
                                PointsConditionOK = (MyCommon.NZ(rstWeb.Rows(0).Item("Amount"), 0) >= QtyRequired)
                            Else
                                'No points entry for this customer
                                PointsConditionOK = False
                            End If
                        Else
                            'No points conditions exist for this web offer
                            PointsConditionOK = True
                        End If

                        If (PointsConditionOK) Then

                            If (GroupDtColumns.Contains("BodyText")) Then PrintedMessage = RetrievePrintedMessage(OfferID)
                            If (GroupDtColumns.Contains("Graphic")) Then GraphicsFileName = RetrieveGraphicPath(OfferID)

                            'Put the resulting data into a table
                            row = dtGroups.NewRow()
                            If (GroupDtColumns.Contains("OfferID")) Then row.Item("OfferID") = OfferID
                            If (GroupDtColumns.Contains("Name")) Then row.Item("Name") = OfferName
                            If (GroupDtColumns.Contains("Description")) Then row.Item("Description") = OfferDesc
                            If (GroupDtColumns.Contains("OfferCategoryID")) Then row.Item("OfferCategoryID") = OfferCategoryID
                            If (GroupDtColumns.Contains("StartDate")) Then row.Item("StartDate") = Date.Parse(OfferStart)
                            If (GroupDtColumns.Contains("EndDate")) Then row.Item("EndDate") = Date.Parse(OfferEnd)
                            If (GroupDtColumns.Contains("CustomerGroupID")) Then row.Item("CustomerGroupID") = RewardGroupID
                            If (GroupDtColumns.Contains("AllowOptOut")) Then row.Item("AllowOptOut") = AllowOptOut
                            If (GroupDtColumns.Contains("EmployeesOnly")) Then row.Item("EmployeesOnly") = EmployeesOnly
                            If (GroupDtColumns.Contains("EmployeesExcluded")) Then row.Item("EmployeesExcluded") = EmployeesExcluded
                            If (GroupDtColumns.Contains("BodyText")) Then row.Item("BodyText") = PrintedMessage
                            If (GroupDtColumns.Contains("Graphic")) Then row.Item("Graphic") = GraphicsFileName
                            If (GroupDtColumns.Contains("OptType")) Then row.Item("OptType") = IIf(OptOutOffer, "out", "in")
                            If (GroupDtColumns.Contains("CPEClientOfferID")) Then row.Item("CPEClientOfferID") = CPEClientOfferID
                            dtGroups.Rows.Add(row)
                        End If
                    Next
                    WriteDebug("Send_XMLGroupOffersFilteredResult - Completed OptIn Offers", DebugState.CurrentTime)
                    If dtGroups.Rows.Count > 0 Then dtGroups.AcceptChanges()
                End If
        
                'Set the general info for each offer found
                If (rst3.Rows.Count > 0) Then
                    WriteDebug("Send_XMLGroupOffersFilteredResult - Processing OptOut Offers", DebugState.CurrentTime)
                    For Each row2 In rst3.Rows
                        If (GroupDtColumns.Contains("OfferID")) Then OfferID = row2.Item("IncentiveID")
                        If (GroupDtColumns.Contains("Name")) Then OfferName = row2.Item("IncentiveName")
                        If (GroupDtColumns.Contains("Description")) Then
                            OfferDesc = row2.Item("Description")
                            If (OfferDesc.Trim() = "") Then OfferDesc = OfferName
                        End If
                                    
                        If (GroupDtColumns.Contains("OfferCategoryID")) Then OfferCategoryID = MyCommon.NZ(row2.Item("OfferCategoryID"), 0)
                        If (GroupDtColumns.Contains("StartDate")) Then OfferStart = row2.Item("StartDate")
                        If (GroupDtColumns.Contains("EndDate")) Then
                            OfferEnd = row2.Item("EndDate")
                            OfferDaysLeft = DateDiff("d", Today, OfferEnd)
                        End If
                        If (GroupDtColumns.Contains("CustomerGroupID")) Then CustomerGroupID = row2.Item("CustomerGroupID")
                        RewardGroupID = MyCommon.NZ(row2.Item("RewardGroup"), -1)
                        If (GroupDtColumns.Contains("AllowOptOut")) Then AllowOptOut = MyCommon.NZ(row2.Item("AllowOptOut"), False)
                        If (GroupDtColumns.Contains("EmployeesOnly")) Then EmployeesOnly = IIf(MyCommon.NZ(row2.Item("EmployeesOnly"), False), 1, 0)
                        If (GroupDtColumns.Contains("EmployeesExcluded")) Then EmployeesExcluded = IIf(MyCommon.NZ(row2.Item("EmployeesExcluded"), False), 1, 0)

                        ROID = MyCommon.NZ(row2.Item("RewardOptionId"), -1)
            
                        OptOutOffer = True

                        ROID = MyCommon.NZ(row2.Item("RewardOptionId"), -1)
                        'If enabled, test the first Incentive related to this Website Offer's Group Membership Customer Group ID for limits.  If blown, stop processing and continue to the next offer, consider thie Website Offer as ineligible
                        If (MyCommon.Fetch_CPE_SystemOption(198) = "1" And RewardGroupID <> -1) Then
                            MyCommon.QueryStr = "select i.incentiveid , i.ClientOfferID " & _
                                  "from CPE_ST_Incentives i with (nolock) " & _
                                  "inner join CPE_ST_RewardOptions as r with (nolock) on i.incentiveid=r.incentiveid " & _
                                  "inner join CPE_ST_IncentiveCustomerGroups CG1 with (nolock) on CG1.rewardoptionid=r.rewardoptionid " & _
                                  "where CG1.customergroupid = " & RewardGroupID & " and CG1.ExcludedUsers <> 1 order by i.incentiveid desc;"
                            rstTemp = MyCommon.LRT_Select()
                            'WriteDebug("RewardGroupID:"&RewardGroupID, DebugState.CurrentTime)
                            If (rstTemp.Rows.Count > 0) Then
                                If (GroupDtColumns.Contains("CPEClientOfferID")) Then CPEClientOfferID = MyCommon.NZ(rstTemp.Rows(0).Item("ClientOfferID"), "")
                                'WriteDebug("Test Limit:"&rstTemp.Rows(0).Item("incentiveID")&":"&CustomerPK, DebugState.CurrentTime)
                                If (CPELimitsBlown(rstTemp.Rows(0).Item("incentiveID"), CustomerPK)) Then
                                    WriteDebug("Test Limit:Limit Blown, skip...", DebugState.CurrentTime)
                                    Continue For
                                End If
                            End If
                        End If

                        'Check if the customer meets any points condition that may exist for this offer
                        MyCommon.QueryStr = "select QtyForIncentive, ProgramID from CPE_IncentivePointsGroups with (NoLock) where RewardOptionID=" & ROID & " and Deleted=0;"
                        rstWeb = MyCommon.LRT_Select
                        If (rstWeb.Rows.Count > 0) Then
                            ' check if the customer has enough points in the program
                            QtyRequired = MyCommon.NZ(rstWeb.Rows(0).Item("QtyForIncentive"), 0)
                            MyCommon.QueryStr = "select Amount from Points with (NoLock) where ProgramID=" & MyCommon.NZ(rstWeb.Rows(0).Item("ProgramID"), -1) & " and CustomerPK=" & CustomerPK & ";"
                            rstWeb = MyCommon.LXS_Select
                            If (rstWeb.Rows.Count > 0) Then
                                PointsConditionOK = (MyCommon.NZ(rstWeb.Rows(0).Item("Amount"), 0) >= QtyRequired)
                            Else
                                'No points entry for this customer
                                PointsConditionOK = False
                            End If
                        Else
                            'No points conditions exist for this web offer
                            PointsConditionOK = True
                        End If

                        If (PointsConditionOK) Then

                            If (GroupDtColumns.Contains("BodyText")) Then PrintedMessage = RetrievePrintedMessage(OfferID)
                            If (GroupDtColumns.Contains("Graphic")) Then GraphicsFileName = RetrieveGraphicPath(OfferID)

                            'Put the resulting data into a table
                            row = dtGroups.NewRow()
                            If (GroupDtColumns.Contains("OfferID")) Then row.Item("OfferID") = OfferID
                            If (GroupDtColumns.Contains("Name")) Then row.Item("Name") = OfferName
                            If (GroupDtColumns.Contains("Description")) Then row.Item("Description") = OfferDesc
                            If (GroupDtColumns.Contains("OfferCategoryID")) Then row.Item("OfferCategoryID") = OfferCategoryID
                            If (GroupDtColumns.Contains("StartDate")) Then row.Item("StartDate") = Date.Parse(OfferStart)
                            If (GroupDtColumns.Contains("EndDate")) Then row.Item("EndDate") = Date.Parse(OfferEnd)
                            If (GroupDtColumns.Contains("CustomerGroupID")) Then row.Item("CustomerGroupID") = RewardGroupID
                            If (GroupDtColumns.Contains("AllowOptOut")) Then row.Item("AllowOptOut") = AllowOptOut
                            If (GroupDtColumns.Contains("EmployeesOnly")) Then row.Item("EmployeesOnly") = EmployeesOnly
                            If (GroupDtColumns.Contains("EmployeesExcluded")) Then row.Item("EmployeesExcluded") = EmployeesExcluded
                            If (GroupDtColumns.Contains("BodyText")) Then row.Item("BodyText") = PrintedMessage
                            If (GroupDtColumns.Contains("Graphic")) Then row.Item("Graphic") = GraphicsFileName
                            If (GroupDtColumns.Contains("OptType")) Then row.Item("OptType") = IIf(OptOutOffer, "out", "in")
                            If (GroupDtColumns.Contains("CPEClientOfferID")) Then row.Item("CPEClientOfferID") = CPEClientOfferID
                            dtGroups.Rows.Add(row)
                        End If
                    Next
                    WriteDebug("Send_XMLGroupOffersFilteredResult - Completed OptOut Offers", DebugState.CurrentTime)
                    If dtGroups.Rows.Count > 0 Then dtGroups.AcceptChanges()
                End If
            End If
        End If

        Return dtGroups
    End Function
    
  Private Function Send_XMLGroupOffers(ByVal CustomerPK As Long) As DataTable
    Dim rst As System.Data.DataTable
    Dim row As System.Data.DataRow
    Dim rst2 As New DataTable()
    Dim row2 As System.Data.DataRow
    Dim rst3 As New DataTable()
    Dim rstTemp As New DataTable()
    Dim rstWeb As System.Data.DataTable
    Dim rstCG As System.Data.DataTable
    Dim cgXml As String
    Dim endDate As String
    Dim reader As SqlDataReader = Nothing
    Dim CurrentOffersClause As String = ""
    Dim CPECurrentOffers As String = ""
    Dim CustomerGroups As New StringBuilder()
    Dim CustomerGroupID As Long
    Dim RewardGroupID As Long
    Dim ROID As Long
    Dim Employee As Boolean
    Dim ExtCustomerID As String = ""
    Dim OfferID As Integer
    Dim OfferName As String
    Dim OfferDesc As String
    Dim OfferCategoryID As Integer
    Dim OfferStart As Date
    Dim OfferEnd As Date
    Dim OfferDaysLeft As Integer
    Dim GraphicsFileName As String = ""
    Dim PrintedMessage As String = ""
    Dim AccumulationBalance As String = ""
    Dim AllowOptOut As Boolean = False
    Dim OptOutOffer As Boolean = False
    Dim EmployeesOnly As Integer
    Dim EmployeesExcluded As Integer
    Dim ExcludedFromOffer As Boolean = False
    Dim PointsConditionOK As Boolean = True
    Dim QtyRequired As Integer = 0
    Dim ProgramID As String = ""
    Dim ProgramName As String = ""
    Dim rowCount As Integer = 0
    Dim i As Integer
    Dim retstring As New StringBuilder
    Dim Handheld As Boolean = False
    Dim CustomerTypeID As Integer = 0
    Dim dtGroups As New System.Data.DataTable
    Dim CPEClientOfferID As String = ""

    If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
    If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
    
    'Create a new datatable to hold the results we'll be assembling
    dtGroups = New DataTable
    dtGroups.Columns.Add("OfferID", System.Type.GetType("System.Int64"))
    dtGroups.Columns.Add("Name", System.Type.GetType("System.String"))
    dtGroups.Columns.Add("Description", System.Type.GetType("System.String"))
    dtGroups.Columns.Add("OfferCategoryID", System.Type.GetType("System.Int32"))
    dtGroups.Columns.Add("StartDate", System.Type.GetType("System.DateTime"))
    dtGroups.Columns.Add("EndDate", System.Type.GetType("System.DateTime"))
    dtGroups.Columns.Add("CustomerGroupID", System.Type.GetType("System.Int64"))
    dtGroups.Columns.Add("AllowOptOut", System.Type.GetType("System.Boolean"))
    dtGroups.Columns.Add("EmployeesOnly", System.Type.GetType("System.Boolean"))
    dtGroups.Columns.Add("EmployeesExcluded", System.Type.GetType("System.Boolean"))
    dtGroups.Columns.Add("Points", System.Type.GetType("System.Int32"))
    dtGroups.Columns.Add("Accumulation", System.Type.GetType("System.Decimal"))
    dtGroups.Columns.Add("BodyText", System.Type.GetType("System.String"))
    dtGroups.Columns.Add("Graphic", System.Type.GetType("System.String"))
    dtGroups.Columns.Add("OptType", System.Type.GetType("System.String"))
  	dtGroups.Columns.Add("CPEClientOfferID", System.Type.GetType("System.String"))

    'First check to see if there's an identifier in the URL
    If (CustomerPK > 0) Then

      'There is, so find the customer's information
      MyCommon.QueryStr = "select C.CustomerPK, C.Employee, C.CustomerStatusID as CardStatusID, CE.Email, CustomerTypeID " & _
                          "from Customers as C with (NoLock) " & _
                          "left join CustomerExt as CE with (NoLock) on CE.CustomerPK=C.CustomerPK " & _
                          "where C.CustomerPK=" & CustomerPK & ";"
      rst = MyCommon.LXS_Select
      If (rst.Rows.Count > 0) Then
        'A customer was found, so assign values to variables
        CustomerPK = rst.Rows(0).Item("CustomerPK")
        CustomerTypeID = MyCommon.NZ(rst.Rows(0).Item("CustomerTypeID"), 0)
        Employee = rst.Rows(0).Item("Employee")

        'Next, get the associated customer groups
        MyCommon.QueryStr = "select CustomerGroupID from GroupMembership with (NoLock) where CustomerPK=" & CustomerPK & " and Deleted=0;"
        rstCG = MyCommon.LXS_Select()
    
        cgXml = "<customergroups><id>1</id><id>2</id>"
          If rstCG.Rows.Count > 0 Then
            For Each row In rstCG.Rows
            cgXml &= "<id>" & MyCommon.NZ(row.Item("CustomerGroupID"), "") & "</id>"
            Next
          End If
        cgXml &= "</customergroups>"
        
        endDate = Today.AddDays(-1).ToString
        
        WriteDebug("Customer Groups: " & cgXML, DebugState.CurrentTime)
        WriteDebug("End Date: " & endDate, DebugState.CurrentTime)
        
        MyCommon.QueryStr = "dbo.pa_CustomerWebOffersOptIn"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@cgXML", SqlDbType.Xml).Value = cgXml
        MyCommon.LRTsp.Parameters.Add("@endDate", SqlDbType.DateTime).Value = endDate
        
        WriteDebug("Executing OptIn Query [LogixRT]", DebugState.CurrentTime)
        reader = MyCommon.LRTsp.ExecuteReader
        WriteDebug("Completed OptIn Query [LogixRT]", DebugState.CurrentTime)
        
        Try
          rst2.Load(reader)
        Catch ex As Exception
          WriteDebug("Exception " & ex.GetType.Name & ":" & ex.Message(), DebugState.CurrentTime)
        End Try
        WriteDebug("Completed Load of OptIn Query with row count=" & rst2.Rows.Count, DebugState.CurrentTime)         
 
        MyCommon.QueryStr = "dbo.pa_CustomerWebOffersOptOut"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@cgXML", SqlDbType.Xml).Value = cgXml
        MyCommon.LRTsp.Parameters.Add("@endDate", SqlDbType.DateTime).Value = endDate
        
        WriteDebug("Executing OptOut Query [LogixRT]", DebugState.CurrentTime)
        reader = MyCommon.LRTsp.ExecuteReader
        WriteDebug("Completed OptOut Query [LogixRT]", DebugState.CurrentTime)
        
        Try
          rst3.Load(reader)
        Catch ex As Exception
          WriteDebug("Exception " & ex.GetType.Name & ":" & ex.Message(), DebugState.CurrentTime)
        End Try
        WriteDebug("Completed Load of OptOut Query with row count=" & rst3.Rows.Count, DebugState.CurrentTime)

        'Set the general info for each offer found
        If (rst2.Rows.Count > 0) Then
          WriteDebug("Send_XMLGroupOffers - Processing OptIn Offers", DebugState.CurrentTime)
          For Each row2 In rst2.Rows
            OfferID = row2.Item("IncentiveID")
            OfferName = row2.Item("IncentiveName")
            OfferDesc = row2.Item("Description")
            If (OfferDesc.Trim() = "") Then OfferDesc = OfferName
            OfferCategoryID = MyCommon.NZ(row2.Item("OfferCategoryID"), 0)
            OfferStart = row2.Item("StartDate")
            OfferEnd = row2.Item("EndDate")
            OfferDaysLeft = DateDiff("d", Today, OfferEnd)
            CustomerGroupID = row2.Item("CustomerGroupID")
            RewardGroupID = MyCommon.NZ(row2.Item("RewardGroup"), -1)
            AllowOptOut = MyCommon.NZ(row2.Item("AllowOptOut"), False)
            EmployeesOnly = IIf(MyCommon.NZ(row2.Item("EmployeesOnly"), False), 1, 0)
            EmployeesExcluded = IIf(MyCommon.NZ(row2.Item("EmployeesExcluded"), False), 1, 0)

            ROID = MyCommon.NZ(row2.Item("RewardOptionId"), -1)
            
            OptOutOffer = false

            'If enabled, test the first Incentive related to this Website Offer's Group Membership Customer Group ID for limits.  If blown, stop processing and continue to the next offer, consider thie Website Offer as ineligible
            If(MyCommon.Fetch_CPE_SystemOption(198) = "1" And RewardGroupID<>-1)
              MyCommon.QueryStr = "select i.incentiveid, i.ClientOfferID " & _
                    "from CPE_ST_Incentives i with (nolock) " & _
                    "inner join CPE_ST_RewardOptions as r with (nolock) on i.incentiveid=r.incentiveid " & _
                    "inner join CPE_ST_IncentiveCustomerGroups CG1 with (nolock) on CG1.rewardoptionid=r.rewardoptionid " & _
                    "where CG1.customergroupid = " & RewardGroupID & " and CG1.ExcludedUsers <> 1 order by i.incentiveid desc;"
              rstTemp=MyCommon.LRT_Select()
              'WriteDebug("RewardGroupID:"&RewardGroupID, DebugState.CurrentTime)
              If (rstTemp.Rows.Count>0) Then
			    CPEClientOfferID = MyCommon.NZ(rstTemp.Rows(0).Item("ClientOfferID"),"")
                'WriteDebug("Test Limit:"&rstTemp.Rows(0).Item("incentiveID")&":"&CustomerPK, DebugState.CurrentTime)
                If(CPELimitsBlown(rstTemp.Rows(0).Item("incentiveID"),CustomerPK)) Then
                  WriteDebug("Test Limit:Limit Blown, skip...", DebugState.CurrentTime)
                  Continue For
                End If
              End If
            End If
            'Check if the customer meets any points condition that may exist for this offer
            MyCommon.QueryStr = "select QtyForIncentive, ProgramID from CPE_IncentivePointsGroups with (NoLock) where RewardOptionID=" & ROID & " and Deleted=0;"
            rstWeb = MyCommon.LRT_Select
            If (rstWeb.Rows.Count > 0) Then
              ' check if the customer has enough points in the program
              QtyRequired = MyCommon.NZ(rstWeb.Rows(0).Item("QtyForIncentive"), 0)
              MyCommon.QueryStr = "select Amount from Points with (NoLock) where ProgramID=" & MyCommon.NZ(rstWeb.Rows(0).Item("ProgramID"), -1) & " and CustomerPK=" & CustomerPK & ";"
              rstWeb = MyCommon.LXS_Select
              If (rstWeb.Rows.Count > 0) Then
                PointsConditionOK = (MyCommon.NZ(rstWeb.Rows(0).Item("Amount"), 0) >= QtyRequired)
              Else
                'No points entry for this customer
                PointsConditionOK = False
              End If
            Else
              'No points conditions exist for this web offer
              PointsConditionOK = True
            End If

            If (PointsConditionOK) Then

              PrintedMessage = RetrievePrintedMessage(OfferID)
              GraphicsFileName = RetrieveGraphicPath(OfferID)

              'Put the resulting data into a table
              row = dtGroups.NewRow()
              row.Item("OfferID") = OfferID
              row.Item("Name") = OfferName
              row.Item("Description") = OfferDesc
              row.Item("OfferCategoryID") = OfferCategoryID
              row.Item("StartDate") = Date.Parse(OfferStart)
              row.Item("EndDate") = Date.Parse(OfferEnd)
              row.Item("CustomerGroupID") = RewardGroupID
              row.Item("AllowOptOut") = AllowOptOut
              row.Item("EmployeesOnly") = EmployeesOnly
              row.Item("EmployeesExcluded") = EmployeesExcluded
              row.Item("BodyText") = PrintedMessage
              row.Item("Graphic") = GraphicsFileName
              row.Item("OptType") = IIf(OptOutOffer, "out", "in")
			  row.Item("CPEClientOfferID") = CPEClientOfferID
              dtGroups.Rows.Add(row)
            End If
          Next
          WriteDebug("Send_XMLGroupOffers - Completed OptIn Offers", DebugState.CurrentTime)
          If dtGroups.Rows.Count > 0 Then dtGroups.AcceptChanges()
        End If     
        
        'Set the general info for each offer found
        If (rst3.Rows.Count > 0) Then
          WriteDebug("Send_XMLGroupOffers - Processing OptOut Offers", DebugState.CurrentTime)
          For Each row2 In rst3.Rows
            OfferID = row2.Item("IncentiveID")
            OfferName = row2.Item("IncentiveName")
            OfferDesc = row2.Item("Description")
            If (OfferDesc.Trim() = "") Then OfferDesc = OfferName
            OfferCategoryID = MyCommon.NZ(row2.Item("OfferCategoryID"), 0)
            OfferStart = row2.Item("StartDate")
            OfferEnd = row2.Item("EndDate")
            OfferDaysLeft = DateDiff("d", Today, OfferEnd)
            CustomerGroupID = row2.Item("CustomerGroupID")
            RewardGroupID = MyCommon.NZ(row2.Item("RewardGroup"), -1)
            AllowOptOut = MyCommon.NZ(row2.Item("AllowOptOut"), False)
            EmployeesOnly = IIf(MyCommon.NZ(row2.Item("EmployeesOnly"), False), 1, 0)
            EmployeesExcluded = IIf(MyCommon.NZ(row2.Item("EmployeesExcluded"), False), 1, 0)
            
            OptOutOffer = True

            ROID = MyCommon.NZ(row2.Item("RewardOptionId"), -1)
            'If enabled, test the first Incentive related to this Website Offer's Group Membership Customer Group ID for limits.  If blown, stop processing and continue to the next offer, consider thie Website Offer as ineligible
            If(MyCommon.Fetch_CPE_SystemOption(198) = "1" And RewardGroupID<>-1)
              MyCommon.QueryStr = "select i.incentiveid, i.ClientOfferID " & _
                    "from CPE_ST_Incentives i with (nolock) " & _
                    "inner join CPE_ST_RewardOptions as r with (nolock) on i.incentiveid=r.incentiveid " & _
                    "inner join CPE_ST_IncentiveCustomerGroups CG1 with (nolock) on CG1.rewardoptionid=r.rewardoptionid " & _
                    "where CG1.customergroupid = " & RewardGroupID & " and CG1.ExcludedUsers <> 1 order by i.incentiveid desc;"
              rstTemp=MyCommon.LRT_Select()
              'WriteDebug("RewardGroupID:"&RewardGroupID, DebugState.CurrentTime)
              If (rstTemp.Rows.Count>0) Then
      			CPEClientOfferID = MyCommon.NZ(rstTemp.Rows(0).Item("ClientOfferID"),"")
                'WriteDebug("Test Limit:"&rstTemp.Rows(0).Item("incentiveID")&":"&CustomerPK, DebugState.CurrentTime)
                If(CPELimitsBlown(rstTemp.Rows(0).Item("incentiveID"),CustomerPK)) Then
                  WriteDebug("Test Limit:Limit Blown, skip...", DebugState.CurrentTime)
                  Continue For
                End If
              End If
            End If

            'Check if the customer meets any points condition that may exist for this offer
            MyCommon.QueryStr = "select QtyForIncentive, ProgramID from CPE_IncentivePointsGroups with (NoLock) where RewardOptionID=" & ROID & " and Deleted=0;"
            rstWeb = MyCommon.LRT_Select
            If (rstWeb.Rows.Count > 0) Then
              ' check if the customer has enough points in the program
              QtyRequired = MyCommon.NZ(rstWeb.Rows(0).Item("QtyForIncentive"), 0)
              MyCommon.QueryStr = "select Amount from Points with (NoLock) where ProgramID=" & MyCommon.NZ(rstWeb.Rows(0).Item("ProgramID"), -1) & " and CustomerPK=" & CustomerPK & ";"
              rstWeb = MyCommon.LXS_Select
              If (rstWeb.Rows.Count > 0) Then
                PointsConditionOK = (MyCommon.NZ(rstWeb.Rows(0).Item("Amount"), 0) >= QtyRequired)
              Else
                'No points entry for this customer
                PointsConditionOK = False
              End If
            Else
              'No points conditions exist for this web offer
              PointsConditionOK = True
            End If

            If (PointsConditionOK) Then

              PrintedMessage = RetrievePrintedMessage(OfferID)
              GraphicsFileName = RetrieveGraphicPath(OfferID)

              'Put the resulting data into a table
              row = dtGroups.NewRow()
              row.Item("OfferID") = OfferID
              row.Item("Name") = OfferName
              row.Item("Description") = OfferDesc
              row.Item("OfferCategoryID") = OfferCategoryID
              row.Item("StartDate") = Date.Parse(OfferStart)
              row.Item("EndDate") = Date.Parse(OfferEnd)
              row.Item("CustomerGroupID") = RewardGroupID
              row.Item("AllowOptOut") = AllowOptOut
              row.Item("EmployeesOnly") = EmployeesOnly
              row.Item("EmployeesExcluded") = EmployeesExcluded
              row.Item("BodyText") = PrintedMessage
              row.Item("Graphic") = GraphicsFileName
              row.Item("OptType") = IIf(OptOutOffer, "out", "in")
			  row.Item("CPEClientOfferID") = CPEClientOfferID
              dtGroups.Rows.Add(row)
            End If
          Next
          WriteDebug("Send_XMLGroupOffers - Completed OptOut Offers", DebugState.CurrentTime)
          If dtGroups.Rows.Count > 0 Then dtGroups.AcceptChanges()
        End If
      End If
    End If

    Return dtGroups
  End Function

    Private Function Send_PointsProgramBalances(ByVal CustomerPK As Long, _
                                                ByRef RetCode As StatusCodes, ByRef RetMsg As String) As DataTable
        Dim dtBalances As New DataTable
        Dim rst As DataTable
        Dim row As DataRow
        Dim MyLookup As New Copient.CustomerLookup
        Dim BalRetCode As Copient.CustomerLookup.RETURN_CODE
        Dim Balances(-1) As Copient.CustomerLookup.PointsBalance
        Dim i As Integer

        RetCode = StatusCodes.SUCCESS
        RetMsg = ""

        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
        If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

        'Create a new datatable to hold the results we'll be assembling
        dtBalances = New DataTable("PointsProgram")
        dtBalances.Columns.Add("ProgramID", System.Type.GetType("System.Int64"))
        dtBalances.Columns.Add("Balance", System.Type.GetType("System.Int32"))

        'First check to see if there's an identifier in the URL
        If (CustomerPK > 0) Then
            'There is, so find the customer's information
            MyCommon.QueryStr = "select C.CustomerPK from Customers as C with (NoLock) " & _
                                "where C.CustomerPK=" & CustomerPK & ";"
            rst = MyCommon.LXS_Select

            If (rst.Rows.Count > 0) Then

                'check for internal points balances (stored in LogixXS.Points)
                Balances = MyLookup.GetCustomerPointsBalances(CustomerPK, False, BalRetCode)
                If BalRetCode = RETURN_CODE.OK Then
                    For i = 0 To Balances.GetUpperBound(0)
                        row = dtBalances.NewRow()
                        row.Item("ProgramID") = Balances(i).ProgramID
                        row.Item("Balance") = Balances(i).Balance
                        dtBalances.Rows.Add(row)
                    Next
                    If Balances.GetUpperBound(0) > 0 Then dtBalances.AcceptChanges()
                Else
                    RetCode = StatusCodes.FAILED_BALANCE_LOOKUP
                    RetMsg = "Failed to load customer's points program balances.  Check log for details on this exception."
                End If

                'check for points balances in an external points program
                Balances = MyLookup.GetCustomerExternalPointsBalances(CustomerPK, False, BalRetCode)
                If BalRetCode = RETURN_CODE.OK Then
                    For i = 0 To Balances.GetUpperBound(0)
                        row = dtBalances.NewRow()
                        row.Item("ProgramID") = Balances(i).ProgramID
                        row.Item("Balance") = Balances(i).Balance
                        dtBalances.Rows.Add(row)
                    Next
                    If Balances.GetUpperBound(0) > 0 Then dtBalances.AcceptChanges()
                Else
                    RetCode = StatusCodes.FAILED_BALANCE_LOOKUP
                    RetMsg = "Failed to load customer's external points program balances.  Check log for details on this exception."
                End If

            Else
                RetCode = StatusCodes.INVALID_CUSTOMERID
                RetMsg = "Failure: Invalid customer ID."
            End If
        End If

        Return dtBalances
    End Function

    Private Function Send_PointsProgramBalances_FilteredResponse(ByVal CustomerPK As Long, ByVal selectedResCols As String, _
                                                ByRef RetCode As StatusCodes, ByRef RetMsg As String) As DataTable
        Dim dtBalances As New DataTable
        Dim rst As DataTable
        Dim row As DataRow
        Dim MyLookup As New Copient.CustomerLookup
        Dim BalRetCode As Copient.CustomerLookup.RETURN_CODE
        Dim Balances(-1) As Copient.CustomerLookup.PointsBalance
        Dim i As Integer

        RetCode = StatusCodes.SUCCESS
        RetMsg = ""

        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
        If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

        'Create a new datatable to hold the results we'll be assembling
        dtBalances = New DataTable("PointsProgram")
        If (selectedResCols.Contains("ProgramID")) Then dtBalances.Columns.Add("ProgramID", System.Type.GetType("System.Int64"))
        If (selectedResCols.Contains("Balance")) Then dtBalances.Columns.Add("Balance", System.Type.GetType("System.Int32"))

        'First check to see if there's an identifier in the URL
        If (CustomerPK > 0) Then
            'There is, so find the customer's information
            MyCommon.QueryStr = "select C.CustomerPK from Customers as C with (NoLock) " & _
                                "where C.CustomerPK=" & CustomerPK & ";"
            rst = MyCommon.LXS_Select

            If (rst.Rows.Count > 0) Then

                'check for internal points balances (stored in LogixXS.Points)
                Balances = MyLookup.GetCustomerPointsBalances_Filtered(CustomerPK, selectedResCols, False, BalRetCode)
                If BalRetCode = RETURN_CODE.OK Then
                    For i = 0 To Balances.GetUpperBound(0)
                        row = dtBalances.NewRow()
                        If (selectedResCols.Contains("ProgramID")) Then row.Item("ProgramID") = Balances(i).ProgramID
                        If (selectedResCols.Contains("Balance")) Then row.Item("Balance") = Balances(i).Balance
                        dtBalances.Rows.Add(row)
                    Next
                    If Balances.GetUpperBound(0) > 0 Then dtBalances.AcceptChanges()
                Else
                    RetCode = StatusCodes.FAILED_BALANCE_LOOKUP
                    RetMsg = "Failed to load customer's points program balances.  Check log for details on this exception."
                End If

                'check for points balances in an external points program
                Balances = MyLookup.GetCustomerExternalPointsBalances_filtered(CustomerPK, selectedResCols, False, BalRetCode)
                If BalRetCode = RETURN_CODE.OK Then
                    For i = 0 To Balances.GetUpperBound(0)
                        row = dtBalances.NewRow()
                        If (selectedResCols.Contains("ProgramID")) Then row.Item("ProgramID") = Balances(i).ProgramID
                        If (selectedResCols.Contains("Balance")) Then row.Item("Balance") = Balances(i).Balance
                        dtBalances.Rows.Add(row)
                    Next
                    If Balances.GetUpperBound(0) > 0 Then dtBalances.AcceptChanges()
                Else
                    RetCode = StatusCodes.FAILED_BALANCE_LOOKUP
                    RetMsg = "Failed to load customer's external points program balances.  Check log for details on this exception."
                End If

            Else
                RetCode = StatusCodes.INVALID_CUSTOMERID
                RetMsg = "Failure: Invalid customer ID."
            End If
        End If

        Return dtBalances
    End Function

    Private Function IsValidGUID(ByVal GUID As String) As Boolean
        Dim IsValid As Boolean = False
        Dim ConnInc As New Copient.ConnectorInc

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            IsValid = ConnInc.IsValidConnectorGUID(MyCommon, 3, GUID)
        Catch ex As Exception
            IsValid = False
        End Try

        Return IsValid
    End Function

    
    
    'function to chk customer card 
    Private Function IsValidCustomerCard(ByVal ExtCardID As String, ByVal CardTypeID As String, ByRef RetCode As StatusCodes, ByRef RetMsg As String) As Boolean
        Dim IsValid As Boolean = False
        Dim validationRespCode As CardValidationResponse
        RetCode = StatusCodes.SUCCESS
        RetMsg = ""
        Try
            If (MyCommon.AllowToProcessCustomerCard(ExtCardID, CardTypeID, validationRespCode) = False) Then
               
                
                If validationRespCode <> CardValidationResponse.SUCCESS Then
                    If validationRespCode = CardValidationResponse.CARDIDNOTNUMERIC OrElse validationRespCode = CardValidationResponse.INVALIDCARDFORMAT Then
                        RetCode = StatusCodes.INVALID_CUSTOMERID
                    ElseIf validationRespCode = CardValidationResponse.CARDTYPENOTFOUND OrElse validationRespCode = CardValidationResponse.INVALIDCARDTYPEFORMAT Then
                        RetCode = StatusCodes.INVALID_CUSTOMERTYPEID
                    ElseIf validationRespCode = CardValidationResponse.ERROR_APPLICATION Then
                        RetCode = StatusCodes.APPLICATION_EXCEPTION
                    End If
                    RetMsg = MyCommon.CardValidationResponseMessage(ExtCardID, CardTypeID, validationRespCode)
                
                End If
                
                
                
            Else
                IsValid = True
            End If
            
        Catch ex As Exception
            IsValid = False
        End Try
        Return IsValid
    End Function
   

    Private Function Send_Transactions(ByVal CustomerPK As Long, ByVal CustomerID As String) As DataTable
        'This function is used by OfferList and returns a list of customer offers
        Dim dt As System.Data.DataTable
        Dim dt2 As System.Data.DataTable
        Dim dtTransactions As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim retString As New StringBuilder
        Dim MyLookup As New Copient.CustomerLookup
        Dim Balances(-1) As Copient.CustomerLookup.PointsBalance
        Dim TransactionDate As New DateTime
        Dim ExtLocationCode As String = ""
        Dim RedemptionAmount As Decimal = 0
        Dim RedemptionCount As Integer = 0
        Dim TerminalNum As String = ""
        Dim TransNum As String = ""
        Dim LogixTransNum As String = ""
        Dim OfferID As Long = 0
        Dim OfferName As String = ""
        Dim DetailRecords As Integer = 0

        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
        If MyCommon.LWHadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixWH()

        'Create a new datatable to hold the results we'll be assembling
        dtTransactions = New DataTable
        dtTransactions.Columns.Add("TransactionDate", System.Type.GetType("System.DateTime"))
        dtTransactions.Columns.Add("ExtLocationCode", System.Type.GetType("System.String"))
        dtTransactions.Columns.Add("RedemptionAmount", System.Type.GetType("System.Decimal"))
        dtTransactions.Columns.Add("RedemptionCount", System.Type.GetType("System.Int32"))
        dtTransactions.Columns.Add("TerminalNum", System.Type.GetType("System.String"))
        dtTransactions.Columns.Add("TransNum", System.Type.GetType("System.String"))
        dtTransactions.Columns.Add("LogixTransNum", System.Type.GetType("System.String"))
        dtTransactions.Columns.Add("OfferID", System.Type.GetType("System.Int64"))
        dtTransactions.Columns.Add("OfferName", System.Type.GetType("System.String"))
        dtTransactions.Columns.Add("DetailRecords", System.Type.GetType("System.Int32"))

        MyCommon.QueryStr = "select top " & MyCommon.Fetch_WebOption(7) & " Max(TransDate) as TransactionDate, ExtLocationCode, " & _
                            "sum(RedemptionAmount) as RedemptionAmount, sum(RedemptionCount) as RedemptionCount, TerminalNum, TransNum, " & _
                            "LogixTransNum, OfferID, count(*) as DetailRecords " & _
                            "from TransRedemptionView with (NoLock) where CustomerPrimaryExtID in ('" & CustomerID & "') " & _
                            "group by CustomerPrimaryExtID, TransNum, TerminalNum, ExtLocationCode, LogixTransNum, OfferID " & _
                            "order by TransactionDate DESC;"
        dt = MyCommon.LWH_Select
        If (dt.Rows.Count > 0) Then
            For Each row In dt.Rows
                TransactionDate = row.Item("TransactionDate")
                ExtLocationCode = row.Item("ExtLocationCode")
                RedemptionAmount = row.Item("RedemptionAmount")
                RedemptionCount = row.Item("RedemptionCount")
                TerminalNum = row.Item("TerminalNum")
                TransNum = row.Item("TransNum")
                LogixTransNum = row.Item("LogixTransNum")
                OfferID = row.Item("OfferID")
                DetailRecords = row.Item("DetailRecords")
                MyCommon.QueryStr = "select IncentiveName as OfferName from CPE_ST_Incentives with (NoLock) where IncentiveID=" & OfferID & _
                                    " union " & _
                                    "select Name as OfferName from Offers with (NoLock) where OfferID=" & OfferID & ";"
                dt2 = MyCommon.LRT_Select
                If dt2.Rows.Count > 0 Then
                    OfferName = MyCommon.NZ(dt2.Rows(0).Item("OfferName"), "")
                End If

                'Put the resulting data into a table
                row = dtTransactions.NewRow()
                row.Item("TransactionDate") = Date.Parse(TransactionDate)
                row.Item("ExtLocationCode") = ExtLocationCode
                row.Item("RedemptionAmount") = RedemptionAmount
                row.Item("RedemptionCount") = RedemptionCount
                row.Item("TerminalNum") = TerminalNum
                row.Item("TransNum") = TransNum
                row.Item("LogixTransNum") = LogixTransNum
                row.Item("OfferID") = OfferID
                row.Item("OfferName") = OfferName
                row.Item("DetailRecords") = DetailRecords
                dtTransactions.Rows.Add(row)

            Next
            If dtTransactions.Rows.Count > 0 Then
                dtTransactions.AcceptChanges()
            End If
        End If

        Return dtTransactions
    End Function

    Private Function Send_List(ByVal CustomerPK As Long, ByVal CustomerID As String) As DataTable
        'This function is used by ShopList and returns a list of shopping items
        Dim dt As System.Data.DataTable
        Dim dtTransactions As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim retString As New StringBuilder
        Dim MyLookup As New Copient.CustomerLookup
        Dim ListPK As String = ""
        Dim Item As String = ""


        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

        'Create a new datatable to hold the results we'll be assembling
        dtTransactions = New DataTable
        dtTransactions.Columns.Add("ListPK", System.Type.GetType("System.String"))
        dtTransactions.Columns.Add("Item", System.Type.GetType("System.String"))

        MyCommon.QueryStr = "select listpk, Item " & _
                            "from ShopList as List with (NoLock) where CustomerPK = " & CustomerPK & _
                            " order by ListPK DESC;"
        dt = MyCommon.LXS_Select




        If (dt.Rows.Count > 0) Then
            For Each row In dt.Rows
                ListPK = row.Item("listpk")
                Item = row.Item("Item")

                'Put the resulting data into a table
                row = dtTransactions.NewRow()
                row.Item("ListPK") = ListPK
                row.Item("Item") = Item


                dtTransactions.Rows.Add(row)

            Next
            If dtTransactions.Rows.Count > 0 Then
                dtTransactions.AcceptChanges()
            End If
        End If

        Return dtTransactions
    End Function

    <WebMethod()> _
    Public Function GetNonTriggeredOffers(ByVal GUID As String) As System.Data.DataSet
        Dim dtOffers As System.Data.DataTable
        Dim dtStatus As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim ResultSet As New System.Data.DataSet("NonTriggeredOffersList")
        Dim bRTConnectionOpened As Boolean = False
        
        InitApp()

        'Initialize the status table, which will report the success or failure of the NonTriggeredOffers operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))
        
        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                MyCommon.Open_LogixRT()
                bRTConnectionOpened = True
            End If
            If Not IsValidGUID(GUID) Then
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
            Else
                dtOffers = GetNonTrigOffersList()
                'Put a success into the Status table
                row = dtStatus.NewRow()
                row.Item("StatusCode") = 0
                             
                If dtOffers.Rows.Count > 0 Then
                    row.Item("Description") = "Success."
                Else
                    row.Item("Description") = "Success.No non-triggered offers are found."
                End If
                dtStatus.Rows.Add(row)
                dtOffers.TableName = "Offers"

                'Populate and return the DataSet
                ResultSet.Tables.Add(dtStatus.Copy())
                ResultSet.Tables.Add(dtOffers.Copy())
            End If
        Catch ex As Exception
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
            row.Item("Description") = "Failure: Application " & ex.ToString
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed AndAlso bRTConnectionOpened Then MyCommon.Close_LogixRT()
        End Try
        Return ResultSet
    End Function
  
    Private Function GetNonTrigOffersList() As DataTable
        Dim dtActOffers As System.Data.DataTable
    Dim dtOffers As System.Data.DataTable
    Dim ActiveOffersTable As System.Data.DataTable
        Dim bRTConnectionOpened As Boolean = False
        Dim bWHConnectionOpened As Boolean = False
        
        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
            MyCommon.Open_LogixRT()
            bRTConnectionOpened = True
        End If
        If MyCommon.LWHadoConn.State = ConnectionState.Closed Then
            MyCommon.Open_LogixWH()
            bWHConnectionOpened = True
        End If
                
        dtOffers = New DataTable
        dtOffers.Columns.Add("OfferID", System.Type.GetType("System.Int64"))
        dtOffers.Columns.Add("ExtOfferID", System.Type.GetType("System.String"))
        dtOffers.Columns.Add("OfferName", System.Type.GetType("System.String"))
        dtOffers.Columns.Add("Engine", System.Type.GetType("System.String"))
        dtOffers.Columns.Add("ExternalSource", System.Type.GetType("System.String"))
        dtOffers.Columns.Add("OfferStartDate", System.Type.GetType("System.DateTime"))
        dtOffers.Columns.Add("OfferEndDate", System.Type.GetType("System.DateTime"))
            
        'Getting all the active and deployed offers
        MyCommon.QueryStr = " SELECT DISTINCT O.OfferID, O.ExtOfferID, O.CMOADeployStatus, O.StatusFlag,O.Name, OID.EngineID, PE.[Description] As Engine," & _
                            "O.ProdStartDate, O.ProdEndDate, O.InboundCRMEngineID, ECT.Name As ExtSource FROM CM_ST_Offers AS O WITH (NOLOCK) " & _
                            "INNER JOIN OfferIDs AS OID WITH (NOLOCK) ON OID.OfferID=O.OfferID " & _
                            "LEFT JOIN PromoEngines PE WITH (NOLOCK) ON PE.EngineID = OID.EngineID " & _
                            "LEFT JOIN ExtCRMInterfaces ECT WITH (NOLOCK) ON ECT.ExtInterfaceID = O.InboundCRMEngineID " & _
                            "WHERE(O.IsTemplate = 0 And O.Deleted = 0 And O.CMOADeployStatus = 1) " & _
                            "AND O.ProdEndDate > '" & Today.AddDays(-1).ToString & "' AND O.ProdStartDate <= GETDATE() " & _
                            "UNION ALL " & _
                            "SELECT DISTINCT I.IncentiveID, I.ClientOfferID, I.CPEOADeployStatus, I.StatusFlag, " & _
                            "I.IncentiveName AS Name, OID.EngineID,PE.[Description] As Engine " & _
                            ",I.StartDate, I.EndDate, I.InboundCRMEngineID, ECT.Name As ExtSource FROM CPE_ST_Incentives AS I WITH (NOLOCK) " & _
                            "INNER JOIN OfferIDs AS OID WITH (NOLOCK) ON OID.OfferID = I.IncentiveID " & _
                            "LEFT JOIN PromoEngines PE WITH (NOLOCK) ON PE.EngineID = OID.EngineID " & _
                            "LEFT JOIN ExtCRMInterfaces ECT WITH (NOLOCK) ON ECT.ExtInterfaceID = I.InboundCRMEngineID " & _
                            "LEFT JOIN cpe_st_OfferTranslations AS OT WITH (NOLOCK) ON OT.OfferID = I.IncentiveID " & _
                            "WHERE(I.IsTemplate = 0 And I.Deleted = 0 AND I.EngineID = 9) AND I.EndDate > '" & Today.AddDays(-1).ToString & "' AND I.StartDate <= GETDATE() "
        
    dtActOffers = MyCommon.LRT_Select
    ActiveOffersTable = dtActOffers
    
    If ActiveOffersTable.Rows.Count > 0 Then
      MyCommon.QueryStr = "dbo.pa_GetNonTriggeredOffers"
      MyCommon.Open_LWHsp()
      MyCommon.LWHsp.Parameters.Add("@Offers", SqlDbType.Structured).Value = ActiveOffersTable
      dtOffers = MyCommon.LWHsp_select()
    End If
       
    If dtOffers.Rows.Count > 0 Then dtOffers.AcceptChanges()

    If MyCommon.LRTadoConn.State <> ConnectionState.Closed AndAlso bRTConnectionOpened Then MyCommon.Close_LogixRT()
    If MyCommon.LWHadoConn.State <> ConnectionState.Closed AndAlso bWHConnectionOpened Then MyCommon.Close_LogixWH()
        
    Return dtOffers
    End Function


  <WebMethod()> _
  Public Function WebOfferList(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As Integer) As System.Data.DataSet
    Return _OfferListWeb(GUID, CustomerID, CustomerTypeID , True)
  End Function
  
  Public Function _OfferListWeb(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As Integer, Optional ByVal WebEngineOnly As Boolean = False) As System.Data.DataSet
    Dim dt As System.Data.DataTable
    Dim dtOffers As System.Data.DataTable
    Dim dtGroups As System.Data.DataTable
    Dim dtStatus As System.Data.DataTable
    Dim row As System.Data.DataRow
    Dim CustomerPK As Long
        Dim ResultSet As New System.Data.DataSet("OfferList")
        Dim LocalMyCommon As New Copient.CommonInc
        WriteDebug("OfferList - CardNumber: " & Copient.MaskHelper.MaskCard(CustomerID, CardTypes.CUSTOMER) & ". CardTypeID: " & CardTypes.CUSTOMER, DebugState.BeginTime)

        'Initialize the status table, which will report the success or failure of the CustomerDetails operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        Try
            If LocalMyCommon.LRTadoConn.State = ConnectionState.Closed Then LocalMyCommon.Open_LogixRT()
            If LocalMyCommon.LXSadoConn.State = ConnectionState.Closed Then LocalMyCommon.Open_LogixXS()

            If Not IsValidGUID(GUID) Then
                'Wrong GUID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_GUID
                row.Item("Description") = "Failure: Invalid GUID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
            ElseIf (CustomerID.Length < 1) Then
                'Bad customer ID
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERID
                row.Item("Description") = "Failure: Invalid customer ID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
            Else
                'Pad the customer ID
                CustomerID = LocalMyCommon.Pad_ExtCardID(CustomerID, CustomerTypeID)
                'Find the Customer PK
                LocalMyCommon.QueryStr = "select CID.CustomerPK from CardIDs as CID with (NoLock) " & _
                                    "INNER JOIN CardTypes AS CT with (NoLock) ON CID.CardTypeID = CT.CardTypeID " & _
                                    "WHERE CT.CustTypeID=" & CustomerTypeID & " AND CID.ExtCardID='" & MyCryptLib.SQL_StringEncrypt(CustomerID, True) & "';"
                dt = LocalMyCommon.LXS_Select
                If dt.Rows.Count = 0 Then
                    'Customer not found
                    If CustomerTypeID = 0 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_CUSTOMER
                        row.Item("Description") = "Failure: Customer " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    ElseIf CustomerTypeID = 1 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_HOUSEHOLD
                        row.Item("Description") = "Failure: Household " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    ElseIf CustomerTypeID = 2 Then
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.NOTFOUND_CAM
                        row.Item("Description") = "Failure: CAM " & CustomerID & " not found."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    Else
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERTYPEID
                        row.Item("Description") = "Failure: Invalid customer type."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    End If
                Else
                    CustomerPK = LocalMyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)

                    'Put a success into the Status table
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = 0
                    row.Item("Description") = "Success."
                    dtStatus.Rows.Add(row)
          
                    'Get the lists of offers
                    'Get the lists of all engine offers other than website
                    If Not WebEngineOnly Then
                        WriteDebug("Entering Send_XMLCurrentOffers", DebugState.CurrentTime)
                        dtOffers = Send_XMLCurrentOffers(CustomerPK)
                        WriteDebug("Exiting Send_XMLCurrentOffers", DebugState.CurrentTime)
                    End If
 
          
                    'Get the lists of website offers
                    If (bEnableFilterForResponse) Then
                        WriteDebug("Entering Send_XMLGroupOffers_FilteredResult", DebugState.CurrentTime)
                        dtGroups = Send_XMLGroupOffers_FilteredResult(CustomerPK)
                        WriteDebug("Exiting Send_XMLGroupOffers_FilteredResult", DebugState.CurrentTime)
                    Else
                        WriteDebug("Entering Send_XMLGroupOffers", DebugState.CurrentTime)
                        dtGroups = Send_XMLGroupOffers(CustomerPK)
                        WriteDebug("Exiting Send_XMLGroupOffers", DebugState.CurrentTime)
                    End If

                    If Not WebEngineOnly Then
                        dtOffers.TableName = "Offers"
                    End If
          
                    dtGroups.TableName = "Groups"

                    'Populate and return the DataSet
                    ResultSet.Tables.Add(dtStatus.Copy())
          
                    If Not WebEngineOnly Then
                        ResultSet.Tables.Add(dtOffers.Copy())
                    End If
          
                    ResultSet.Tables.Add(dtGroups.Copy())

                End If
            End If
        Catch ex As Exception
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
            row.Item("Description") = "Failure: Application " & ex.ToString
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
        Finally
            If LocalMyCommon.LRTadoConn.State <> ConnectionState.Closed Then LocalMyCommon.Close_LogixRT()
            If LocalMyCommon.LXSadoConn.State <> ConnectionState.Closed Then LocalMyCommon.Close_LogixXS()
        End Try
    WriteDebug("OfferList - " & System.Net.Dns.GetHostName(), DebugState.EndTime)
    WriteLogLines()

    Return ResultSet
  End Function

  Public Class PickupStore
    Public CmStoreNumber As String = ""
    Public CmStoreName As String = ""
    Public UeLocationCode As String = ""
  End Class

  Public Class PickupStoreList
    Public Status As CustWebStatus
    Public PickupStoreList() As PickupStore
  End Class

  <WebMethod()> _
  Public Function GetPickupStoreList(ByVal GUID As String) As PickupStoreList
    Dim oPickupStoreList As New PickupStoreList
    Dim oStatus As New CustWebStatus
    Dim dt As DataTable
    Dim dr As DataRow
    Dim i As Integer
    Dim bRTConnectionOpened As Boolean = False
    
    InitApp()
    oPickupStoreList.Status = oStatus
    If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
      MyCommon.Open_LogixRT()
      bRTConnectionOpened = True
    End If

    Try
      MyCommon.QueryStr = "select CM.ExtLocationCode as CmStoreNumber, CM.LocationName as CmStoreName, UE.ExtLocationCode as UeLocationCode " & _
                          "from Locations UE with (NoLock) " & _
                          "inner join Locations CM with (NoLock) on CM.LocationId = UE.BrickAndMortarLocationID " & _
                          "where UE.Deleted = 0;"
      dt = MyCommon.LRT_Select
    
      If dt.Rows.Count > 0 Then
        ReDim oPickupStoreList.PickupStoreList(dt.Rows.Count - 1)
        Dim oPickupStore As PickupStore
        i = 0
        For Each dr In dt.Rows
          oPickupStore = New PickupStore
          oPickupStore.CmStoreNumber = dr.Item("CmStoreNumber")
          oPickupStore.CmStoreName = dr.Item("CmStoreName")
          oPickupStore.UeLocationCode = dr.Item("UeLocationCode")
          oPickupStoreList.PickupStoreList(i) = oPickupStore
          i += 1
        Next
      Else
        oPickupStoreList.Status.StatusCode = StatusCodes.NOTFOUND_RECORDS
        oPickupStoreList.Status.Description = "No Pick Up stores are defined for UE!"
      End If
    Catch ex As Exception
      oPickupStoreList.Status.StatusCode = StatusCodes.APPLICATION_EXCEPTION
      oPickupStoreList.Status.Description = ex.Message
    End Try

    If bRTConnectionOpened Then
      If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then
        MyCommon.Close_LogixRT()
      End If
    End If
    Return oPickupStoreList
  End Function
  
    Private Sub InitApp()
        MyCommon.AppName = "CustWeb.asmx"
    End Sub

    Private Sub WriteDebug(ByVal sText As String, ByVal mode As DebugState)
        If bDebugLogOn Then
            Dim TotalSeconds As Double
            Dim sIndent As String
	  
            Select Case mode
                Case DebugState.BeginTime
                    ' first call
                    DebugStartTimes.Add(Now)
		  
                    If DebugStartTimes.Count = 1 Then
                        Dim sReturnMsg As String = ""
                        sReturnMsg = GetComputerID("WebServer", False)
                        WriteLog(scDashes & sReturnMsg, MessageType.Debug)
                    Else
                        sIndent = New String(" "c, 2 * (DebugStartTimes.Count - 1))
                        sText = sIndent & sText
                    End If
                    sText = sText & " - Begin"
		  
                Case DebugState.EndTime
                    ' last call
                    If DebugStartTimes.Count > 0 Then
                        TotalSeconds = Now.Subtract(DebugStartTimes(DebugStartTimes.Count - 1)).TotalSeconds
                        sIndent = New String(" "c, 2 * (DebugStartTimes.Count - 1))
                        sText = sIndent & sText & " - End elapsed time: " & Format(TotalSeconds, "00.000") & "(sec)"
                        DebugStartTimes.RemoveAt(DebugStartTimes.Count - 1)
                    End If
		  
                Case Else
                    ' interim call
                    If DebugStartTimes.Count > 0 Then
                        TotalSeconds = Now.Subtract(DebugStartTimes(DebugStartTimes.Count - 1)).TotalSeconds
                        sIndent = New String(" "c, 2 * (DebugStartTimes.Count - 1))
                        sText = sIndent & sText & " - Current elapsed time: " & Format(TotalSeconds, "0.000") & "(sec)"
                    End If
            End Select
            WriteLog(sText, MessageType.Debug)
        End If
    End Sub

    Private Sub WriteLog(ByVal sText As String, ByVal eType As MessageType)
        Dim sLogText As String = ""

        If eType = MessageType.Debug Then
            sLogText = "[" & Date.Now.ToString("MM/dd/yyyy HH:mm:ss.ffff") & " (Type=" & eType.ToString & ")] " & sText
        Else
            If eType <> MessageType.Info Then
                sText = sText.Replace(ControlChars.CrLf, " ")
            End If
            If sInputForLog.Length > 0 Then
                sLogText = "[" & Date.Now.ToString("MM/dd/yyyy HH:mm:ss.ffff") & " (Type=" & eType.ToString & ")] " & sInputForLog & ControlChars.CrLf
                sInputForLog = ""
            End If
            sLogText = sLogText & "[" & Date.Now.ToString("MM/dd/yyyy HH:mm:ss.ffff") & " (Type=" & eType.ToString & ")] " & sText
        End If
        sLogLines = sLogLines & sLogText & vbCrLf
    End Sub

    Private Sub WriteLogLines()
        Dim sFileName As String
        Dim sLogText As String = ""

        If sLogLines.Length > 2 Then
            Try
                sLogLines = sLogLines.Remove(sLogLines.Length - 2)
                sFileName = sLogFileName & "." & MyCommon.Leading_Zero_Fill(Year(Today), 4) & MyCommon.Leading_Zero_Fill(Month(Today), 2) & MyCommon.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
                MyCommon.Write_Log(sFileName, sLogLines)
            Catch ex As Exception
                Try
                    MyCommon.Error_Processor(, "WriteLog", sAppName, sInstallationName)
                Catch
                End Try
            End Try
        End If
    End Sub
  
    Private Function GetComputerID(ByVal sServerText As String, ByVal bAdapter As Boolean) As String
        Dim sIPAddress As String
        Dim sReturnMessage As String
        Dim adapters As NetworkInterface() = NetworkInterface.GetAllNetworkInterfaces()
        Dim adapter As NetworkInterface
 
        sReturnMessage = " " & sServerText & ": " & System.Net.Dns.GetHostName()
	
        sIPAddress = HttpContext.Current.Request.ServerVariables("HTTP_VIA")
        If sIPAddress = "" Then
            sIPAddress = HttpContext.Current.Request.ServerVariables("REMOTE_ADDR")
        Else
            sIPAddress = HttpContext.Current.Request.ServerVariables("HTTP_X_FORWARDED_FOR")
        End If
        sReturnMessage = sReturnMessage & "  HttpContext: " & sIPAddress & "  InterNetwork:"
			
        For Each IPA As Net.IPAddress In System.Net.Dns.GetHostAddresses(Net.Dns.GetHostName())
            If IPA.AddressFamily.ToString() = "InterNetwork" Then
                sIPAddress = IPA.ToString()
                sReturnMessage = sReturnMessage & " - " & sIPAddress
            End If
        Next IPA
    
        If bAdapter = True Then
            For Each adapter In adapters
                sReturnMessage = sReturnMessage & " - " & adapter.Description & " " & adapter.Name
            Next adapter
        End If

        Return sReturnMessage
    End Function
    
    '******************************************************************************
    'GetTransHistoryByStoreNumber returns Transaction Information for a specified Store
    'ExtLocationCode is required
    'StartDateTime, EndDateTime, IncludeItemDetails, and RedemptionFilter are optional
    'Expected Format for optional parameters:
    ' StartDateTime and EndDateTime - mm/dd/yyyy hh:mm
    ' IncludeItemDetails - Y or N
    '  RedemptionFilters - 1 (all transactions), 2 (only transaction with redemptions), 
    '                      3 (only transactions without redemptions)
    '******************************************************************************
    <WebMethod()> _
    Public Function GetTransHistoryByStoreNumber(ByVal GUID As String, ByVal ExtLocationCode As String, ByVal StartDateTime As String, ByVal EndDateTime as String, ByVal sLast4CardID As String, ByVal IncludeItemDetails As String, ByVal RedemptionFilter As String) As System.Data.DataSet
        Dim bIncludeItemDetails As Boolean = False
        Dim iRedemptionFilter As Integer = 1
        
        InitApp()
        If (String.Compare(IncludeItemDetails, "Y", True) = 0) OrElse (String.Compare(IncludeItemDetails, "Yes", True) = 0) Then 
          bIncludeItemDetails = True
        End If
        Try
          iRedemptionFilter = CInt(RedemptionFilter)
          If (iRedemptionFilter < 1) OrElse (iRedemptionFilter > 3) Then 
            iRedemptionFilter = 1
          End If
        Catch ex As Exception
            iRedemptionFilter = 1
        End Try
        
        Return _GetTransHistoryByStoreNumber(GUID, ExtLocationCode, StartDateTime, EndDateTime, sLast4CardID, bIncludeItemDetails, iRedemptionFilter)
    End Function

    Public Function _GetTransHistoryByStoreNumber(ByVal GUID As String, ByVal ExtLocationCode As String, Optional ByVal StartDateTime As String = "", Optional  ByVal EndDateTime as String = "", Optional  ByVal sLast4CardID As String = "", Optional  ByVal bIncludeItemDetails As Boolean = False, Optional  ByVal iRedemptionFilter As Integer = 1) As DataSet
        Dim tempstr As String = ""
        Dim DateFilter As String = ""
        Dim ResultSet As New System.Data.DataSet("TransactionHistoryByStore")
        Dim bOpenedWHConnection As Boolean = False
        Dim bOpenedTRXConnection As Boolean = False
        Dim bOpenedRTConnection As Boolean = False
        Dim dtStatus As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim dt As DataTable
        Dim dtTrans As DataTable
        Dim drTrans As System.Data.DataRow
        Dim dtTransItemByStoreID As DataTable
        Dim drItems As DataRow
        Dim LocationName As String
        Dim StartDT As DateTime
        Dim EndDT As DateTime
        Dim TransItemInfo As String = ""
        bIncludeItemDetails = bIncludeItemDetails AndAlso MyCommon.Fetch_CPE_SystemOption(169)
        
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
            _GetTransHistoryByStoreNumber = ResultSet
            Exit Function
        End If
        If ExtLocationCode = "" Then
            'Store ID Value is empty
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_EXTLOCATIONCODE
            row.Item("Description") = "Failure: Invalid Location Code"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _GetTransHistoryByStoreNumber = ResultSet
            Exit Function
        End If
        If StartDateTime <> "" Then
            'StartDateTime needs to be validated
            Try
              StartDT = DateTime.Parse(StartDateTime)
            Catch ex as exception
              row = dtStatus.NewRow()
              row.Item("StatusCode") = StatusCodes.INVALID_STARTDATE
              row.Item("Description") = "Failure: Invalid Start Date - " & ex.ToString
              dtStatus.Rows.Add(row)
              dtStatus.AcceptChanges()
              ResultSet.Tables.Add(dtStatus.Copy())
              _GetTransHistoryByStoreNumber = ResultSet
              Exit Function
            End Try
        End If
        If EndDateTime <> "" Then
            'EndDateTime needs to be validated
            Try
              EndDT = DateTime.Parse(EndDateTime)
            Catch ex as exception
              row = dtStatus.NewRow()
              row.Item("StatusCode") = StatusCodes.INVALID_ENDDATE
              row.Item("Description") = "Failure: Invalid End Date - " & ex.ToString
              dtStatus.Rows.Add(row)
              dtStatus.AcceptChanges()
              ResultSet.Tables.Add(dtStatus.Copy())
              _GetTransHistoryByStoreNumber = ResultSet
              Exit Function
            End Try
        End If
        If sLast4CardID <> "" Then
            'sLast4CardID needs to be validated
            If sLast4CardID.Length <> 4 Then
              row = dtStatus.NewRow()
              row.Item("StatusCode") = StatusCodes.INVALID_LAST4_CARDID
              row.Item("Description") = "Failure: Invalid Last 4 of Card ID "
              dtStatus.Rows.Add(row)
              dtStatus.AcceptChanges()
              ResultSet.Tables.Add(dtStatus.Copy())
              _GetTransHistoryByStoreNumber = ResultSet
              Exit Function
            End If
        End If
        If GUID.Contains("'") = True Or GUID.Contains(Chr(34)) = True Then
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_GUID
            row.Item("Description") = "Failure: Invalid GUID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _GetTransHistoryByStoreNumber = ResultSet
            Exit Function
        End If
        If ExtLocationCode.Contains("'") = True Or ExtLocationCode.Contains(Chr(34)) = True Then
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_EXTLOCATIONCODE
            row.Item("Description") = "Failure: Invalid ExtLocationCode"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _GetTransHistoryByStoreNumber = ResultSet
            Exit Function
        End If
        
        Dim dtTransByStoreID As System.Data.DataTable
        Dim drTransByStoreID As DataRow
        dtTransByStoreID = New DataTable
        dtTransByStoreID.TableName = "Transaction"
        dtTransByStoreID.Columns.Add("LogixTransactionNumber", System.Type.GetType("System.String"))
        dtTransByStoreID.Columns.Add("TransactionItemInfo", System.Type.GetType("System.String"))
        dtTransByStoreID.Columns.Add("Date", System.Type.GetType("System.String"))
        dtTransByStoreID.Columns.Add("ExtLocationCode", System.Type.GetType("System.String"))
        dtTransByStoreID.Columns.Add("LocationName", System.Type.GetType("System.String"))
        dtTransByStoreID.Columns.Add("ExtCardID", System.Type.GetType("System.String"))
        dtTransByStoreID.Columns.Add("CardType", System.Type.GetType("System.String"))
        dtTransByStoreID.Columns.Add("PresentedCardID", System.Type.GetType("System.String"))
        dtTransByStoreID.Columns.Add("HouseHoldID", System.Type.GetType("System.String"))
        dtTransByStoreID.Columns.Add("TerminalNumber", System.Type.GetType("System.String"))
        dtTransByStoreID.Columns.Add("RedemptionCount", System.Type.GetType("System.String"))
        dtTransByStoreID.Columns.Add("RedemptionAmount", System.Type.GetType("System.String"))
        
        Try
        
          If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
            bOpenedRTConnection = True
            MyCommon.Open_LogixRT()
          End If
          If MyCommon.LWHadoConn.State = ConnectionState.Closed Then
            bOpenedWHConnection = True
            MyCommon.Open_LogixWH()
          End If
          If bIncludeItemDetails AndAlso MyCommon.LTRXadoConn.State = ConnectionState.Closed Then
              bOpenedTRXConnection = True
              MyCommon.Open_LogixTRX()
          End If

          If Not IsValidGUID(GUID) Then
            'Wrong GUID
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_GUID
            row.Item("Description") = "Failure: Invalid GUID."
          Else
            'Fetch transactions based on StoreID from TransHist table
            
            If sLast4CardID <> "" Then
              tempstr = " And (TH.PresentedCustomerID like '%" & sLast4CardID & "') "
            End If
           
            If (StartDateTime <> "") And (EndDateTime <> "")  Then
              DateFilter = " having max(TH.TransDate) between '" & StartDateTime & "'" & _
                 " and '" & EndDateTime & "' "  
            Else If (StartDateTime <> "") Then
              DateFilter = " having max(TH.TransDate) > '" & StartDateTime & "' "               
            Else If (EndDateTime <> "")  Then
              DateFilter = " having max(TH.TransDate) < '" & EndDateTime & "' "  
            End If
            
            MyCommon.QueryStr = "select LocationName from Locations where ExtLocationCode = '" & ExtLocationCode & "';"
            dt = MyCommon.LRT_Select
            If dt.Rows.Count > 0 Then
              LocationName = dt.Rows(0).Item("LocationName")
              
              MyCommon.QueryStr = "dbo.pc_Transaction_Select"
              MyCommon.Open_LWHsp()
              MyCommon.LWHsp.Parameters.Add("@ExtLocationCode", SqlDbType.NVarChar, 20).Value = ExtLocationCode
              MyCommon.LWHsp.Parameters.Add("@Max", SqlDbType.Int).Value = 1000
              MyCommon.LWHsp.Parameters.Add("@SearchFilter", SqlDbType.NVarChar, 1000).Value = tempstr
              MyCommon.LWHsp.Parameters.Add("@DateFilter", SqlDbType.NVarChar, 1000).Value = DateFilter
              MyCommon.LWHsp.Parameters.Add("@RedemptionFilter", SqlDbType.Int).Value = iRedemptionFilter              
              dtTrans = MyCommon.LWHsp_select()
              MyCommon.Close_LXSsp()
              
              If dtTrans.Rows.Count > 0 Then
                'Transaction Information found
                If dtTrans.Rows.Count > 1000
                  row = dtStatus.NewRow()
                  row.Item("StatusCode") = StatusCodes.FAILED_TOOMANYRECORDS
                  row.Item("Description") = "Failure: Too Many Records Returned" 
                  _GetTransHistoryByStoreNumber = ResultSet
                  Exit Function                  
                Else
                  row = dtStatus.NewRow()
                  row.Item("StatusCode") = StatusCodes.SUCCESS
                  row.Item("Description") = "Success"                  
                End If
                
                For Each drTrans In dtTrans.Rows
                  drTransByStoreID = dtTransByStoreID.NewRow()
                  drTransByStoreID.Item("LogixTransactionNumber") = drTrans.Item("LogixTransNum")
                  
                  'Check if Item details should be returned
                  If bIncludeItemDetails Then
                    MyCommon.QueryStr = "Select ItemID, Quantity, Price, Description from TransactionItem with (NoLock) where LogixTransNum='" & drTrans.Item("LogixTransNum") & "';" 
                    dtTransItemByStoreID = MyCommon.LTRX_Select
                    If dtTransItemByStoreID.Rows.Count > 0 Then
                      '' Item information found
                      For Each drItems In dtTransItemByStoreID.Rows
                        TransItemInfo = TransItemInfo & drItems.Item("ItemID") & "," & drItems.Item("Description") & "," & drItems.Item("Quantity") & "," & drItems.Item("Price") & ","  & vbcrlf
                      Next drItems
                      drTransByStoreID.Item("TransactionItemInfo") = "<![CDATA[" & TransItemInfo & "]]>"
                    Else
                      'No records found
                      drTransByStoreID.Item("TransactionItemInfo") = "<![CDATA[]]>"
                    End If
                  End If
                  
                  drTransByStoreID.Item("Date") = drTrans.Item("TransactionDate")
                  drTransByStoreID.Item("ExtLocationCode") = ExtLocationCode
                  drTransByStoreID.Item("LocationName") = LocationName
                            drTransByStoreID.Item("ExtCardID") = drTrans.Item("CustomerPrimaryExtID").ToString()
                  drTransByStoreID.Item("CardType") = drTrans.Item("PresentedCardTypeID")
                  drTransByStoreID.Item("PresentedCardID") = drTrans.Item("PresentedCustomerID")
                  drTransByStoreID.Item("HouseHoldID") = drTrans.Item("HHID")
                  drTransByStoreID.Item("TerminalNumber") = drTrans.Item("TerminalNum")
                  drTransByStoreID.Item("RedemptionCount") = drTrans.Item("RedemptionCount")
                  drTransByStoreID.Item("RedemptionAmount") = drTrans.Item("RedemptionAmount")
                  dtTransByStoreID.Rows.Add(drTransByStoreID)
                Next 

                dtTransByStoreID.AcceptChanges()
                ResultSet.Tables.Add(dtTransByStoreID.Copy())
                                      
                  
              Else
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.NOTFOUND_TRANSACTIONS
                row.Item("Description") = "Failure: No transactions Found " 
              End If 'Transactions
            Else
              row = dtStatus.NewRow()
              row.Item("StatusCode") = StatusCodes.INVALID_EXTLOCATIONCODE
              row.Item("Description") = "Failure: Invalid ExtLocationCode"
            End If 'ExtLocationCode
          End If 'Valid GUI
        
        Catch ex As Exception
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
            row.Item("Description") = "Failure: Application " & ex.ToString
        Finally
            If bOpenedWHConnection = True And MyCommon.LWHadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixWH()
            If bOpenedTRXConnection = True And MyCommon.LTRXadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixTRX()
            If bOpenedRTConnection = True And MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
        End Try
        Return ResultSet
    End Function
    
    '******************************************************************************
    'GetTransHistoryByCard returns Transaction Information for a specified Card ID
    'ExtCardID and CardTypeID are required
    'IncludeItemDetails is optional, format - Y or N
    '******************************************************************************
    <WebMethod()> _
    Public Function GetTransHistoryByCard(ByVal GUID As String, ByVal ExtCardID As String, ByVal CardTypeID As String, ByVal IncludeItemDetails As String) As System.Data.DataSet 
        Dim bIncludeItemDetails As Boolean = False
        
        InitApp()
        If (String.Compare(IncludeItemDetails, "Y", True) = 0) OrElse (String.Compare(IncludeItemDetails, "Yes", True) = 0) Then 
          bIncludeItemDetails = True
        End If

        Return _GetTransHistoryByCard(GUID, ExtCardID, CardTypeID, bIncludeItemDetails)
    End Function

    Public Function _GetTransHistoryByCard(ByVal GUID As String, ByVal ExtCardID As String, ByVal CardTypeID As String, Optional  ByVal bIncludeItemDetails As Boolean = False) As DataSet
        Dim ResultSet As New System.Data.DataSet("TransactionHistoryByCard")
        Dim bOpenedWHConnection As Boolean = False
        Dim bOpenedTRXConnection As Boolean = False
        Dim bOpenedRTConnection As Boolean = False
        Dim bOpenedXSConnection As Boolean = False
        Dim dtStatus As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim dt As DataTable
        Dim dtTrans As DataTable
        Dim drTrans As System.Data.DataRow
        Dim dtTransItemByStoreID As DataTable
        Dim drItems As DataRow
        Dim ExtLocationCode As String
        Dim LocationName As String
        Dim TransItemInfo As String = ""
        bIncludeItemDetails = bIncludeItemDetails AndAlso MyCommon.Fetch_CPE_SystemOption(169)
        
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
            _GetTransHistoryByCard = ResultSet
            Exit Function
        End If
        If ExtCardID = "" Then
            'ExtCardID Value is empty
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERID
            row.Item("Description") = "Failure: Invalid Customer ID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _GetTransHistoryByCard = ResultSet
            Exit Function
        End If
        If CardTypeID = "" Then
            'CardTypeID Value is empty
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERTYPEID
            row.Item("Description") = "Failure: Invalid Customer Type ID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _GetTransHistoryByCard = ResultSet
            Exit Function
        End If
        If GUID.Contains("'") = True Or GUID.Contains(Chr(34)) = True Then
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_GUID
            row.Item("Description") = "Failure: Invalid GUID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _GetTransHistoryByCard = ResultSet
            Exit Function
        End If
        If ExtCardID.Contains("'") = True Or ExtCardID.Contains(Chr(34)) = True Then
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERID
            row.Item("Description") = "Failure: Invalid ExtCardID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _GetTransHistoryByCard = ResultSet
            Exit Function
        End If
        If CardTypeID.Contains("'") = True Or CardTypeID.Contains(Chr(34)) = True Then
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERTYPEID
            row.Item("Description") = "Failure: Invalid CardTypeID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _GetTransHistoryByCard = ResultSet
            Exit Function
        End If
        
        Dim dtTransByStoreID As System.Data.DataTable
        Dim drTransByStoreID As DataRow
        dtTransByStoreID = New DataTable
        dtTransByStoreID.TableName = "Transaction"
        dtTransByStoreID.Columns.Add("LogixTransactionNumber", System.Type.GetType("System.String"))
        dtTransByStoreID.Columns.Add("TransactionItemInfo", System.Type.GetType("System.String"))
        dtTransByStoreID.Columns.Add("Date", System.Type.GetType("System.String"))
        dtTransByStoreID.Columns.Add("ExtLocationCode", System.Type.GetType("System.String"))
        dtTransByStoreID.Columns.Add("LocationName", System.Type.GetType("System.String"))
        dtTransByStoreID.Columns.Add("ExtCardID", System.Type.GetType("System.String"))
        dtTransByStoreID.Columns.Add("CardType", System.Type.GetType("System.String"))
        dtTransByStoreID.Columns.Add("PresentedCardID", System.Type.GetType("System.String"))
        dtTransByStoreID.Columns.Add("HouseHoldID", System.Type.GetType("System.String"))
        dtTransByStoreID.Columns.Add("TerminalNumber", System.Type.GetType("System.String"))
        dtTransByStoreID.Columns.Add("RedemptionCount", System.Type.GetType("System.String"))
        dtTransByStoreID.Columns.Add("RedemptionAmount", System.Type.GetType("System.String"))
        
        Try
        
          If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
            bOpenedRTConnection = True
            MyCommon.Open_LogixRT()
          End If
          If MyCommon.LWHadoConn.State = ConnectionState.Closed Then
            bOpenedWHConnection = True
            MyCommon.Open_LogixWH()
          End If
          If MyCommon.LXSadoConn.State = ConnectionState.Closed Then
            bOpenedXSConnection = True
            MyCommon.Open_LogixXS()
          End If
          If bIncludeItemDetails AndAlso MyCommon.LTRXadoConn.State = ConnectionState.Closed Then
              bOpenedTRXConnection = True
              MyCommon.Open_LogixTRX()
          End If

          If Not IsValidGUID(GUID) Then
            'Wrong GUID
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_GUID
            row.Item("Description") = "Failure: Invalid GUID."
          Else
            'Fetch transactions based on StoreID from TransHist table
                        
                MyCommon.QueryStr = "select TH.CustomerPrimaryExtID, Max(TH.TransDate) as TransactionDate, TH.ExtLocationCode, sum(TR.RedemptionAmount) as RedemptionAmount, sum(TR.RedemptionCount) as RedemptionCount, " & _
                          "TH.TerminalNum, TH.LogixTransNum, POSTransNum as TransNum, TH.CustomerTypeID, TH.PresentedCustomerID, TH.PresentedCardTypeID, TH.HHID, TH.Replayed, isnull(TransTotal,0) as TransTotal  " & _
                          "from TransHist as TH with (NoLock) " & _
                          "Left Outer Join TransRedemptionView as TR with (NoLock) " & _
                          "on TH.LogixTransNum=TR.LogixTransNum " & _
                          "where TH.CustomerPrimaryExtID='" & ExtCardID & "' and TH.CustomerTypeID='" & CardTypeID & "' group by TH.CustomerPrimaryExtID, TH.HHID, TH.CustomerTypeID, TH.PresentedCustomerID, TH.PresentedCardTypeID, TH.LogixTransNum, POSTransNum, TH.TerminalNum, TH.ExtLocationCode, TH.Replayed, TransTotal " & _
                          "order by TH.LogixTransNum"
            dtTrans = MyCommon.LWH_Select
            If dtTrans.Rows.Count > 0 Then
              'Transaction Information found
              row = dtStatus.NewRow()
              row.Item("StatusCode") = StatusCodes.SUCCESS
              row.Item("Description") = "Success"                  
                
              For Each drTrans In dtTrans.Rows
                drTransByStoreID = dtTransByStoreID.NewRow()
                drTransByStoreID.Item("LogixTransactionNumber") = drTrans.Item("LogixTransNum")
                
                'Check if Item details should be returned
                If bIncludeItemDetails Then
                  MyCommon.QueryStr = "Select ItemID, Quantity, Price, Description from TransactionItem with (NoLock) where LogixTransNum='" & drTrans.Item("LogixTransNum") & "';" 
                  dtTransItemByStoreID = MyCommon.LTRX_Select
                  If dtTransItemByStoreID.Rows.Count > 0 Then
                    '' Item information found
                    For Each drItems In dtTransItemByStoreID.Rows
                      TransItemInfo = TransItemInfo & drItems.Item("ItemID") & "," & drItems.Item("Description") & "," & drItems.Item("Quantity") & "," & drItems.Item("Price") & ","  & vbcrlf
                    Next drItems
                    drTransByStoreID.Item("TransactionItemInfo") = "<![CDATA[" & TransItemInfo & "]]>"
                  Else
                    'No records found
                    drTransByStoreID.Item("TransactionItemInfo") = "<![CDATA[]]>"
                  End If
                End If
                
                ExtLocationCode = drTrans.Item("ExtLocationCode")
                MyCommon.QueryStr = "select LocationName from Locations where ExtLocationCode = '" & ExtLocationCode & "';"
                dt = MyCommon.LRT_Select
                If dt.Rows.Count > 0 Then
                  LocationName = dt.Rows(0).Item("LocationName")
                End If
                
                drTransByStoreID.Item("Date") = drTrans.Item("TransactionDate")
                drTransByStoreID.Item("ExtLocationCode") = ExtLocationCode
                drTransByStoreID.Item("LocationName") = LocationName
                drTransByStoreID.Item("ExtCardID") = ExtCardID
                drTransByStoreID.Item("CardType") = CardTypeID
                drTransByStoreID.Item("PresentedCardID") = drTrans.Item("PresentedCustomerID")
                drTransByStoreID.Item("HouseHoldID") = drTrans.Item("HHID")
                drTransByStoreID.Item("TerminalNumber") = drTrans.Item("TerminalNum")
                drTransByStoreID.Item("RedemptionCount") = drTrans.Item("RedemptionCount")
                drTransByStoreID.Item("RedemptionAmount") = drTrans.Item("RedemptionAmount")
                dtTransByStoreID.Rows.Add(drTransByStoreID)
              Next 

              dtTransByStoreID.AcceptChanges()
              ResultSet.Tables.Add(dtTransByStoreID.Copy())
                                    
                
            Else
                    MyCommon.QueryStr = "select CardPK from CardIDs where ExtCardID = '" & MyCryptLib.SQL_StringEncrypt(ExtCardID) & "' and CardTypeID = '" & CardTypeID & "';"
              dt = MyCommon.LXS_Select
              If dt.Rows.Count > 0 Then
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.NOTFOUND_CUSTOMER
                row.Item("Description") = "Failure: No transactions Found " 
              Else
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.NOTFOUND_TRANSACTIONS
                row.Item("Description") = "Failure: No transactions Found " 
              End If
            End If 'Transactions
          
          End If 'Valid GUI
        
        Catch ex As Exception
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
            row.Item("Description") = "Failure: Application " & ex.ToString
        Finally
            If bOpenedWHConnection = True And MyCommon.LWHadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixWH()
            If bOpenedTRXConnection = True And MyCommon.LTRXadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixTRX()
            If bOpenedRTConnection = True And MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If bOpenedXSConnection = True And MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
            
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
        End Try
        Return ResultSet
    End Function
    
    '************************************************************************************
    'GetDetailsByTransactionNumber returns Item Details for a specified Transaction
    'TransNum is required
    '************************************************************************************
    <WebMethod()> _
    Public Function GetDetailsByTransactionNumber(ByVal GUID As String, ByVal TransNum As String) As System.Data.DataSet 

        InitApp()

        Return _GetDetailsByTransactionNumber(GUID, TransNum)
    End Function

    Public Function _GetDetailsByTransactionNumber(ByVal GUID As String, ByVal TransNum As String) As DataSet
        Dim ResultSet As New System.Data.DataSet("DetailsByTransactionNumber")
        Dim bOpenedWHConnection As Boolean = False
        Dim bOpenedTRXConnection As Boolean = False
        Dim bOpenedRTConnection As Boolean = False
        Dim bOpenedXSConnection As Boolean = False
        Dim dtStatus As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim dt As DataTable
        Dim dtTrans As DataTable
        Dim drTrans As System.Data.DataRow
        Dim dtTransItemByStoreID As DataTable
        Dim drItems As DataRow
        Dim ExtLocationCode As String
        Dim LocationName As String
        Dim TransItemInfo As String = ""
        
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
            _GetDetailsByTransactionNumber = ResultSet
            Exit Function
        End If
        If TransNum = "" Then
            'TransNum Value is empty
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_TRANSNUM
            row.Item("Description") = "Failure: Invalid TransactionNumber"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _GetDetailsByTransactionNumber = ResultSet
            Exit Function
        End If
        If GUID.Contains("'") = True Or GUID.Contains(Chr(34)) = True Then
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_GUID
            row.Item("Description") = "Failure: Invalid GUID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _GetDetailsByTransactionNumber = ResultSet
            Exit Function
        End If
        
        Dim dtTransByStoreID As System.Data.DataTable
        Dim drTransByStoreID As DataRow
        dtTransByStoreID = New DataTable
        dtTransByStoreID.TableName = "Transaction"
        'dtTransByStoreID.Columns.Add("LogixTransactionNumber", System.Type.GetType("System.String"))
        dtTransByStoreID.Columns.Add("TransactionItemInfo", System.Type.GetType("System.String"))
        dtTransByStoreID.Columns.Add("Date", System.Type.GetType("System.String"))
        dtTransByStoreID.Columns.Add("ExtLocationCode", System.Type.GetType("System.String"))
        dtTransByStoreID.Columns.Add("LocationName", System.Type.GetType("System.String"))
        dtTransByStoreID.Columns.Add("ExtCardID", System.Type.GetType("System.String"))
        dtTransByStoreID.Columns.Add("CardType", System.Type.GetType("System.String"))
        dtTransByStoreID.Columns.Add("PresentedCardID", System.Type.GetType("System.String"))
        dtTransByStoreID.Columns.Add("HouseHoldID", System.Type.GetType("System.String"))
        dtTransByStoreID.Columns.Add("TerminalNumber", System.Type.GetType("System.String"))
        dtTransByStoreID.Columns.Add("RedemptionCount", System.Type.GetType("System.String"))
        dtTransByStoreID.Columns.Add("RedemptionAmount", System.Type.GetType("System.String"))
        
        Try
        
          If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
            bOpenedRTConnection = True
            MyCommon.Open_LogixRT()
          End If
          If MyCommon.LWHadoConn.State = ConnectionState.Closed Then
            bOpenedWHConnection = True
            MyCommon.Open_LogixWH()
          End If
          If MyCommon.LXSadoConn.State = ConnectionState.Closed Then
            bOpenedXSConnection = True
            MyCommon.Open_LogixXS()
          End If
          If MyCommon.LTRXadoConn.State = ConnectionState.Closed Then
              bOpenedTRXConnection = True
              MyCommon.Open_LogixTRX()
          End If

          If Not IsValidGUID(GUID) Then
            'Wrong GUID
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_GUID
            row.Item("Description") = "Failure: Invalid GUID."
          Else
            'Fetch items based on Transaction Number from TransactionItems Table
              
             MyCommon.QueryStr = "select TH.CustomerPrimaryExtID, Max(TH.TransDate) as TransactionDate, TH.ExtLocationCode, sum(TR.RedemptionAmount) as RedemptionAmount, sum(TR.RedemptionCount) as RedemptionCount, " & _
                      "TH.TerminalNum, TH.LogixTransNum, POSTransNum as TransNum,  TH.CustomerTypeID, TH.PresentedCustomerID, TH.PresentedCardTypeID, TH.HHID, TH.Replayed, isnull(TransTotal,0) as TransTotal  " & _
                      "from TransHist as TH with (NoLock) " & _
                      "Left Outer Join TransRedemptionView as TR with (NoLock) " & _
                      "on TH.LogixTransNum=TR.LogixTransNum " & _
                      "where TH.LogixTransNum='" & TransNum & "' " & _
                      "group by TH.CustomerPrimaryExtID, TH.HHID, TH.CustomerTypeID, TH.PresentedCustomerID, TH.PresentedCardTypeID, TH.LogixTransNum, POSTransNum, TH.TerminalNum, TH.ExtLocationCode, TH.Replayed, TransTotal " & _
                      "order by TH.LogixTransNum" 
              dtTrans = MyCommon.LWH_Select
              If dtTrans.Rows.Count > 0 Then
                drTrans = dtTrans.Rows(0)
                drTransByStoreID = dtTransByStoreID.NewRow()
                drTransByStoreID.Item("LogixTransactionNumber") = TransNum
                
                MyCommon.QueryStr = "Select ItemID, Quantity, Price, Description from TransactionItem with (NoLock) where LogixTransNum='" & TransNum & "';" 
                dtTransItemByStoreID = MyCommon.LTRX_Select
                If dtTransItemByStoreID.Rows.Count > 0 Then
                  '' Item information found
                  row = dtStatus.NewRow()
                  row.Item("StatusCode") = StatusCodes.SUCCESS
                  row.Item("Description") = "Success" 
                  
                  For Each drItems In dtTransItemByStoreID.Rows
                    TransItemInfo = TransItemInfo & drItems.Item("ItemID") & "," & drItems.Item("Description") & "," & drItems.Item("Quantity") & "," & drItems.Item("Price") & ","  & vbcrlf
                  Next drItems
                  drTransByStoreID.Item("TransactionItemInfo") = "<![CDATA[" & TransItemInfo & "]]>"
                Else
                  'No records found
                  row = dtStatus.NewRow()
                  row.Item("StatusCode") = StatusCodes.NOTFOUND_TRANSACTIONITEMS
                  row.Item("Description") = "Failure: No Transaction Items Found" 
                End If
                
                ExtLocationCode = drTrans.Item("ExtLocationCode")
                MyCommon.QueryStr = "select LocationName from Locations where ExtLocationCode = '" & ExtLocationCode & "';"
                dt = MyCommon.LRT_Select
                If dt.Rows.Count > 0 Then
                  LocationName = dt.Rows(0).Item("LocationName")
                End If
                
                drTransByStoreID.Item("Date") = drTrans.Item("TransactionDate")
                drTransByStoreID.Item("ExtLocationCode") = ExtLocationCode
                drTransByStoreID.Item("LocationName") = LocationName
                    drTransByStoreID.Item("ExtCardID") = drTrans.Item("CustomerPrimaryExtID").ToString()
                drTransByStoreID.Item("CardType") = drTrans.Item("CustomerTypeID")
                drTransByStoreID.Item("PresentedCardID") = drTrans.Item("PresentedCustomerID")
                drTransByStoreID.Item("HouseHoldID") = drTrans.Item("HHID")
                drTransByStoreID.Item("TerminalNumber") = drTrans.Item("TerminalNum")
                drTransByStoreID.Item("RedemptionCount") = drTrans.Item("RedemptionCount")
                drTransByStoreID.Item("RedemptionAmount") = drTrans.Item("RedemptionAmount")
                dtTransByStoreID.Rows.Add(drTransByStoreID)
                dtTransByStoreID.AcceptChanges()
                ResultSet.Tables.Add(dtTransByStoreID.Copy())
              Else
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.NOTFOUND_TRANSACTIONS
                row.Item("Description") = "Failure: No Transactions Found" 
              
              End If 'Valid Transaction
          End If 'Valid GUI
        
        Catch ex As Exception
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.APPLICATION_EXCEPTION
            row.Item("Description") = "Failure: Application " & ex.ToString
        Finally
            If bOpenedWHConnection = True And MyCommon.LWHadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixWH()
            If bOpenedTRXConnection = True And MyCommon.LTRXadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixTRX()
            If bOpenedRTConnection = True And MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If bOpenedXSConnection = True And MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
            
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
        End Try
        Return ResultSet
    End Function

  Function CPELimitsBlown(ByVal iIncentiveID As Integer, ByVal lCustomerPK As Long) As Boolean
    Dim bIncentiveLimitReached As Boolean
    Dim dtOffer As DataTable
    Dim dt2 As DataTable
    bIncentiveLimitReached = false
    'iIncentiveID = dr.Item(0)
    'gather the necessary information to determine if the customer has exceeded their limit
    MyCommon.QueryStr = "select i.incentiveid, i.StartDate, i.EndDate, i.IncentiveName, i.Description, i.employeesonly, " & _
                        "i.employeesexcluded, i.P3DistPeriod, i.P3DistTimeType, i.P3DistQtyLimit, CG1.CustomerGroupID " & _
                        "from CPE_ST_Incentives I with (nolock) " & _
                        "inner join CPE_ST_RewardOptions as r with (nolock) on i.incentiveid=r.incentiveid " & _
                        "inner join CPE_ST_IncentiveCustomerGroups CG1 with (nolock) on CG1.rewardoptionid=r.rewardoptionid " & _
                        "where i.incentiveid = " & iIncentiveID & " and CG1.ExcludedUsers <> 1;"

    dtOffer = MyCommon.LRT_Select
    If dtOffer.Rows.Count > 0 Then
      'see if a customer has exceeded their limit for this Incentive
      ' Limits need to be checked only if Period and limit <> 0
      If ((dtOffer.Rows(0).Item("P3DistPeriod") <> 0) Or (dtOffer.Rows(0).Item("P3DistQtyLimit") <> 0)) Then
        Dim NumberRedemptions As Integer = 0
        Dim NumberHours As Integer = 0
        MyCommon.QueryStr = "select count(RD.DistributionID) as NumberRedemptions" & _
                               " from CPE_RewardDistribution rd" & _
                               " where rd.DistributionDate >= DATEADD(day, -" & dtOffer.Rows(0).Item("P3DistPeriod") & ", convert(datetime, GETDATE()))" & _
                               " and rd.DistributionDate < convert(datetime, GETDATE())" & _
                               " and rd.IncentiveID=" & dtOffer.Rows(0).Item("IncentiveID") & _
                               " and rd.CustomerPK=" & lCustomerPK & ";"
        dt2 = MyCommon.LXS_Select
        NumberRedemptions = dt2.Rows(0).Item("NumberRedemptions")

        MyCommon.QueryStr = "select DATEDIFF(day, MAX(rd.DistributionDate), getdate()) as NumberHours" & _
                               " from CPE_RewardDistribution RD" & _
                               " where rd.IncentiveID=" & dtOffer.Rows(0).Item("IncentiveID") & _
                               " and rd.CustomerPK=" & lCustomerPK & ";"
        dt2 = MyCommon.LXS_Select
        NumberHours = MyCommon.NZ(dt2.Rows(0).Item("NumberHours"),-1)
          
        MyCommon.QueryStr = "dbo.pa_CPE_TestLimits"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@incentiveid", SqlDbType.BigInt).Value = dtOffer.Rows(0).Item("IncentiveID")
        MyCommon.LRTsp.Parameters.Add("@NumberDays", SqlDbType.Int).Value = dtOffer.Rows(0).Item("P3DistPeriod")
        MyCommon.LRTsp.Parameters.Add("@NumberHours", SqlDbType.Int).Value = NumberHours
        MyCommon.LRTsp.Parameters.Add("@LimitNumber", SqlDbType.Int).Value = dtOffer.Rows(0).Item("P3DistQtyLimit")
        MyCommon.LRTsp.Parameters.Add("@TimeType", SqlDbType.Int).Value = dtOffer.Rows(0).Item("P3DistTimeType")
        MyCommon.LRTsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = lCustomerPK
        MyCommon.LRTsp.Parameters.Add("@NumberRedemptions", SqlDbType.Int).Value = NumberRedemptions
        MyCommon.LRTsp.Parameters.Add("@limitexceeded", SqlDbType.Bit).Direction = ParameterDirection.Output
        MyCommon.LRTsp.ExecuteNonQuery()
        If MyCommon.LRTsp.Parameters("@limitexceeded").Value = True Then
          bIncentiveLimitReached = True
        End If
        MyCommon.Close_LRTsp()
      End If '(iPeriod <> 0)
    End If
    return bIncentiveLimitReached
  End Function

End Class
