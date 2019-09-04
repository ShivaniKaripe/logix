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
Imports Copient.CommonInc
Imports Copient.AlternateID
Imports Copient.CustomerLookup
Imports Copient.ConnectorInc

<WebService(Namespace:="http://www.copienttech.com/CustomerFacingWebsite/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class Service
  Inherits System.Web.Services.WebService

  Private MyCommon As New Copient.CommonInc
  Private MyAltID As New Copient.AlternateID

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
    APPLICATION_EXCEPTION = 9999
  End Enum


  <WebMethod()> _
  Public Function CustomerDetails(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As Integer) As System.Data.DataSet
    Dim dt, dt2 As System.Data.DataTable
    Dim dtStatus, dtBalances As System.Data.DataTable
    Dim row As System.Data.DataRow
    Dim ResultSet As New System.Data.DataSet("CustomerDetails")
    Dim CustomerPK As Long = 0
    Dim BalRetMsg As String = ""
    Dim BalRetCode As StatusCodes = StatusCodes.SUCCESS

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
		CustomerID = MyCommon.Pad_ExtCardID(CustomerID,CustomerTypeID)
        
		'Find the Customer PK
        MyCommon.QueryStr = "select CID.CustomerPK from CardIDs as CID with (NoLock) " & _
                            "where ExtCardID='" & CustomerID & "' and CardTypeID in " & _
                            "(select CardTypeID from CardTypes with (NoLock) where CustTypeID=" & CustomerTypeID & ");"
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
          CustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)
          'Get the customer's details
          MyCommon.QueryStr = "select C.CustomerPK, C.FirstName, C.LastName, C.Employee, C.CustomerStatusID, C.CurrYearSTD, C.LastYearSTD, " & _
                              "C.CustomerTypeID, C.HHPK, E.Address, E.City, E.State, E.Zip, E.Country, isnull(E.PhoneAsEntered,'') as Phone, E.Email, E.DOB, '' as HHID, " & _
                              "'" & CustomerID & "' as CustomerID " & _
                              "from Customers as C with (NoLock) " & _
                              "left join CustomerEXT as E with (NoLock) on C.CustomerPK=E.CustomerPK " & _
                              "where C.CustomerPK=" & CustomerPK & ";"
          dt = MyCommon.LXS_Select
          If dt.Rows.Count > 0 Then
            If MyCommon.NZ(dt.Rows(0).Item("HHPK"), 0) > 0 Then
              MyCommon.QueryStr = "select top 1 ExtCardID from CardIDs with (NoLock) where CustomerPK=" & dt.Rows(0).Item("HHPK") & ";"
              dt2 = MyCommon.LXS_Select
              If dt2.Rows.Count > 0 Then
                dt.Rows(0).Item("HHID") = MyCommon.NZ(dt2.Rows(0).Item("ExtCardID"), "")
              End If
            End If
            dtBalances = Send_PointsProgramBalances(CustomerPK, BalRetCode, BalRetMsg)
            If BalRetCode = StatusCodes.SUCCESS Then
              ' send customer details and program balances
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

  <WebMethod()> _
  Public Function CustomerDetails_ByCardID(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As Integer,ByVal CardTypeID As Integer) As System.Data.DataSet
    Dim dt, dt2 As System.Data.DataTable
    Dim dtStatus, dtBalances As System.Data.DataTable
    Dim row As System.Data.DataRow
    Dim ResultSet As New System.Data.DataSet("CustomerDetails")
    Dim CustomerPK As Long = 0
    Dim BalRetMsg As String = ""
    Dim BalRetCode As StatusCodes = StatusCodes.SUCCESS

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
        CustomerID = MyCommon.Pad_ExtCardID(CustomerID,CardTypeID)   
                
        'Find the Customer PK
        MyCommon.QueryStr = "select CID.CustomerPK from CardIDs as CID with (NoLock) " & _
                            "where ExtCardID='" & CustomerID & "' and CardTypeID = " &  CardTypeID & ";"
                            
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
          CustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)
          'Get the customer's details
          MyCommon.QueryStr = "select C.CustomerPK, C.FirstName, C.LastName, C.Employee, C.CustomerStatusID, C.CurrYearSTD, C.LastYearSTD, " & _
                              "C.CustomerTypeID, C.HHPK, E.Address, E.City, E.State, E.Zip, E.Country, isnull(E.PhoneAsEntered,'') as Phone, E.Email, E.DOB, '' as HHID, " & _
                              "'" & CustomerID & "' as CustomerID " & _
                              "from Customers as C with (NoLock) " & _
                              "left join CustomerEXT as E with (NoLock) on C.CustomerPK=E.CustomerPK " & _
                              "where C.CustomerPK=" & CustomerPK & ";"
          dt = MyCommon.LXS_Select
          If dt.Rows.Count > 0 Then
            If MyCommon.NZ(dt.Rows(0).Item("HHPK"), 0) > 0 Then
              MyCommon.QueryStr = "select top 1 ExtCardID from CardIDs with (NoLock) where CustomerPK=" & dt.Rows(0).Item("HHPK") & ";"
              dt2 = MyCommon.LXS_Select
              If dt2.Rows.Count > 0 Then
                dt.Rows(0).Item("HHID") = MyCommon.NZ(dt2.Rows(0).Item("ExtCardID"), "")
              End If
            End If
            dtBalances = Send_PointsProgramBalances(CustomerPK, BalRetCode, BalRetMsg)
            If BalRetCode = StatusCodes.SUCCESS Then
              ' send customer details and program balances
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

  <WebMethod()> _
  Public Function MembershipEdit(ByVal GUID As String, ByVal Mode As String, ByVal CustomerGroupID As String, ByVal CustomerID As String, ByVal CustomerTypeID As Integer) As System.Data.DataSet
    Dim dt As System.Data.DataTable
    Dim dtStatus As System.Data.DataTable
    Dim row As System.Data.DataRow
    Dim ResultSet As New System.Data.DataSet("CustomerDetails")
    Dim CustomerPK As Long

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
		CustomerID = MyCommon.Pad_ExtCardID(CustomerID,CustomerTypeID)
		
        'Find the Customer PK
        MyCommon.QueryStr = "select CID.CustomerPK from CardIDs as CID with (NoLock) " & _
                            "where ExtCardID='" & CustomerID & "' and CardTypeID in " & _
                            "(select CardTypeID from CardTypes with (NoLock) where CustTypeID=" & CustomerTypeID & ");"
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
                row.Item("Description") = "Success: Customer " & CustomerID & " opted into customer group " & CustomerGroupID & "."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
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

    Return ResultSet
  End Function
    <WebMethod()> _
  Public Function MembershipEdit_ByCardID(ByVal GUID As String, ByVal Mode As String, ByVal CustomerGroupID As String, ByVal CustomerID As String, ByVal CustomerTypeID As Integer, ByVal CardTypeID As Integer) As System.Data.DataSet
    Dim dt As System.Data.DataTable
    Dim dtStatus As System.Data.DataTable
    Dim row As System.Data.DataRow
    Dim ResultSet As New System.Data.DataSet("CustomerDetails")
    Dim CustomerPK As Long

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
        CustomerID = MyCommon.Pad_ExtCardID(CustomerID,CardTypeID)   
        'Find the Customer PK
        MyCommon.QueryStr = "select CID.CustomerPK from CardIDs as CID with (NoLock) " & _
                            "where ExtCardID='" & CustomerID & "' and CardTypeID = " & CardTypeID & ";"
                         
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
                row.Item("Description") = "Success: Customer " & CustomerID & " opted into customer group " & CustomerGroupID & "."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
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

    Return ResultSet
  End Function

  <WebMethod()> _
  Public Function OfferList(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As Integer) As System.Data.DataSet
    Dim dt As System.Data.DataTable
    Dim dtOffers As System.Data.DataTable
    Dim dtGroups As System.Data.DataTable
    Dim dtStatus As System.Data.DataTable
    Dim row As System.Data.DataRow
    Dim CustomerPK As Long
    Dim ResultSet As New System.Data.DataSet("OfferList")

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
		CustomerID = MyCommon.Pad_ExtCardID(CustomerID,CustomerTypeID)
		
        'Find the Customer PK
        MyCommon.QueryStr = "select CID.CustomerPK from CardIDs as CID with (NoLock) " & _
                            "where ExtCardID='" & CustomerID & "' and CardTypeID in " & _
                            "(select CardTypeID from CardTypes with (NoLock) where CustTypeID=" & CustomerTypeID & ");"
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
    
  <WebMethod()> _
  Public Function OfferList_ByCardID(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As Integer,ByVal CardTypeID As Integer) As System.Data.DataSet
    Dim dt As System.Data.DataTable
    Dim dtOffers As System.Data.DataTable
    Dim dtGroups As System.Data.DataTable
    Dim dtStatus As System.Data.DataTable
    Dim row As System.Data.DataRow
    Dim CustomerPK As Long
    Dim ResultSet As New System.Data.DataSet("OfferList")

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
        CustomerID = MyCommon.Pad_ExtCardID(CustomerID,CardTypeID)   
                
        'Find the Customer PK
        MyCommon.QueryStr = "select CID.CustomerPK from CardIDs as CID with (NoLock) " & _
                            "where ExtCardID='" & CustomerID & "' and CardTypeID = " & CardTypeID & ";"
                            
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
  
  
  <WebMethod()> _
  Public Function TransactionList(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As Integer) As System.Data.DataSet
    Dim dt As System.Data.DataTable
    Dim dtStatus As System.Data.DataTable
    Dim dtTransactions As System.Data.DataTable
    Dim row As System.Data.DataRow
    Dim ResultSet As New System.Data.DataSet("Transactions")
    Dim CustomerPK As Long = 0
    Dim TransRetMsg As String = ""
    Dim TransRetCode As StatusCodes = StatusCodes.SUCCESS
    
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
        'Bad customer ID8
        row = dtStatus.NewRow()
        row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERID
        row.Item("Description") = "Failure: Invalid customer ID."
        dtStatus.Rows.Add(row)
        dtStatus.AcceptChanges()
        ResultSet.Tables.Add(dtStatus.Copy())
      Else
        'Pad the customer ID
		CustomerID = MyCommon.Pad_ExtCardID(CustomerID,CustomerTypeID)
		
        'Find the Customer PK
        MyCommon.QueryStr = "select CID.CustomerPK from CardIDs as CID with (NoLock) " & _
                            "where ExtCardID='" & CustomerID & "' and CardTypeID in " & _
                            "(select CardTypeID from CardTypes with (NoLock) where CustTypeID=" & CustomerTypeID & ");"
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
          CustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)

          'Put a success into the Status table
          row = dtStatus.NewRow()
          row.Item("StatusCode") = 0
          row.Item("Description") = "Success."
          dtStatus.Rows.Add(row)

          'Get the lists of offers
          dtTransactions = Send_Transactions(CustomerPK, CustomerID)
          dtTransactions.TableName = "Offers"

          'Populate and return the DataSet
          ResultSet.Tables.Add(dtStatus.Copy())
          ResultSet.Tables.Add(dtTransactions.Copy())
          
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
  Public Function TransactionList_ByCardID(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As Integer,ByVal CardTypeID As Integer) As System.Data.DataSet
    Dim dt As System.Data.DataTable
    Dim dtStatus As System.Data.DataTable
    Dim dtTransactions As System.Data.DataTable
    Dim row As System.Data.DataRow
    Dim ResultSet As New System.Data.DataSet("Transactions")
    Dim CustomerPK As Long = 0
    Dim TransRetMsg As String = ""
    Dim TransRetCode As StatusCodes = StatusCodes.SUCCESS
    
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
        'Bad customer ID8
        row = dtStatus.NewRow()
        row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERID
        row.Item("Description") = "Failure: Invalid customer ID."
        dtStatus.Rows.Add(row)
        dtStatus.AcceptChanges()
        ResultSet.Tables.Add(dtStatus.Copy())
      Else
        'Pad the customer ID
        CustomerID = MyCommon.Pad_ExtCardID(CustomerID,CardTypeID)   
        'Find the Customer PK
        MyCommon.QueryStr = "select CID.CustomerPK from CardIDs as CID with (NoLock) " & _
                            "where ExtCardID='" & CustomerID & "' and CardTypeID = " & CardTypeID  & ";"
                            
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
          CustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)

          'Put a success into the Status table
          row = dtStatus.NewRow()
          row.Item("StatusCode") = 0
          row.Item("Description") = "Success."
          dtStatus.Rows.Add(row)

          'Get the lists of offers
          dtTransactions = Send_Transactions(CustomerPK, CustomerID)
          dtTransactions.TableName = "Offers"

          'Populate and return the DataSet
          ResultSet.Tables.Add(dtStatus.Copy())
          ResultSet.Tables.Add(dtTransactions.Copy())
          
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
  Public Function ShopList(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As Integer) As System.Data.DataSet
    Dim dt As System.Data.DataTable
    Dim dtStatus As System.Data.DataTable
    Dim dtTransactions As System.Data.DataTable
    Dim row As System.Data.DataRow
    Dim ResultSet As New System.Data.DataSet("List")
    Dim CustomerPK As Long = 0
    Dim TransRetMsg As String = ""
    Dim TransRetCode As StatusCodes = StatusCodes.SUCCESS
    
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
        'Bad customer ID8
        row = dtStatus.NewRow()
        row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERID
        row.Item("Description") = "Failure: Invalid customer ID."
        dtStatus.Rows.Add(row)
        dtStatus.AcceptChanges()
        ResultSet.Tables.Add(dtStatus.Copy())
      Else
        'Pad the customer ID
		CustomerID = MyCommon.Pad_ExtCardID(CustomerID,CustomerTypeID)
		
        'Find the Customer PK
        MyCommon.QueryStr = "select CID.CustomerPK from CardIDs as CID with (NoLock) " & _
                            "where ExtCardID='" & CustomerID & "' and CardTypeID in " & _
                            "(select CardTypeID from CardTypes with (NoLock) where CustTypeID=" & CustomerTypeID & ");"
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
          CustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)

          'Put a success into the Status table
          row = dtStatus.NewRow()
          row.Item("StatusCode") = 0
          row.Item("Description") = "Success."
          dtStatus.Rows.Add(row)

          'Get the lists of offers
          dtTransactions = Send_List(CustomerPK, CustomerID)
          dtTransactions.TableName = "ListElement"

          'Populate and return the DataSet
          ResultSet.Tables.Add(dtStatus.Copy())
          ResultSet.Tables.Add(dtTransactions.Copy())
          
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
  Public Function ShopList_ByCardID(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As Integer,ByVal CardTypeID As Integer) As System.Data.DataSet
    Dim dt As System.Data.DataTable
    Dim dtStatus As System.Data.DataTable
    Dim dtTransactions As System.Data.DataTable
    Dim row As System.Data.DataRow
    Dim ResultSet As New System.Data.DataSet("List")
    Dim CustomerPK As Long = 0
    Dim TransRetMsg As String = ""
    Dim TransRetCode As StatusCodes = StatusCodes.SUCCESS
    
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
        'Bad customer ID8
        row = dtStatus.NewRow()
        row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERID
        row.Item("Description") = "Failure: Invalid customer ID."
        dtStatus.Rows.Add(row)
        dtStatus.AcceptChanges()
        ResultSet.Tables.Add(dtStatus.Copy())
      Else
        'Pad the customer ID
        CustomerID = MyCommon.Pad_ExtCardID(CustomerID,CardTypeID)   
                
        'Find the Customer PK
        MyCommon.QueryStr = "select CID.CustomerPK from CardIDs as CID with (NoLock) " & _
                            "where ExtCardID='" & CustomerID & "' and CardTypeID = " & CardTypeID & ";"
                          
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
          CustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)

          'Put a success into the Status table
          row = dtStatus.NewRow()
          row.Item("StatusCode") = 0
          row.Item("Description") = "Success."
          dtStatus.Rows.Add(row)

          'Get the lists of offers
          dtTransactions = Send_List(CustomerPK, CustomerID)
          dtTransactions.TableName = "ListElement"

          'Populate and return the DataSet
          ResultSet.Tables.Add(dtStatus.Copy())
          ResultSet.Tables.Add(dtTransactions.Copy())
          
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
  Public Function ShopListDelete(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As Integer, ByVal ListPK As Integer) As System.Data.DataSet
    Dim dt As System.Data.DataTable
    Dim dtStatus As System.Data.DataTable
    Dim dtTransactions As System.Data.DataTable
    Dim row As System.Data.DataRow
    Dim ResultSet As New System.Data.DataSet("List")
    Dim CustomerPK As Long = 0
    Dim TransRetMsg As String = ""
    Dim TransRetCode As StatusCodes = StatusCodes.SUCCESS
    
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
        'Bad customer ID8
        row = dtStatus.NewRow()
        row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERID
        row.Item("Description") = "Failure: Invalid customer ID."
        dtStatus.Rows.Add(row)
        dtStatus.AcceptChanges()
        ResultSet.Tables.Add(dtStatus.Copy())
      Else
        'Pad the customer ID
		CustomerID = MyCommon.Pad_ExtCardID(CustomerID,CustomerTypeID)
		
        'Find the Customer PK
        MyCommon.QueryStr = "select CID.CustomerPK from CardIDs as CID with (NoLock) " & _
                            "where ExtCardID='" & CustomerID & "' and CardTypeID in " & _
                            "(select CardTypeID from CardTypes with (NoLock) where CustTypeID=" & CustomerTypeID & ");"
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
          CustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)

          'Put a success into the Status table
          row = dtStatus.NewRow()
          row.Item("StatusCode") = 0
          row.Item("Description") = "Success."
          dtStatus.Rows.Add(row)

          'delete the list item
          MyCommon.QueryStr = "delete from ShopList where listpk=" & ListPK & ";"
          MyCommon.LXS_Execute()
		  

          'Populate and return the DataSet
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

    Return ResultSet
  End Function
    
     <WebMethod()> _
  Public Function ShopListDelete_ByCardID(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As Integer, ByVal ListPK As Integer, ByVal CardTypeID As Integer) As System.Data.DataSet
    Dim dt As System.Data.DataTable
    Dim dtStatus As System.Data.DataTable
    Dim dtTransactions As System.Data.DataTable
    Dim row As System.Data.DataRow
    Dim ResultSet As New System.Data.DataSet("List")
    Dim CustomerPK As Long = 0
    Dim TransRetMsg As String = ""
    Dim TransRetCode As StatusCodes = StatusCodes.SUCCESS
    
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
        'Bad customer ID8
        row = dtStatus.NewRow()
        row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERID
        row.Item("Description") = "Failure: Invalid customer ID."
        dtStatus.Rows.Add(row)
        dtStatus.AcceptChanges()
        ResultSet.Tables.Add(dtStatus.Copy())
      Else
        'Pad the customer ID
        CustomerID = MyCommon.Pad_ExtCardID(CustomerID,CardTypeID)   
                
        'Find the Customer PK
        MyCommon.QueryStr = "select CID.CustomerPK from CardIDs as CID with (NoLock) " & _
                            "where ExtCardID='" & CustomerID & "' and CardTypeID = " & CardTypeID & ";"
                          
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
          CustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)

          'Put a success into the Status table
          row = dtStatus.NewRow()
          row.Item("StatusCode") = 0
          row.Item("Description") = "Success."
          dtStatus.Rows.Add(row)

          'delete the list item
          MyCommon.QueryStr = "delete from ShopList where listpk=" & ListPK & ";"
          MyCommon.LXS_Execute()
		  

          'Populate and return the DataSet
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

    Return ResultSet
  End Function
  
       <WebMethod()> _
  Public Function ShopListAdd(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As Integer, ByVal NewItem As String) As System.Data.DataSet
    Dim dt As System.Data.DataTable
    Dim dtStatus As System.Data.DataTable
    Dim dtTransactions As System.Data.DataTable
    Dim row As System.Data.DataRow
    Dim ResultSet As New System.Data.DataSet("List")
    Dim CustomerPK As Long = 0
    Dim TransRetMsg As String = ""
    Dim TransRetCode As StatusCodes = StatusCodes.SUCCESS
    
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
        'Bad customer ID8
        row = dtStatus.NewRow()
        row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERID
        row.Item("Description") = "Failure: Invalid customer ID."
        dtStatus.Rows.Add(row)
        dtStatus.AcceptChanges()
        ResultSet.Tables.Add(dtStatus.Copy())
      Else
        'Pad the customer ID
		CustomerID = MyCommon.Pad_ExtCardID(CustomerID,CustomerTypeID)
        'Find the Customer PK
        MyCommon.QueryStr = "select CID.CustomerPK from CardIDs as CID with (NoLock) " & _
                            "where ExtCardID='" & CustomerID & "' and CardTypeID in " & _
                            "(select CardTypeID from CardTypes with (NoLock) where CustTypeID=" & CustomerTypeID & ");"
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
          CustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)

          'Put a success into the Status table
          row = dtStatus.NewRow()
          row.Item("StatusCode") = 0
          row.Item("Description") = "Success."
          dtStatus.Rows.Add(row)

          'delete the list item
          MyCommon.QueryStr = "insert into ShopList (CustomerPK,Item) values (" &   CustomerPK   &" ,' " & NewItem & "');"
          MyCommon.LXS_Execute()
		  

          'Populate and return the DataSet
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

    Return ResultSet
  End Function
    
  <WebMethod()> _
  Public Function ShopListAdd_ByCardID(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As Integer, ByVal NewItem As String,ByVal CardTypeID As Integer) As System.Data.DataSet
    Dim dt As System.Data.DataTable
    Dim dtStatus As System.Data.DataTable
    Dim dtTransactions As System.Data.DataTable
    Dim row As System.Data.DataRow
    Dim ResultSet As New System.Data.DataSet("List")
    Dim CustomerPK As Long = 0
    Dim TransRetMsg As String = ""
    Dim TransRetCode As StatusCodes = StatusCodes.SUCCESS
    
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
        'Bad customer ID8
        row = dtStatus.NewRow()
        row.Item("StatusCode") = StatusCodes.INVALID_CUSTOMERID
        row.Item("Description") = "Failure: Invalid customer ID."
        dtStatus.Rows.Add(row)
        dtStatus.AcceptChanges()
        ResultSet.Tables.Add(dtStatus.Copy())
      Else
        'Pad the customer ID
         CustomerID = MyCommon.Pad_ExtCardID(CustomerID,CardTypeID)   
                
        'Find the Customer PK
        MyCommon.QueryStr = "select CID.CustomerPK from CardIDs as CID with (NoLock) " & _
                            "where ExtCardID='" & CustomerID & "' and CardTypeID = " & CardTypeID  & ";"
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
          CustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)

          'Put a success into the Status table
          row = dtStatus.NewRow()
          row.Item("StatusCode") = 0
          row.Item("Description") = "Success."
          dtStatus.Rows.Add(row)

          'delete the list item
          MyCommon.QueryStr = "insert into ShopList (CustomerPK,Item) values (" &   CustomerPK   &" ,' " & NewItem & "');"
          MyCommon.LXS_Execute()
		  

          'Populate and return the DataSet
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

    Return ResultSet
  End Function
  
  <WebMethod()> _
 Public Function OfferListCM(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As Integer) As System.Data.DataSet
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
		CustomerID = MyCommon.Pad_ExtCardID(CustomerID,CustomerTypeID)
        'Find the Customer PK
        MyCommon.QueryStr = "select CID.CustomerPK, C.HHPK from CardIDs as CID with (NoLock) " & _
                            "inner join Customers as C with (NoLock) on C.CustomerPK=CID.CustomerPK " & _
                            "where CID.ExtCardID='" & CustomerID & "' and CID.CardTypeID in " & _
                            "(select CardTypeID from CardTypes where CustTypeID=" & CustomerTypeID & ");"
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

          bStatus = GetCustomerCmOffers(lCustomerPK, lHouseholdPK, dtOffers, dtPrograms)


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
 Public Function OfferListCM_ByCardID(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As Integer, ByVal CardTypeID As Integer) As System.Data.DataSet
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
        CustomerID = MyCommon.Pad_ExtCardID(CustomerID,CardTypeID)   
                
        'Find the Customer PK
        MyCommon.QueryStr = "select CID.CustomerPK, C.HHPK from CardIDs as CID with (NoLock) " & _
                            "inner join Customers as C with (NoLock) on C.CustomerPK=CID.CustomerPK " & _
                            "where CID.ExtCardID='" & CustomerID & "' and CID.CardTypeID = " & CardTypeID & ";"
  
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

          bStatus = GetCustomerCmOffers(lCustomerPK, lHouseholdPK, dtOffers, dtPrograms)


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
 Public Function TransactionHistoryCM(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As Integer, ByVal StartDate As Date, ByVal EndDate As Date) As System.Data.DataSet
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
        CustomerID = MyCommon.Pad_ExtCardID(CustomerID,CustomerTypeID)
        'Find the Customer PK
        MyCommon.QueryStr = "select CID.CustomerPK, C.HHPK from CardIDs as CID with (NoLock) " & _
                            "inner join Customers as C with (NoLock) on C.CustomerPK=CID.CustomerPK " & _
                            "where CID.ExtCardID='" & CustomerID & "' and CID.CardTypeID in " & _
                            "(select CardTypeID from CardTypes where CustTypeID=" & CustomerTypeID & ");"
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
          dtStoredValues.TableName = "Storedvalues"
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

          bStatus = GetCustomerTransactions(lCustomerPK, lHouseholdPK, CustomerID, CustomerTypeID, StartDate, EndDate, dtTransactions, dtPoints, dtStoredValues)

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
 Public Function TransactionHistoryCM_ByCardID(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As Integer, ByVal StartDate As Date, ByVal EndDate As Date, ByVal CardTypeID As Integer) As System.Data.DataSet
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
        CustomerID = MyCommon.Pad_ExtCardID(CustomerID,CardTypeID)   
                
        'Find the Customer PK
        MyCommon.QueryStr = "select CID.CustomerPK, C.HHPK from CardIDs as CID with (NoLock) " & _
                            "inner join Customers as C with (NoLock) on C.CustomerPK=CID.CustomerPK " & _
                            "where CID.ExtCardID='" & CustomerID & "' and CID.CardTypeID = " & CardTypeID & ";"

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
          dtStoredValues.TableName = "Storedvalues"
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

          bStatus = GetCustomerTransactions(lCustomerPK, lHouseholdPK, CustomerID, CustomerTypeID, StartDate, EndDate, dtTransactions, dtPoints, dtStoredValues)

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
  Public Function PointsBalancesCM(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As Integer) As DataSet
    Dim dt As System.Data.DataTable
    Dim dtBalances As System.Data.DataTable
    Dim dtStatus As System.Data.DataTable
    Dim row As System.Data.DataRow
    Dim lCustomerPK As Long
    Dim lHouseholdPK As Long
    Dim bStatus As Boolean
    Dim ResultSet As New System.Data.DataSet("PointsBalancesCM")

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
      	CustomerID = MyCommon.Pad_ExtCardID(CustomerID,CustomerTypeID)
        'Find the Customer PK
        MyCommon.QueryStr = "select CID.CustomerPK, C.HHPK from CardIDs as CID with (NoLock) " & _
                            "inner join Customers as C with (NoLock) on C.CustomerPK=CID.CustomerPK " & _
                            "where CID.ExtCardID='" & CustomerID & "' and CID.CardTypeID in " & _
                            "(select CardTypeID from CardTypes where CustTypeID=" & CustomerTypeID & ");"
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


          dtBalances = New DataTable("PointsProgram")
          dtBalances.Columns.Add("ProgramID", System.Type.GetType("System.Int64"))
          dtBalances.Columns.Add("ProgramName", System.Type.GetType("System.String"))
          dtBalances.Columns.Add("Category", System.Type.GetType("System.String"))
          dtBalances.Columns.Add("Balance", System.Type.GetType("System.Decimal"))

          bStatus = GetPointsBalances(lCustomerPK, lHouseholdPK, dtBalances)


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
  Public Function PointsBalancesCM_ByCardID(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As Integer, ByVal CardTypeID As Integer) As DataSet
    Dim dt As System.Data.DataTable
    Dim dtBalances As System.Data.DataTable
    Dim dtStatus As System.Data.DataTable
    Dim row As System.Data.DataRow
    Dim lCustomerPK As Long
    Dim lHouseholdPK As Long
    Dim bStatus As Boolean
    Dim ResultSet As New System.Data.DataSet("PointsBalancesCM")

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
        CustomerID = MyCommon.Pad_ExtCardID(CustomerID,CardTypeID)   
                
        'Find the Customer PK
        MyCommon.QueryStr = "select CID.CustomerPK, C.HHPK from CardIDs as CID with (NoLock) " & _
                            "inner join Customers as C with (NoLock) on C.CustomerPK=CID.CustomerPK " & _
                            "where CID.ExtCardID='" & CustomerID & "' and CID.CardTypeID = " & CardTypeID  & ";"
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


          dtBalances = New DataTable("PointsProgram")
          dtBalances.Columns.Add("ProgramID", System.Type.GetType("System.Int64"))
          dtBalances.Columns.Add("ProgramName", System.Type.GetType("System.String"))
          dtBalances.Columns.Add("Category", System.Type.GetType("System.String"))
          dtBalances.Columns.Add("Balance", System.Type.GetType("System.Decimal"))

          bStatus = GetPointsBalances(lCustomerPK, lHouseholdPK, dtBalances)


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
  Public Function StoredValueBalancesCM(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As Integer, ByVal AboutToExpireDays As Integer) As DataSet
    Dim dt As System.Data.DataTable
    Dim dtBalances As System.Data.DataTable
    Dim dtStatus As System.Data.DataTable
    Dim row As System.Data.DataRow
    Dim lCustomerPK As Long
    Dim lHouseholdPK As Long
    Dim bStatus As Boolean
    Dim ResultSet As New System.Data.DataSet("StoredValueBalancesCM")

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
      	CustomerID = MyCommon.Pad_ExtCardID(CustomerID,CustomerTypeID)
		
        'Find the Customer PK
        MyCommon.QueryStr = "select CID.CustomerPK, C.HHPK from CardIDs as CID with (NoLock) " & _
                            "inner join Customers as C with (NoLock) on C.CustomerPK=CID.CustomerPK " & _
                            "where CID.ExtCardID='" & CustomerID & "' and CID.CardTypeID in " & _
                            "(select CardTypeID from CardTypes where CustTypeID=" & CustomerTypeID & ");"
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
  Public Function StoredValueBalancesCM_ByCardID(ByVal GUID As String, ByVal CustomerID As String, ByVal CustomerTypeID As Integer, ByVal AboutToExpireDays As Integer, ByVal CardTypeID As Integer) As DataSet
    Dim dt As System.Data.DataTable
    Dim dtBalances As System.Data.DataTable
    Dim dtStatus As System.Data.DataTable
    Dim row As System.Data.DataRow
    Dim lCustomerPK As Long
    Dim lHouseholdPK As Long
    Dim bStatus As Boolean
    Dim ResultSet As New System.Data.DataSet("StoredValueBalancesCM")

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
        CustomerID = MyCommon.Pad_ExtCardID(CustomerID,CardTypeID)   
                
        'Find the Customer PK
        MyCommon.QueryStr = "select CID.CustomerPK, C.HHPK from CardIDs as CID with (NoLock) " & _
                            "inner join Customers as C with (NoLock) on C.CustomerPK=CID.CustomerPK " & _
                            "where CID.ExtCardID='" & CustomerID & "' and CID.CardTypeID = " & CardTypeID & ";"

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


  Private Function GetCustomerCmOffers(ByVal lCustomerPK As Long, ByVal lHouseholdPK As Long, ByRef dtOffers As DataTable, ByRef dtPrograms As DataTable) As Boolean
    Const sDateFormatBusStart As String = "yyyy-MM-ddT23:59:59"
    Const sDateFormatBusEnd As String = "yyyy-MM-ddT00:00:00"
    Const sDateFormat As String = "yyyy-MM-ddTHH:mm:ss"

    Dim dt As DataTable = Nothing
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
      sGroupCustomerList += ",1,2)"
    Else
      ' Member, but no specific groups assigned
      sGroupCustomerList = "(1,2)"
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
      If dtOffers.Rows.Count > 0 Then dtOffers.AcceptChanges()
      If dtPrograms.Rows.Count > 0 Then dtPrograms.AcceptChanges()
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
        MyCommon.QueryStr = "select ExtCardID from CardIDs with (NoLock) where CustomerPK = " & lHouseholdPK & ";"
        dtXS = MyCommon.LXS_Select
        If dtXS.Rows.Count > 0 Then
          sHouseHoldID = MyCommon.NZ(dtXS.Rows(0).Item(0), "")
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
    Else
      bstatus = False
    End If

    Return bstatus
  End Function

  Private Function GetPointsBalances(ByVal lCustomerPK As Long, ByVal lHouseholdPK As Long, ByRef dtBalances As DataTable) As Boolean
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

    MyCommon.QueryStr = "select ProgramID, Sum(Amount) from Points with (NoLock) where CustomerPK = " & lUseCustomerPK & _
                        " group by ProgramID order by ProgramID;"
    dt = MyCommon.LXS_Select
    If dt.Rows.Count > 0 Then
      For Each dr In dt.Rows
        decAmount = Decimal.Parse(dr.Item(1))
        If decAmount > 0.0 Then
          sProgramId = dr.Item(0)
          sAmount = dr.Item(1)
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

    Return GraphicsFileName
  End Function

  Private Function RetrievePrintedMessage(ByVal OfferID As Long) As String
    'This function is used by Send_XMLCurrentOffers and Send_XMLGroupOffers to get the text of an offer's printed message  
    Dim PMsgBuf As New StringBuilder()
    Dim rst As System.Data.DataTable

    MyCommon.Open_LogixRT()
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
    MyCommon.Close_LogixRT()

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
    Dim CustomerGroups As String = "0"
    Dim CgBuf As New StringBuilder()

    If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
    If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

    'Create a new datatable to hold the results we'll be assembling
    dtOffers = New DataTable
    dtOffers.Columns.Add("OfferID", System.Type.GetType("System.Int64"))
    dtOffers.Columns.Add("Name", System.Type.GetType("System.String"))
    dtOffers.Columns.Add("Description", System.Type.GetType("System.String"))
    dtOffers.Columns.Add("OfferCategoryID", System.Type.GetType("System.Int32"))
    dtOffers.Columns.Add("StartDate", System.Type.GetType("System.DateTime"))
    dtOffers.Columns.Add("EndDate", System.Type.GetType("System.DateTime"))
    dtOffers.Columns.Add("CustomerGroupID", System.Type.GetType("System.Int64"))
    dtOffers.Columns.Add("AllowOptOut", System.Type.GetType("System.Boolean"))
    dtOffers.Columns.Add("EmployeesOnly", System.Type.GetType("System.Boolean"))
    dtOffers.Columns.Add("EmployeesExcluded", System.Type.GetType("System.Boolean"))
    dtOffers.Columns.Add("Points", System.Type.GetType("System.Int32"))
    dtOffers.Columns.Add("Accumulation", System.Type.GetType("System.Decimal"))
    dtOffers.Columns.Add("BodyText", System.Type.GetType("System.String"))
    dtOffers.Columns.Add("Graphic", System.Type.GetType("System.String"))

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
      'Build a list of customer group IDs for this customer
      For Each row In rst.Rows
        If (CgBuf.Length > 0) Then CgBuf.Append(",")
        CgBuf.Append(MyCommon.NZ(row.Item("CustomerGroupID"), 0))
      Next
      CustomerGroups = CgBuf.ToString

      'The customer's in at least one group, so for each one we'll grab the associated offer(s)
      For Each row In rst.Rows
        MyCommon.QueryStr = "select distinct O.OfferID, O.ExtOfferID, O.IsTemplate, O.CMOADeployStatus, O.StatusFlag, O.OddsOfWinning, O.InstantWin, " & _
                            "O.Name, O.Description, O.OfferCategoryID, O.ProdStartDate, O.ProdEndDate, 0 as AllowOptOut, O.EmployeeFiltering as EmployeesOnly, O.NonEmployeesOnly as EmployeesExcluded, LinkID, OID.EngineID " & _
                            "from Offers as O with (NoLock) " & _
                            "left join OfferConditions as OC with (NoLock) on OC.OfferID=O.OfferID " & _
                            "inner join OfferIDs as OID with (NoLock) on OID.OfferID=O.OfferID " & _
                            "where O.IsTemplate=0 and O.Deleted=0 and O.CMOADeployStatus=1 and OC.Deleted=0 and OC.ConditionTypeID=1 " & _
                            "and O.DisabledOnCFW=0 and ProdEndDate>'" & Today.AddDays(-1).ToString & "' and LinkID=" & row.Item("CustomerGroupID") & _
                            "union all " & _
                            "select distinct I.IncentiveID, I.ClientOfferID, I.IsTemplate, I.CPEOADeployStatus, I.StatusFlag, 0 as OddsOfWinning, 0 as InstantWin, " & _
                            "I.IncentiveName, Convert(nvarchar(2000),I.Description) as Description, I.PromoClassID as OfferCategoryID, I.StartDate, I.EndDate, I.AllowOptOut, I.EmployeesOnly, I.EmployeesExcluded, ICG.CustomerGroupID, OID.EngineID " & _
                            "from CPE_Incentives as I with (NoLock) " & _
                            "left join CPE_RewardOptions as RO with (NoLock) on I.IncentiveID=RO.IncentiveID and RO.TouchResponse=0 and RO.Deleted=0 " & _
                            "left join CPE_IncentiveCustomerGroups as ICG with (NoLock) on RO.RewardOptionID=ICG.RewardOptionID and ICG.ExcludedUsers=0 " & _
                            "inner join OfferIDs as OID with (NoLock) on OID.OfferID=I.IncentiveID " & _
                            "where (I.IsTemplate=0 and I.Deleted=0 and ICG.Deleted=0) " & _
                            "and I.DisabledOnCFW=0 and I.EndDate>'" & Today.AddDays(-1).ToString & "' and CustomerGroupID=" & row.Item("CustomerGroupID") & ";"
        rst2 = MyCommon.LRT_Select

        'Set the general info for each offer found
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

          'Filter out the website offers
          MyCommon.QueryStr = "select OfferID from OfferIDs with (NoLock) where OfferID=" & OfferID & " and EngineID=3;"
          rstWeb = MyCommon.LRT_Select

          'Filter out the offers where the customer is in the excluded customer group
          MyCommon.QueryStr = "select ExcludedID from OfferConditions with (NoLock) " & _
                              "where OfferID=" & OfferID & " and ExcludedID in (" & CustomerGroups & ") " & _
                              "union " & _
                              "select CustomerGroupID as ExcludedID from CPE_IncentiveCustomerGroups ICG with (NoLock) " & _
                              "inner join CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID=ICG.RewardOptionID " & _
                              "where ICG.Deleted=0 and RO.Deleted=0 and RO.IncentiveID=" & OfferID & " and ExcludedUsers=1 " & _
                              "and CustomerGroupID in (" & CustomerGroups & ");"
          rstExcluded = MyCommon.LRT_Select

          If (rstWeb.Rows.Count = 0 AndAlso rstExcluded.Rows.Count = 0) Then

            'Find the name of the associated (and non-excluding) location group
            MyCommon.QueryStr = "select OL.OfferID, OL.LocationGroupID, OL.Excluded, LG.Name from OfferLocations as OL with (NoLock) " & _
                                "left join LocationGroups as LG with (NoLock) on LG.LocationGroupID=OL.LocationGroupID " & _
                                "where OL.OfferID=" & OfferID & " and OL.Excluded=0;"
            rst3 = MyCommon.LRT_Select

            'Find any associated points programs
            MyCommon.QueryStr = "select O.Offerid, LinkID, ProgramName, PP.ProgramID, PromoVarID from OfferRewards as OFR with (NoLock) " & _
                                "left join RewardPoints as RP with (NoLock) on RP.RewardPointsID=OFR.LinkID " & _
                                "left join PointsPrograms as PP with (NoLock) on RP.ProgramID=PP.ProgramID " & _
                                "left join Offers as O with (NoLock) on O.OfferID=OFR.OfferID " & _
                                "where (RewardTypeID=2 and O.Deleted=0 and OFR.Deleted=0) " & _
                                "and RP.ProgramID is not null " & _
                                "and O.OfferID=" & OfferID & _
                                " union " & _
                                "select " & OfferID & " as OfferID, D.OutputID, PP.ProgramName, PP.ProgramID, PP.PromoVarID " & _
                                "from CPE_Deliverables D with (NoLock) inner join CPE_DeliverablePoints DP with (NoLock) on D.OutputID=DP.PKID " & _
                                "inner join PointsPrograms PP with (NoLock) on DP.ProgramID=PP.ProgramID " & _
                                "where D.RewardOptionID in (select RO.RewardOptionID from CPE_RewardOptions RO with (NoLock) where IncentiveID=" & OfferID & ") " & _
                                "and D.Deleted=0 and DP.Deleted=0 and PP.Deleted=0 and D.DeliverableTypeID=8;"
            rst4 = MyCommon.LRT_Select
            For Each row4 In rst4.Rows
              ProgramID = row4.Item("ProgramID")
              ProgramName = MyCommon.NZ(row4.Item("ProgramName"), "unknown").ToString.Replace(",", " ")
            Next

            If (ProgramName <> "" And ProgramID <> "") Then
              'Find the balance in the points program
              For Each row4 In rst4.Rows
                ProgID = MyCommon.NZ(row4.Item("ProgramID"), -1)
                MyCommon.QueryStr = "select Amount from Points with (NoLock) where CustomerPK=" & CustomerPK & " and ProgramID=" & ProgID
                rst5 = MyCommon.LXS_Select
                If (rst5.Rows.Count > 0) Then
                  Amount = MyCommon.NZ(rst5.Rows(0).Item("Amount"), 0)
                Else
                  Amount = 0
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
            row.Item("CustomerGroupID") = CustomerGroupID
            row.Item("AllowOptOut") = AllowOptOut
            row.Item("EmployeesOnly") = EmployeesOnly
            row.Item("EmployeesExcluded") = EmployeesExcluded
            row.Item("Points") = Amount
            row.Item("Accumulation") = AccumulationBalance
            row.Item("BodyText") = PrintedMessage
            row.Item("Graphic") = GraphicsFileName
            dtOffers.Rows.Add(row)
            ProgramID = ""
            ProgramName = ""
          End If
        Next
        If dtOffers.Rows.Count > 0 Then dtOffers.AcceptChanges()
      Next
    End If

    Return dtOffers
  End Function

  Private Function Send_XMLGroupOffers(ByVal CustomerPK As Long) As DataTable
    Dim rst As System.Data.DataTable
    Dim row As System.Data.DataRow
    Dim rst2 As System.Data.DataTable
    Dim row2 As System.Data.DataRow
    Dim rst3 As System.Data.DataTable
    Dim rstWeb As System.Data.DataTable
    Dim rstCG As System.Data.DataTable
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
        rstCG.Rows.Add(New String() {"1"})
        rstCG.Rows.Add(New String() {"2"})
        rowCount = rstCG.Rows.Count
        For i = 0 To rowCount - 1
          CustomerGroups.Append(MyCommon.NZ(rstCG.Rows(i).Item("CustomerGroupID"), -1))
          If (i < rowCount - 1) Then CustomerGroups.Append(",")
        Next
        MyCommon.QueryStr = "select I.IncentiveID, I.IncentiveName, I.Description, I.PromoClassID as OfferCategoryID, I.StartDate, I.EndDate, ICG.CustomerGroupID, " & _
                            "ICG.ExcludedUsers, D.OutputID as RewardGroup, I.AllowOptOut, I.EmployeesOnly, I.EmployeesExcluded, RO.RewardOptionID " & _
                            "from CPE_Incentives as I with (NoLock) " & _
                            "inner join OfferIDs as OID with (NoLock) on I.IncentiveID=OID.OfferID " & _
                            "inner join CPE_RewardOptions as RO with (NoLock) on RO.IncentiveID=I.IncentiveID and RO.TouchResponse=0 and RO.Deleted=0 " & _
                            "inner join CPE_Deliverables as D with (NoLock) on D.RewardOptionID=RO.RewardOptionID and D.Deleted=0 and DeliverableTypeID=5 and RewardOptionPhase=3 " & _
                            "inner join CPE_IncentiveCustomerGroups as ICG with (NoLock) on ICG.RewardOptionID=RO.RewardOptionID " & _
                            " and ICG.Deleted=0 and ICG.CustomerGroupID in (" & CustomerGroups.ToString & ") and ICG.ExcludedUsers=0 " & _
                            "where I.Deleted=0 And I.StatusFlag=0 and I.EndDate>='" & Today.AddDays(-1).ToString & "' and OID.EngineID=3 " & CPECurrentOffers & ";"
        rst2 = MyCommon.LRT_Select

        'Set the general info for each offer found
        If (rst2.Rows.Count > 0) Then
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

            'Filter out offers where the customer is already a member of the reward group
            MyCommon.QueryStr = "select MembershipID from GroupMembership with (NoLock) where CustomerPK=" & CustomerPK & " " & _
                                "and Deleted=0 and CustomerGroupID=" & RewardGroupID & ";"
            rstWeb = MyCommon.LXS_Select
            OptOutOffer = (rstWeb.Rows.Count > 0)

            'Check if the customer is in a group that is excluded from this offer
            MyCommon.QueryStr = "select CustomerGroupID from CPE_IncentiveCustomerGroups ICG with (NoLock) " & _
                                "where RewardOptionID=" & ROID & " and CustomerGroupID in (" & CustomerGroups.ToString & ") and ExcludedUsers=1 and Deleted=0;"
            rstWeb = MyCommon.LRT_Select
            ExcludedFromOffer = (rstWeb.Rows.Count > 0)

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

            If (PointsConditionOK) AndAlso (Not ExcludedFromOffer And ((Not OptOutOffer) Or (OptOutOffer And AllowOptOut))) Then
              'Find the name of the associated (and non-excluding) location group
              MyCommon.QueryStr = "select OL.OfferID, OL.LocationGroupID, OL.Excluded, LG.Name from OfferLocations as OL with (NoLock) " & _
                                  "left join LocationGroups as LG with (NoLock) on LG.LocationGroupID=OL.LocationGroupID " & _
                                  "where OL.OfferID=" & OfferID & " and OL.Excluded=0;"
              rst3 = MyCommon.LRT_Select

              'Find the name of the associated (and non-excluding) location group
              MyCommon.QueryStr = "select OL.OfferID, OL.LocationGroupID, OL.Excluded, LG.Name from OfferLocations as OL with (NoLock) " & _
                                  "left join LocationGroups as LG with (NoLock) on LG.LocationGroupID=OL.LocationGroupID " & _
                                  "where OL.OfferID=" & OfferID & " and OL.Excluded=0;"
              rst3 = MyCommon.LRT_Select

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
              dtGroups.Rows.Add(row)
            End If
          Next
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

  Private Function Send_Transactions(ByVal CustomerPK As Long, ByVal CustomerID As String) As DataTable
    'This function is used by OfferList and returns a list of customer offers
    Dim dt As System.Data.DataTable
    Dim dtTransactions As System.Data.DataTable
    Dim row As System.Data.DataRow
    Dim retString As New StringBuilder
    Dim MyLookup As New Copient.CustomerLookup
    Dim BalRetCode As Copient.CustomerLookup.RETURN_CODE
    Dim Balances(-1) As Copient.CustomerLookup.PointsBalance
    Dim TransactionDate As New DateTime
    Dim ExtLocationCode As String = ""
    Dim RedemptionAmount As Decimal = 0
    Dim RedemptionCount As Integer = 0
    Dim TerminalNum As String = ""
    Dim TransNum As String = ""
    Dim LogixTransNum As String = ""
    Dim OfferID As Long = 0
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
                        "from ShopList as List with (NoLock) where CustomerPK = " & CustomerPK  & _
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
End Class