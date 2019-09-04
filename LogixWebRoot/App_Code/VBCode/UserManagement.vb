Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.IO
Imports System.Xml
Imports System.Xml.Schema
Imports System.Data
Imports System.Data.SqlClient
Imports System
Imports System.Text
Imports Copient.CommonInc
Imports CMS.CryptLib
Imports Copient.LogixInc
Imports Copient.ConnectorInc
Imports CMS.AMS
Imports CMS.AMS.Contract

' To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line.
' <System.Web.Script.Services.ScriptService()> _
<WebService(Namespace:="http://www.copienttech.com/UserManagement/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class UserManagement
    Inherits System.Web.Services.WebService

    Private UMLogFile As String = "UserManagementWSLog." & Date.Now.ToString("yyyyMMdd") & ".txt"

    Dim MyCommon As New Copient.CommonInc
    ''insert  into connectors table
    Private Const CONNECTOR_ID As Integer = 63

    Public Enum StatusCodes As Integer
        SUCCESS = 0
        INVALID_GUID = 1
        INVALID_USERNAME = 2
        INVALID_NAME = 3
        INVALID_FIRSTNAME = 4
        INVALID_LASTNAME = 5
        INVALID_EMAILD = 6
        INVALID_USERID = 7
        INVALID_AUTHENTICATION = 8
        INVALID_DESCRIPTION = 9
        INVALID_ROLEID = 10
        INVALID_PASSWORDMISMATCH = 11
        INVALID_PASSWORD = 12
        INVALID_PASSWORDCON = 13
        FAILED_OPTIN = 14
        NOTFOUND_RECORDS = 15
        INVALID_ROLENAMES = 16
        INVALID_EMPLOYEEID = 17
        INVALID_EXTBANNERID = 18
        INVALID_BANNERNAME = 19
        INVALID_PASSWORDREUSED = 20
        APPLICATION_EXCEPTION = 9999
    End Enum


    Private Sub InitApp()
        MyCommon.AppName = "UserManagement.asmx"

        Try
        Catch eXmlSch As XmlSchemaException
        Catch ex As Exception

        End Try
    End Sub

    <WebMethod()> _
    Public Function GetUserDetailsByID(ByVal GUID As String, ByVal UserId As String) As DataSet

        InitApp()
        Dim iUserId As Integer = -1
        Try
            iUserId = CInt(UserId)
        Catch ex As Exception
            iUserId = -1
        End Try
        Return _GetUserDetailsByID(GUID, iUserId)

    End Function

    <WebMethod()> _
    Public Function GetUserDetailsByName(ByVal GUID As String, ByVal Username As String) As DataSet

        InitApp()
        Return _GetUserDetailsByName(GUID, Username)

    End Function

    <WebMethod()> _
    Public Function GetUserRolesByName(ByVal GUID As String, ByVal Username As String) As DataSet
        InitApp()
        Return _GetUserRolesByName(GUID, Username)
    End Function

    <WebMethod()> _
    Public Function GetUserRolesByID(ByVal GUID As String, ByVal UserId As String) As DataSet
        InitApp()
        Dim lUserId As Integer = -1
        Try
            If UserId.Length > 0 Then If IsNumeric(UserId) = True Then lUserId = CInt(UserId)
        Catch ex As Exception
            lUserId = -1
        End Try
        Return _GetUserRolesByID(GUID, lUserId)
    End Function

    <WebMethod()> _
    Public Function GetRolePermissions(ByVal GUID As String, ByVal RoleId As String) As DataSet

        InitApp()
        Dim lRoleId As Integer = -1
        Try
            If RoleId.Length > 0 Then If IsNumeric(RoleId) = True Then lRoleId = CInt(RoleId)
        Catch ex As Exception
            lRoleId = -1
        End Try
        Return _GetRolePermissionsList(GUID, lRoleId)

    End Function

    <WebMethod()> _
    Public Function GetUsersRoles(ByVal GUID As String) As DataSet

        InitApp()
        Return _GetUsersRoles(GUID)

    End Function

    <WebMethod()> _
    Public Function GetUsersList(ByVal GUID As String) As DataSet

        InitApp()
        Return _GetUserList(GUID)

    End Function

    <WebMethod()> _
    Public Function AddUser(ByVal GUID As String, ByVal Username As String, ByVal FirstName As String, ByVal LastName As String, ByVal Password As String, _
                           ByVal Email As String, ByVal RoleId As String, ByVal EmployeeId As String) As DataSet
        InitApp()
        Return _AddUser(GUID, Username, FirstName, LastName, Password, Email, RoleId, EmployeeId)

    End Function

    <WebMethod()> _
    Public Function AddUserRoles(ByVal GUID As String, ByVal Username As String, ByVal RoleNames As String) As DataSet
        InitApp()
        Return _AddUserRoles(GUID, Username, RoleNames)

    End Function

    'AMSPS-2157
    <WebMethod()> _
    Public Function AddUserWithExtBannerId(ByVal GUID As String, ByVal Username As String, ByVal FirstName As String, ByVal LastName As String, ByVal Password As String, _
                           ByVal Email As String, ByVal RoleId As String, ByVal EmployeeId As String, ByVal ExtBannerId As String) As DataSet

        InitApp()
        'Method to add user with External BannerID
        Dim MethodName As String = "AddUserWithExtBannerId"
        Return _AddUserWithBanner(GUID, Username, FirstName, LastName, Password, Email, RoleId, EmployeeId, ExtBannerId, MethodName)

    End Function

    'AMSPS-2157
    <WebMethod()> _
    Public Function AddUserWithBannerName(ByVal GUID As String, ByVal Username As String, ByVal FirstName As String, ByVal LastName As String, ByVal Password As String, _
                           ByVal Email As String, ByVal RoleId As String, ByVal EmployeeId As String, ByVal BannerName As String) As DataSet

        InitApp()
        'Method to add user with Banner Name
        Dim MethodName As String = "AddUserWithBannerName"
        Return _AddUserWithBanner(GUID, Username, FirstName, LastName, Password, Email, RoleId, EmployeeId, BannerName, MethodName)

    End Function

    <WebMethod()> _
    Public Function ModifyUser(ByVal GUID As String, ByVal Username As String, ByVal FirstName As String, ByVal LastName As String, _
                             ByVal Email As String, ByVal RoleId As String, ByVal EmployeeId As String) As DataSet
        InitApp()
        Return _ModifyUser(GUID, Username, FirstName, LastName, Email, RoleId, EmployeeId)

    End Function

    <WebMethod()> _
    Public Function DeleteUserById(ByVal GUID As String, ByVal UserId As String) As DataSet

        InitApp()
        Dim iUserId As Integer = -1
        Try
            iUserId = CInt(UserId)
        Catch ex As Exception
            iUserId = -1
        End Try
        Return _DeleteUserById(GUID, iUserId)

    End Function

    <WebMethod()> _
    Public Function DeleteUserByName(ByVal GUID As String, ByVal Username As String) As DataSet

        InitApp()
        Return _DeleteUserByName(GUID, Username)

    End Function

    <WebMethod()> _
    Public Function DeleteUserByAuthentication(ByVal GUID As String, ByVal DeleteUserName As String, ByVal AdminUserName As String, _
                              ByVal Password As String) As DataSet
        InitApp()
        Return _DeleteUserByAuthentication(GUID, AdminUserName, Password, DeleteUserName)

    End Function

    <WebMethod()> _
    Public Function ChangePassword(ByVal GUID As String, ByVal Username As String, ByVal OldPassword As String, _
                      ByVal NewPassword As String) As DataSet
        InitApp()
        Return _ChangePassword(GUID, Username, OldPassword, NewPassword)

    End Function

    Private Function _AddUser(ByVal GUID As String, ByVal Username As String, ByVal FirstName As String, ByVal LastName As String, ByVal Password As String, _
                            ByVal Email As String, ByVal RoleId As String, ByVal EmployeeId As String) As DataSet

        Dim ResultSet As New System.Data.DataSet("UserManagement")
        Dim MethodName As String = "AddUser"
        Dim dtStatus As DataTable
        Dim rst As DataTable
        Dim row As System.Data.DataRow
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = ""
        Dim u_UID As Integer
        Dim MyCryptlib As New CMS.CryptLib
        Dim HashLib As New CMS.HashLib.CryptLib
        Dim UpdateString As String = ""

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        If isValidInputText(GUID) = False Then
            'GUID Value is empty
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_GUID
            row.Item("Description") = "Failure.Invalid GUID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _AddUser = ResultSet
            Exit Function
        End If
        If Username.Trim.Length = 0 Or Username.Trim.Length > 50 Then
            'Username Value is empty or invalid
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_USERNAME
            row.Item("Description") = "Failure.Invalid Username."
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _AddUser = ResultSet
            Exit Function
        End If
        If FirstName.Trim.Length = 0 And LastName.Trim.Length = 0 Then
            'FirstName Value is empty or invalid
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_NAME
            row.Item("Description") = "Failure.Enter Firstname or Lastname."
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _AddUser = ResultSet
            Exit Function
        End If
        If FirstName.Trim <> "" Then
            If FirstName.Length > 50 Then
                'FirstName Value is empty or invalid
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_FIRSTNAME
                row.Item("Description") = "Failure.Invalid FirstName."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                _AddUser = ResultSet
                Exit Function
            End If
        End If
        If LastName.Trim <> "" Then
            If LastName.Length > 50 Then
                'FirstName Value is empty or invalid
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_LASTNAME
                row.Item("Description") = "Failure.Invalid LastName."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                _AddUser = ResultSet
                Exit Function
            End If
        End If
        If Password.Trim.Length = 0 Or Password.Trim.Length > 20 Then
            'Password Value is empty or invalid
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_PASSWORD
            row.Item("Description") = "Failure.Invalid Password."
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _AddUser = ResultSet
            Exit Function
        End If
        If Email <> "" Then
            If EmailAddressCheck(Email) = False Then
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_EMAILD
                row.Item("Description") = "Failure.Invalid Email."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                _AddUser = ResultSet
                Exit Function
            End If
        End If
        If RoleId <> "" Then
            Dim lRoleID As Long = 0
            Try
                lRoleID = RoleId.Replace(",", "")
            Catch ex As Exception
                lRoleID = -1
            End Try
            If lRoleID = -1 Then
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_ROLEID
                row.Item("Description") = "Failure.Invalid RoleID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                _AddUser = ResultSet
                Exit Function
            End If
        End If
        If EmployeeId <> "" Then
            Dim IsAlpha As Boolean
            If System.Text.RegularExpressions.Regex.IsMatch(EmployeeId, "^[a-zA-Z0-9]+$") Then
                IsAlpha = True
            Else
                IsAlpha = False
            End If
            If (IsAlpha) Then
                EmployeeId = MyCommon.CleanString(EmployeeId)
                MyCommon.QueryStr = "select EmployeeId from AdminUsers with (NoLock) where EmployeeId=N'" & EmployeeId & "'"
                rst = MyCommon.LRT_Select
                If rst.Rows.Count > 0 Then
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = StatusCodes.INVALID_EMPLOYEEID
                    row.Item("Description") = "Failure. EmployeeId already exists."
                    dtStatus.Rows.Add(row)
                    dtStatus.AcceptChanges()
                    ResultSet.Tables.Add(dtStatus.Copy())
                    _AddUser = ResultSet
                    Exit Function
                End If
                UpdateString = ", EmployeeId = N'" & EmployeeId & "'"
            Else
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_EMPLOYEEID
                row.Item("Description") = "Failure. EmployeeId can have only alphanumeric characters"
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                _AddUser = ResultSet
                Exit Function
            End If
        End If
        Try
            Using MyCommon.LRTadoConn
                If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                    MyCommon.Open_LogixRT()
                End If

                If Not IsValidGUID(GUID, MethodName) Then
                    RetCode = StatusCodes.INVALID_GUID
                    RetMsg = "Failure.Invalid GUID"
                Else
                    MyCommon.QueryStr = "Select UserName from AdminUsers with (NoLock) where UserName=@Username"
                    MyCommon.DBParameters.Add("@Username", SqlDbType.NVarChar).Value = Username
                    rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                    If rst.Rows.Count > 0 Then
                        RetCode = StatusCodes.INVALID_USERNAME
                        RetMsg = "Failure. Username already exists."
                    Else
                        MyCommon.QueryStr = "dbo.pt_AdminUsers_Insert"
                        MyCommon.Open_LRTsp()
                        MyCommon.LRTsp.Parameters.Add("@UserName ", SqlDbType.NVarChar, 50).Value = Username
                        MyCommon.LRTsp.Parameters.Add("@AdminUserID", SqlDbType.Int).Direction = ParameterDirection.Output
                        MyCommon.LRTsp.ExecuteNonQuery()
                        u_UID = MyCommon.LRTsp.Parameters("@AdminUserID").Value
                        MyCommon.Write_Log(UMLogFile, "User: " & u_UID & " is created successfully.", True)
                        MyCommon.Close_LRTsp()
                        If (u_UID <> -1) Then
                            'Generate a Salt
                            Dim USalt As String = HashLib.GenerateNewSalt()
                            MyCommon.QueryStr = "update AdminUsers with (RowLock) set FirstName = N'" & FirstName & "',LastName = N'" & LastName & "',UserName =  N'" & Username & "',Password = N'" & HashLib.SQL_LoginHash(Password, USalt) & _
                                                 "',Email = N'" & MyCryptlib.SQL_StringEncrypt(Email) & "'" & UpdateString & ", LanguageID = 1 ,StartPageID = 1,StyleID = 4, USalt= N'" & USalt & "' where AdminUserID=" & u_UID
                            MyCommon.LRT_Execute()
                        End If
                        If RoleId <> "" Then
                            MyCommon.QueryStr = "Delete from AdminUserRoles with (RowLock) where AdminUserID=" & u_UID & ""
                            MyCommon.LRT_Execute()
                            Dim i As Integer
                            Dim aryTextFile() As String
                            aryTextFile = RoleId.Split(",")
                            Dim Flag As Boolean
                            For i = 0 To UBound(aryTextFile)
                                Flag = IsValidRoleId(aryTextFile(i))
                                If Flag Then
                                    MyCommon.QueryStr = "INSERT INTO AdminUserRoles (AdminUserID,RoleID) VALUES (" & u_UID & "," & aryTextFile(i) & ")"
                                    MyCommon.LRT_Execute()
                                Else
                                    RetCode = StatusCodes.INVALID_ROLEID
                                    RetMsg = "Failure:Invalid RoleID"
                                    Exit For
                                End If

                            Next i
                        End If
                    End If
                    If RetCode = StatusCodes.SUCCESS Then
                        RetMsg = "Success: User - " & Username & "(" & u_UID & ") added successfully."
                    End If
                End If
            End Using
        Catch ex As Exception
            RetCode = StatusCodes.APPLICATION_EXCEPTION
            RetMsg = "Failure.Application " & ex.ToString
        Finally
            If dtStatus IsNot Nothing AndAlso dtStatus.Rows.Count = 0 Then
                row = dtStatus.NewRow()
                row.Item("StatusCode") = RetCode
                row.Item("Description") = RetMsg
                dtStatus.Rows.Add(row)
            End If
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
        End Try
        Return ResultSet
    End Function

    'AMSPS-2157
    Private Function _AddUserWithBanner(ByVal GUID As String, ByVal Username As String, ByVal FirstName As String, ByVal LastName As String, ByVal Password As String, _
               ByVal Email As String, ByVal RoleId As String, ByVal EmployeeId As String, ByVal Banner As String, _
               ByVal MethodName As String) As DataSet

        Dim ResultSet As New System.Data.DataSet("UserManagement")
        Dim rst As DataTable
        Dim row As System.Data.DataRow
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = ""
        Dim MyCryptlib As New CMS.CryptLib
        Dim HashLib As New CMS.HashLib.CryptLib
        Dim UpdateString As String = ""
        Dim dtBanner As DataTable
        Dim u_UID As Integer
        Dim dtStatus As DataTable
        Dim BannerId As Integer = 0

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        If isValidInputText(GUID) = False Then
            'GUID Value is empty
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_GUID
            row.Item("Description") = "Failure.Invalid GUID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _AddUserWithBanner = ResultSet
            Exit Function
        End If
        If Username.Trim.Length = 0 Or Username.Trim.Length > 50 Then
            'Username Value is empty or invalid
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_USERNAME
            row.Item("Description") = "Failure.Invalid Username."
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _AddUserWithBanner = ResultSet
            Exit Function
        End If
        If FirstName.Trim.Length = 0 And LastName.Trim.Length = 0 Then
            'FirstName Value is empty or invalid
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_NAME
            row.Item("Description") = "Failure.Enter Firstname or Lastname."
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _AddUserWithBanner = ResultSet
            Exit Function
        End If
        If FirstName.Trim <> "" Then
            If FirstName.Length > 50 Then
                'FirstName Value is empty or invalid
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_FIRSTNAME
                row.Item("Description") = "Failure.Invalid FirstName."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                _AddUserWithBanner = ResultSet
                Exit Function
            End If
        End If
        If LastName.Trim <> "" Then
            If LastName.Length > 50 Then
                'FirstName Value is empty or invalid
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_LASTNAME
                row.Item("Description") = "Failure.Invalid LastName."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                _AddUserWithBanner = ResultSet
                Exit Function
            End If
        End If
        If Password.Trim.Length = 0 Or Password.Trim.Length > 20 Then
            'Password Value is empty or invalid
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_PASSWORD
            row.Item("Description") = "Failure.Invalid Password."
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _AddUserWithBanner = ResultSet
            Exit Function
        End If
        If Email <> "" Then
            If EmailAddressCheck(Email) = False Then
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_EMAILD
                row.Item("Description") = "Failure.Invalid Email."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                _AddUserWithBanner = ResultSet
                Exit Function
            End If
        End If
        If RoleId <> "" Then
            Dim lRoleID As Long = 0
            Try
                lRoleID = RoleId.Replace(",", "")
            Catch ex As Exception
                lRoleID = -1
            End Try
            If lRoleID = -1 Then
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_ROLEID
                row.Item("Description") = "Failure.Invalid RoleID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                _AddUserWithBanner = ResultSet
                Exit Function
            End If
        End If
        If EmployeeId <> "" Then
            Dim IsAlpha As Boolean
            If System.Text.RegularExpressions.Regex.IsMatch(EmployeeId, "^[a-zA-Z0-9]+$") Then
                IsAlpha = True
            Else
                IsAlpha = False
            End If
            If (IsAlpha) Then
                EmployeeId = MyCommon.CleanString(EmployeeId)
                MyCommon.QueryStr = "select EmployeeId from AdminUsers with (NoLock) where EmployeeId=N'" & EmployeeId & "'"
                rst = MyCommon.LRT_Select
                If rst.Rows.Count > 0 Then
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = StatusCodes.INVALID_EMPLOYEEID
                    row.Item("Description") = "Failure. EmployeeId already exists."
                    dtStatus.Rows.Add(row)
                    dtStatus.AcceptChanges()
                    ResultSet.Tables.Add(dtStatus.Copy())
                    _AddUserWithBanner = ResultSet
                    Exit Function
                End If
                UpdateString = ", EmployeeId = N'" & EmployeeId & "'"
            Else
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_EMPLOYEEID
                row.Item("Description") = "Failure. EmployeeId can have only alphanumeric characters"
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                _AddUserWithBanner = ResultSet
                Exit Function
            End If
        End If
        If Banner = "" Then
            'Checking if External Banner ID or Banner Name are blank
            row = dtStatus.NewRow()
            If MethodName.Contains("ExtBanner") Then
                row.Item("StatusCode") = StatusCodes.INVALID_EXTBANNERID
                row.Item("Description") = "Failure. Invalid ExtBannerId."
            Else
                row.Item("StatusCode") = StatusCodes.INVALID_BANNERNAME
                row.Item("Description") = "Failure. Invalid BannerName."
            End If
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _AddUserWithBanner = ResultSet
            Exit Function
        End If
        If Banner <> "" Then
            'If External Banner ID or Banner Name are not blank, checking whether provided banner exists in database. 
            If MethodName.Contains("ExtBanner") Then
                MyCommon.QueryStr = "select BannerID, Name from Banners with (NoLock) where ExtBannerID=N'" & Banner & "'"
            Else
                MyCommon.QueryStr = "select BannerID, Name from Banners with (NoLock) where Name=N'" & Banner & "'"
            End If
            rst = MyCommon.LRT_Select
            If rst.Rows.Count = 0 Then
                row = dtStatus.NewRow()
                If MethodName.Contains("ExtBanner") Then
                    row.Item("StatusCode") = StatusCodes.INVALID_EXTBANNERID
                    row.Item("Description") = "Failure. Invalid ExtBannerID"
                Else
                    row.Item("StatusCode") = StatusCodes.INVALID_BANNERNAME
                    row.Item("Description") = "Failure. Invalid BannerName"
                End If
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                _AddUserWithBanner = ResultSet
                Exit Function
            Else
                BannerId = Convert.ToInt32(rst.Rows(0).Item("BannerID"))
            End If
        End If
        Try
            Using MyCommon.LRTadoConn
                If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                    MyCommon.Open_LogixRT()
                End If

                If Not IsValidGUID(GUID, MethodName) Then
                    RetCode = StatusCodes.INVALID_GUID
                    RetMsg = "Failure.Invalid GUID"
                Else
                    MyCommon.QueryStr = "Select AdminUserID,UserName from AdminUsers with (NoLock) where UserName=@Username"
                    MyCommon.DBParameters.Add("@Username", SqlDbType.NVarChar).Value = Username
                    rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                    If rst.Rows.Count > 0 Then
                        u_UID = Convert.ToInt32(rst.Rows(0).Item("AdminUserID"))
                        If u_UID <> -1 AndAlso BannerId <> 0 Then
                            'Add banner to user if user already exists
                            MyCommon.QueryStr = "Select AdminUserID, BannerID from AdminUserBanners with (NoLock) where AdminUserID=" & u_UID & " and BannerID=" & BannerId
                            dtBanner = MyCommon.LRT_Select
                            If dtBanner.Rows.Count > 0 Then
                                If MethodName.Contains("ExtBanner") Then
                                    RetCode = StatusCodes.INVALID_EXTBANNERID
                                    RetMsg = "Failure:User- " & Username & " is already associated to the banner with ExtBannerID: " & Banner
                                Else
                                    RetCode = StatusCodes.INVALID_BANNERNAME
                                    RetMsg = "Failure:User- " & Username & " is already associated to the banner with BannerName: " & Banner
                                End If

                            Else
                                MyCommon.QueryStr = "INSERT INTO AdminUserBanners (AdminUserID,BannerID) VALUES (" & u_UID & "," & BannerId & ")"
                                MyCommon.LRT_Execute()
                                RetCode = StatusCodes.SUCCESS
                                If MethodName.Contains("ExtBanner") Then
                                    RetMsg = "Username already exists. Associated user to the banner with ExtBannerId: " & Banner & "."
                                Else
                                    RetMsg = "Username already exists. Associated user to the banner with BannerName: " & Banner & "."
                                End If
                            End If
                        End If
                    Else
                        'Creating user if not exists in database.
                        MyCommon.QueryStr = "dbo.pt_AdminUsers_Insert"
                        MyCommon.Open_LRTsp()
                        MyCommon.LRTsp.Parameters.Add("@UserName ", SqlDbType.NVarChar, 50).Value = Username
                        MyCommon.LRTsp.Parameters.Add("@AdminUserID", SqlDbType.Int).Direction = ParameterDirection.Output
                        MyCommon.LRTsp.ExecuteNonQuery()
                        u_UID = MyCommon.LRTsp.Parameters("@AdminUserID").Value
                        MyCommon.Write_Log(UMLogFile, "User: " & u_UID & " is created successfully.", True)
                        MyCommon.Close_LRTsp()
                        If (u_UID <> -1) Then
                            'Generate a Salt
                            Dim USalt As String = HashLib.GenerateNewSalt()
                            MyCommon.QueryStr = "update AdminUsers with (RowLock) set FirstName = N'" & FirstName & "',LastName = N'" & LastName & "',UserName =  N'" & Username & "',Password = N'" & HashLib.SQL_LoginHash(Password, USalt) & _
                                                 "',Email = N'" & MyCryptlib.SQL_StringEncrypt(Email) & "'" & UpdateString & ", LanguageID = 1 ,StartPageID = 1,StyleID = 4, USalt= N'" & USalt & "' where AdminUserID=" & u_UID
                            MyCommon.LRT_Execute()
                        End If
                        If RoleId <> "" Then
                            MyCommon.QueryStr = "Delete from AdminUserRoles with (RowLock) where AdminUserID=" & u_UID & ""
                            MyCommon.LRT_Execute()
                            Dim i As Integer
                            Dim aryTextFile() As String
                            aryTextFile = RoleId.Split(",")
                            Dim Flag As Boolean
                            For i = 0 To UBound(aryTextFile)
                                Flag = IsValidRoleId(aryTextFile(i))
                                If Flag Then
                                    MyCommon.QueryStr = "INSERT INTO AdminUserRoles (AdminUserID,RoleID) VALUES (" & u_UID & "," & aryTextFile(i) & ")"
                                    MyCommon.LRT_Execute()
                                Else
                                    RetCode = StatusCodes.INVALID_ROLEID
                                    RetMsg = "Failure:Invalid RoleID"
                                    Exit For
                                End If

                            Next i
                        End If

                        'Associate banner to the user created
                        If u_UID <> -1 AndAlso BannerId <> 0 Then
                            MyCommon.QueryStr = "Select AdminUserID, BannerID from AdminUserBanners with (NoLock) where AdminUserID=" & u_UID & " and BannerID=" & BannerId
                            dtBanner = MyCommon.LRT_Select
                            If dtBanner.Rows.Count > 0 Then
                                If MethodName.Contains("ExtBanner") Then
                                    RetCode = StatusCodes.INVALID_EXTBANNERID
                                    RetMsg = "Failure: User- " & Username & " is already associated to the banner with ExtBannerId: " & Banner
                                Else
                                    RetCode = StatusCodes.INVALID_BANNERNAME
                                    RetMsg = "Failure: User- " & Username & " is already associated to the banner with BannerName: " & Banner
                                End If
                            Else
                                MyCommon.QueryStr = "INSERT INTO AdminUserBanners (AdminUserID,BannerID) VALUES (" & u_UID & "," & BannerId & ")"
                                MyCommon.LRT_Execute()
                                RetCode = StatusCodes.SUCCESS
                                If MethodName.Contains("ExtBanner") Then
                                    RetMsg = "Success: User - " & Username & "(" & u_UID & ") added successfully with ExtBannerId: " & Banner & "."
                                Else
                                    RetMsg = "Success: User - " & Username & "(" & u_UID & ") added successfully with BannerName: " & Banner & "."
                                End If
                            End If
                        End If
                    End If
                End If
            End Using
        Catch ex As Exception
            RetCode = StatusCodes.APPLICATION_EXCEPTION
            RetMsg = "Failure.Application " & ex.ToString
        Finally
            If dtStatus IsNot Nothing AndAlso dtStatus.Rows.Count = 0 Then
                row = dtStatus.NewRow()
                row.Item("StatusCode") = RetCode
                row.Item("Description") = RetMsg
                dtStatus.Rows.Add(row)
            End If
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
        End Try
        Return ResultSet
    End Function

    Private Function _AddUserRoles(ByVal GUID As String, ByVal Username As String, ByVal RoleNames As String) As DataSet
        Dim ResultSet As New System.Data.DataSet("UserManagement")
        Dim MethodName As String = "AddUser"
        Dim dtStatus As DataTable
        Dim rst As DataTable
        Dim row As System.Data.DataRow
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = ""
        Dim u_UID As Integer
        Dim RoleId As Integer = 0
        Dim FailedRolecnt As Integer = 0
        Dim FirstRec As Boolean = False
        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        If isValidInputText(GUID) = False Then
            'GUID Value is empty
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_GUID
            row.Item("Description") = "Failure.Invalid GUID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _AddUserRoles = ResultSet
            Exit Function
        End If
        If Username.Trim.Length = 0 Or Username.Trim.Length > 50 Then
            'Username Value is empty or invalid
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_USERNAME
            row.Item("Description") = "Failure.Invalid Username."
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _AddUserRoles = ResultSet
            Exit Function
        End If
        If RoleNames.Trim.Length = 0 Then
            'Username Value is empty or invalid
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_ROLEID
            row.Item("Description") = "Failure.Invalid Role Names."
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _AddUserRoles = ResultSet
            Exit Function
        End If
        If RoleNames <> "" Then
            Dim lRoleNames As String = Nothing
            Try
                lRoleNames = RoleNames.Replace(",", "")
            Catch ex As Exception
                lRoleNames = Nothing
            End Try
            If lRoleNames = Nothing Then
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_ROLENAMES
                row.Item("Description") = "Failure.Invalid Role Names."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                _AddUserRoles = ResultSet
                Exit Function
            End If
        End If
        Try
            Using MyCommon.LRTadoConn
                If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                    MyCommon.Open_LogixRT()
                End If
                If Not IsValidGUID(GUID, MethodName) Then
                    RetCode = StatusCodes.INVALID_GUID
                    RetMsg = "Failure.Invalid GUID"
                Else
                    MyCommon.QueryStr = "Select AdminUserID,UserName from AdminUsers with (NoLock) where UserName=@Username"
                    MyCommon.DBParameters.Add("@Username", SqlDbType.NVarChar).Value = Username
                    rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                    If rst.Rows.Count > 0 Then
                        u_UID = MyCommon.NZ(rst.Rows(0).Item("AdminUserID"), 0)
                        If RoleNames <> "" Then
                            Dim i As Integer
                            Dim aryTextFile() As String
                            aryTextFile = RoleNames.Split(",")
                            For i = 0 To UBound(aryTextFile)
                                MyCommon.QueryStr = "SELECT RoleId FROM AdminRoles WHERE RoleName = N'" & aryTextFile(i).Trim & "'"
                                rst = MyCommon.LRT_Select
                                If rst.Rows.Count > 0 Then
                                    RoleId = MyCommon.NZ(rst.Rows(0).Item("RoleId"), 0)
                                    'If FirstRec = False Then
                                    '   MyCommon.QueryStr = "Delete from AdminUserRoles with (RowLock) where AdminUserID=" & u_UID & ""
                                    '  MyCommon.LRT_Execute()
                                    ' FirstRec = True
                                    'End If
                                    If Not UserRoleExists(u_UID, RoleId) Then
                                        MyCommon.QueryStr = "INSERT INTO AdminUserRoles (AdminUserID,RoleID) VALUES (" & u_UID & "," & RoleId & ")"
                                        MyCommon.LRT_Execute()
                                    End If
                                Else
                                    FailedRolecnt += 1
                                End If
                            Next i
                            If FailedRolecnt = aryTextFile.Length Then
                                RetCode = StatusCodes.INVALID_ROLENAMES
                                RetMsg = "Failure.Role name does not exist in the database."
                            End If
                        End If
                    Else
                        RetCode = StatusCodes.INVALID_USERNAME
                        RetMsg = "Failure.Username does not exist in the database."
                    End If

                    If RetCode = StatusCodes.SUCCESS Then
                        RetMsg = "Success: User Roles for - " & Username & " added successfully."
                    End If
                End If
            End Using
        Catch ex As Exception
            RetCode = StatusCodes.APPLICATION_EXCEPTION
            RetMsg = "Failure.Application " & ex.ToString
        Finally
            If dtStatus IsNot Nothing AndAlso dtStatus.Rows.Count = 0 Then
                row = dtStatus.NewRow()
                row.Item("StatusCode") = RetCode
                row.Item("Description") = RetMsg
                dtStatus.Rows.Add(row)
            End If
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
        End Try
        Return ResultSet
    End Function

    Private Function _ModifyUser(ByVal GUID As String, ByVal Username As String, ByVal FirstName As String, ByVal LastName As String, _
                                 ByVal Email As String, ByVal RoleId As String, ByVal EmployeeId As String) As DataSet

        Dim ResultSet As New System.Data.DataSet("UserManagement")
        Dim MethodName As String = "ModifyUser"
        Dim dtStatus As DataTable
        Dim rst As DataTable
        Dim row As System.Data.DataRow
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = ""
        Dim u_UID As Integer
        Dim MyCryptlib As New CMS.CryptLib
        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        If isValidInputText(GUID) = False Then
            'GUID Value is empty
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_GUID
            row.Item("Description") = "Failure.Invalid GUID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _ModifyUser = ResultSet
            Exit Function
        End If
        If Username.Trim = "0" Or Username.Length > 50 Then
            'Username Value is empty or invalid
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_USERNAME
            row.Item("Description") = "Failure.Invalid Username."
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _ModifyUser = ResultSet
            Exit Function
        End If
        If FirstName = "" And LastName = "" Then
            'FirstName Value is empty or invalid
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_NAME
            row.Item("Description") = "Failure."
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _ModifyUser = ResultSet
            Exit Function
        End If
        If FirstName <> "" Then
            If FirstName.Length > 50 Then
                'FirstName Value is empty or invalid
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_FIRSTNAME
                row.Item("Description") = "Failure.Invalid FirstName."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                _ModifyUser = ResultSet
                Exit Function
            End If
        End If
        If LastName <> "" Then
            If LastName.Length > 50 Then
                'FirstName Value is empty or invalid
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_LASTNAME
                row.Item("Description") = "Failure.Invalid LastName."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                _ModifyUser = ResultSet
                Exit Function
            End If
        End If
        If Email <> "" Then
            If EmailAddressCheck(Email) = False Then
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_EMAILD
                row.Item("Description") = "Failure.Invalid Email."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                _ModifyUser = ResultSet
                Exit Function
            End If
        End If
        If RoleId <> "" Then
            Dim lRoleID As Long = 0
            Try
                lRoleID = RoleId.Replace(",", "")
            Catch ex As Exception
                lRoleID = -1
            End Try
            If lRoleID = -1 Then
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_ROLEID
                row.Item("Description") = "Failure.Invalid RoleID."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                _ModifyUser = ResultSet
                Exit Function
            End If
        End If
        If EmployeeId <> "" Then
            Dim IsAlpha As Boolean
            If System.Text.RegularExpressions.Regex.IsMatch(EmployeeId, "^[a-zA-Z0-9]+$") Then
                IsAlpha = True
            Else
                IsAlpha = False
            End If
            If (IsAlpha) Then
                EmployeeId = MyCommon.CleanString(EmployeeId)
                MyCommon.QueryStr = "select EmployeeId from AdminUsers with (NoLock) where EmployeeId=N'" & EmployeeId & "'"
                rst = MyCommon.LRT_Select
                If rst.Rows.Count > 0 Then
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = StatusCodes.INVALID_EMPLOYEEID
                    row.Item("Description") = "Failure. EmployeeId already exists."
                    dtStatus.Rows.Add(row)
                    dtStatus.AcceptChanges()
                    ResultSet.Tables.Add(dtStatus.Copy())
                    _ModifyUser = ResultSet
                    Exit Function
                End If
            Else
                row = dtStatus.NewRow()
                row.Item("StatusCode") = StatusCodes.INVALID_EMPLOYEEID
                row.Item("Description") = "Failure. EmployeeId can have only alphanumeric characters"
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
                _ModifyUser = ResultSet
                Exit Function
            End If
        End If
        Try
            Using MyCommon.LRTadoConn
                If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                    MyCommon.Open_LogixRT()
                End If
                If Not IsValidGUID(GUID, MethodName) Then
                    RetCode = StatusCodes.INVALID_GUID
                    RetMsg = "Failure.Invalid GUID"
                Else
                    MyCommon.QueryStr = "Select AdminUserID,UserName from AdminUsers with (NoLock) where UserName=@Username"
                    MyCommon.DBParameters.Add("@Username", SqlDbType.NVarChar).Value = Username
                    rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                    If rst.Rows.Count > 0 Then
                        u_UID = MyCommon.NZ(rst.Rows(0).Item("AdminUserID"), 0)
                        If (u_UID <> -1) Then
                            MyCommon.QueryStr = "update AdminUsers with (RowLock) set FirstName = N'" & FirstName & "',LastName = N'" & LastName & "',UserName =  N'" & Username & "',Email = N'" & MyCryptlib.SQL_StringEncrypt(Email) & _
                                                 "', EmployeeId = N'" & EmployeeId & "', LanguageID = 1 ,StartPageID = 1,StyleID = 4 where AdminUserID=" & u_UID
                            MyCommon.LRT_Execute()
                        End If
                        If RoleId <> "" Then
                            MyCommon.QueryStr = "Delete from AdminUserRoles with (RowLock) where AdminUserID=" & u_UID & ""
                            MyCommon.LRT_Execute()
                            Dim i As Integer
                            Dim aryTextFile() As String
                            aryTextFile = RoleId.Split(",")
                            Dim flag As Boolean
                            For i = 0 To UBound(aryTextFile)
                                flag = IsValidRoleId(aryTextFile(i))
                                If flag Then
                                    MyCommon.QueryStr = "INSERT INTO AdminUserRoles (AdminUserID,RoleID) VALUES (" & u_UID & "," & aryTextFile(i) & ")"
                                    MyCommon.LRT_Execute()
                                Else
                                    RetCode = StatusCodes.INVALID_ROLEID
                                    RetMsg = "Failure: RoleID does not Exist"
                                    Exit For
                                End If

                            Next i
                        End If
                    Else
                        RetCode = StatusCodes.INVALID_USERID
                        RetMsg = "Failure.Username does not exist in database."
                    End If
                    If RetCode = StatusCodes.SUCCESS Then
                        RetMsg = "Success: User - " & Username & "(" & u_UID & ") modified successfully."
                    End If
                End If
            End Using
        Catch ex As Exception
            RetCode = StatusCodes.APPLICATION_EXCEPTION
            RetMsg = "Failure.Application " & ex.ToString
        Finally
            If dtStatus IsNot Nothing AndAlso dtStatus.Rows.Count = 0 Then
                row = dtStatus.NewRow()
                row.Item("StatusCode") = RetCode
                row.Item("Description") = RetMsg
                dtStatus.Rows.Add(row)
            End If
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
        End Try
        Return ResultSet
    End Function

    Private Function _DeleteUserById(ByVal GUID As String, ByVal UserId As Integer) As DataSet

        Dim ResultSet As New System.Data.DataSet("UserManagement")
        Dim MethodName As String = "DeleteUserById"
        Dim dtStatus As DataTable
        Dim rst As DataTable
        Dim row As System.Data.DataRow
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = ""
        Dim Logix As New Copient.LogixInc

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        If isValidInputText(GUID) = False Then
            'GUID Value is empty
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_GUID
            row.Item("Description") = "Failure.Invalid GUID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _DeleteUserById = ResultSet
            Exit Function
        End If
        If UserId = -1 Then
            'Username Value is empty or invalid
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_USERID
            row.Item("Description") = "Failure.Invalid UserId."
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _DeleteUserById = ResultSet
            Exit Function
        End If
        Try
            Using MyCommon.LRTadoConn
                If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                    MyCommon.Open_LogixRT()
                End If
                If Not IsValidGUID(GUID, MethodName) Then
                    RetCode = StatusCodes.INVALID_GUID
                    RetMsg = "Failure.Invalid GUID"
                Else
                    MyCommon.QueryStr = "Select UserName from AdminUsers with (NoLock) where AdminUserID = " & UserId
                    rst = MyCommon.LRT_Select
                    If rst.Rows.Count > 0 Then
                        If Logix.IsAdministrator(UserId, MyCommon) Then
                            RetCode = StatusCodes.INVALID_USERID
                            RetMsg = "Failure.Admin User cannot be deleted."
                        Else
                            MyCommon.QueryStr = "DELETE FROM AdminUsers with (RowLock) WHERE AdminUserID= " & UserId
                            MyCommon.LRT_Execute()
                            MyCommon.QueryStr = "DELETE from AdminUserRoles with (RowLock) where AdminUserID=" & UserId
                            MyCommon.LRT_Execute()
                        End If
                    Else
                        RetCode = StatusCodes.INVALID_USERID
                        RetMsg = "Failure.Userid does not exist in database."
                    End If
                    If RetCode = StatusCodes.SUCCESS Then
                        RetMsg = "Success: User - " & MyCommon.NZ(rst.Rows(0).Item("UserName"), 0) & "(" & UserId & ") deleted successfully."
                    End If
                End If
            End Using
        Catch ex As Exception
            RetCode = StatusCodes.APPLICATION_EXCEPTION
            RetMsg = "Failure.Application " & ex.ToString
        Finally
            If dtStatus IsNot Nothing AndAlso dtStatus.Rows.Count = 0 Then
                row = dtStatus.NewRow()
                row.Item("StatusCode") = RetCode
                row.Item("Description") = RetMsg
                dtStatus.Rows.Add(row)
            End If
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
        End Try
        Return ResultSet
    End Function

    Private Function _DeleteUserByName(ByVal GUID As String, ByVal Username As String) As DataSet

        Dim ResultSet As New System.Data.DataSet("UserManagement")
        Dim MethodName As String = "DeleteUserByName"
        Dim dtStatus As DataTable
        Dim rst As DataTable
        Dim row As System.Data.DataRow
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = ""
        Dim Logix As New Copient.LogixInc
        Dim UserId As Integer

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        If isValidInputText(GUID) = False Then
            'GUID Value is empty
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_GUID
            row.Item("Description") = "Failure.Invalid GUID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _DeleteUserByName = ResultSet
            Exit Function
        End If
        If Username.Trim = "0" Then
            'Email Value is empty or invalid
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_USERNAME
            row.Item("Description") = "Failure.Invalid Username."
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _DeleteUserByName = ResultSet
            Exit Function
        End If
        Try
            Using MyCommon.LRTadoConn
                If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                    MyCommon.Open_LogixRT()
                End If
                If Not IsValidGUID(GUID, MethodName) Then
                    RetCode = StatusCodes.INVALID_GUID
                    RetMsg = "Failure.Invalid GUID"
                Else
                    MyCommon.QueryStr = "Select AdminUserID from AdminUsers with (NoLock) where UserName=@Username"
                    MyCommon.DBParameters.Add("@Username", SqlDbType.NVarChar).Value = Username
                    rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                    If rst.Rows.Count > 0 Then
                        UserId = MyCommon.NZ(rst.Rows(0).Item("AdminUserID"), 0)
                        If Logix.IsAdministrator(UserId, MyCommon) Then
                            RetCode = StatusCodes.INVALID_USERID
                            RetMsg = "Failure.Admin User cannot be deleted."
                        Else
                            MyCommon.QueryStr = "DELETE FROM AdminUsers with (RowLock) where AdminUserID = " & UserId
                            MyCommon.LRT_Execute()
                            MyCommon.QueryStr = "DELETE from AdminUserRoles with (RowLock) where AdminUserID = " & UserId
                            MyCommon.LRT_Execute()
                        End If
                    Else
                        RetCode = StatusCodes.INVALID_USERID
                        RetMsg = "Failure.Username does not exist in database."
                    End If
                    If RetCode = StatusCodes.SUCCESS Then
                        RetMsg = "Success: User - " & Username & "(" & UserId & ") deleted successfully."
                    End If
                End If
            End Using
        Catch ex As Exception
            RetCode = StatusCodes.APPLICATION_EXCEPTION
            RetMsg = "Failure.Application " & ex.ToString
        Finally
            If dtStatus IsNot Nothing AndAlso dtStatus.Rows.Count = 0 Then
                row = dtStatus.NewRow()
                row.Item("StatusCode") = RetCode
                row.Item("Description") = RetMsg
                dtStatus.Rows.Add(row)
            End If
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
        End Try
        Return ResultSet
    End Function

    Private Function _ChangePassword(ByVal GUID As String, ByVal Username As String, ByVal OldPassword As String, _
                           ByVal NewPassword As String) As DataSet

        Dim ResultSet As New System.Data.DataSet("UserManagement")
        Dim MethodName As String = "ChangePassword"
        Dim dtStatus As DataTable
        Dim rst As DataTable
        Dim row As System.Data.DataRow
        Dim ErrorMsg As String = ""
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = ""
        Dim MyCryptlib As New CMS.CryptLib
        Dim HashLib As New CMS.HashLib.CryptLib

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        If isValidInputText(GUID) = False Then
            'GUID Value is empty
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_GUID
            row.Item("Description") = "Failure.Invalid GUID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _ChangePassword = ResultSet
            Exit Function
        End If
        If Username.Trim = "0" Then
            'Username Value is empty or invalid
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_USERNAME
            row.Item("Description") = "Failure.Invalid Username."
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _ChangePassword = ResultSet
            Exit Function
        End If
        If OldPassword.Trim.Length = 0 Then
            'Password Value is empty or invalid
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_PASSWORD
            row.Item("Description") = "Failure.Invalid Old Password."
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _ChangePassword = ResultSet
            Exit Function
        End If
        If NewPassword.Trim.Length = 0 Then
            'PasswordCon Value is empty or invalid
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_PASSWORD
            row.Item("Description") = "Failure.Invalid New Password."
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _ChangePassword = ResultSet
            Exit Function
        End If
        Try
            Using MyCommon.LRTadoConn
                If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                    MyCommon.Open_LogixRT()
                End If
                If Not IsValidGUID(GUID, MethodName) Then
                    RetCode = StatusCodes.INVALID_GUID
                    RetMsg = "Failure.Invalid GUID"
                Else
                    MyCommon.QueryStr = "Select Password from AdminUsers with (NoLock) where UserName=@Username"
                    MyCommon.DBParameters.Add("@Username", SqlDbType.NVarChar).Value = Username
                    rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                    If rst.Rows.Count = 0 Then
                        RetCode = StatusCodes.INVALID_USERNAME
                        RetMsg = "Failure.Username does not exists."
                    Else
                        Dim old_USalt As String = MyCommon.GetUserSalt(MyCommon, Username)
                        Dim EncryptedOldpassword As String = String.Empty
                        Dim EncryptedNewpassword As String = String.Empty

                        'Here checking the Usalt for the user, if it is blank or null then convert both new and old password accordingly for validation 
                        If Not String.IsNullOrEmpty(old_USalt) Then
                            EncryptedOldpassword = HashLib.SQL_LoginHash(OldPassword, old_USalt)
                            EncryptedNewpassword = HashLib.SQL_LoginHash(NewPassword, old_USalt)
                        Else
                            EncryptedOldpassword = MyCryptlib.SQL_LegacyLoginEncrypt(OldPassword)
                            EncryptedNewpassword = MyCryptlib.SQL_LegacyLoginEncrypt(NewPassword)
                        End If
                        If rst.Rows(0)(0) = EncryptedOldpassword Then
                            'Validate current password and new passowrd should not be same
                            If rst.Rows(0)(0) = EncryptedNewpassword Then
                                RetCode = StatusCodes.INVALID_PASSWORDREUSED
                                RetMsg = "Failure.Password should not be same as current. Please change the password.."
                            Else
                                CurrentRequest.Resolver.AppName = "UserManagement.vb"
                                Dim m_adminUserDataService As IAdminUserData = CurrentRequest.Resolver.Resolve(Of IAdminUserData)()
                                If m_adminUserDataService.ValidatePassword(NewPassword, Username).ResultType = CMS.AMS.Models.AMSResultType.Success Then
                                    'Generate a Salt
                                    Dim USalt As String = HashLib.GenerateNewSalt()
                                    MyCommon.QueryStr = "UPDATE AdminUsers with (RowLock) SET Password=@Password, PasswordChangedDate=@PasswordChangedDate, USalt=@USalt " & _
                                                        "WHERE UserName = @UserName"
                                    MyCommon.DBParameters.Add("@UserName", SqlDbType.NVarChar, 50).Value = Username
                                    MyCommon.DBParameters.Add("@Password", SqlDbType.NVarChar, 1000).Value = HashLib.SQL_LoginHash(NewPassword, USalt)
                                    MyCommon.DBParameters.Add("@PasswordChangedDate", SqlDbType.DateTime).Value = DateTime.Now
                                    MyCommon.DBParameters.Add("@USalt", System.Data.SqlDbType.NChar).Value = USalt
                                    MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
                                    RetCode = StatusCodes.SUCCESS
                                    RetMsg = "Success"
                                Else
                                    RetCode = StatusCodes.INVALID_PASSWORD
                                    RetMsg = "Failure.Invalid New Password. Password should contain at least one uppercase, lowercase character, digit and special character."
                                End If
                            End If
                        Else
                            RetCode = StatusCodes.INVALID_PASSWORD
                            RetMsg = "Failure.Invalid Old Password."
                        End If
                        If RetCode = StatusCodes.SUCCESS Then
                            RetMsg = "Success: Password changed successfully for user - " & Username
                        End If
                    End If
                End If
            End Using
        Catch ex As Exception
            RetCode = StatusCodes.APPLICATION_EXCEPTION
            RetMsg = "Failure.Application " & ex.ToString
        Finally
            If dtStatus IsNot Nothing AndAlso dtStatus.Rows.Count = 0 Then
                row = dtStatus.NewRow()
                row.Item("StatusCode") = RetCode
                row.Item("Description") = RetMsg
                dtStatus.Rows.Add(row)
            End If
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
        End Try
        Return ResultSet

    End Function

    Private Function _DeleteUserByAuthentication(ByVal GUID As String, ByVal AdminUserName As String, _
                                                      ByVal Password As String, ByVal DeleteUserName As String) As DataSet
        Dim ResultSet As New System.Data.DataSet("UserManagement")
        Dim MethodName As String = "DeleteUserByAuthentication"
        Dim dtStatus As DataTable
        Dim rstAuthenticate, rst As DataTable
        Dim row As System.Data.DataRow
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = ""
        Dim Logix As New Copient.LogixInc
        Dim MyCryptlib As New CMS.CryptLib
        Dim HashLib As New CMS.HashLib.CryptLib
        Dim UserId As Integer

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        If isValidInputText(GUID) = False Then
            'GUID Value is empty
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_GUID
            row.Item("Description") = "Failure.Invalid GUID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _DeleteUserByAuthentication = ResultSet
            Exit Function
        End If
        If DeleteUserName.Trim = "0" Then
            'Username Value is empty or invalid
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_USERNAME
            row.Item("Description") = "Failure.Invalid Username."
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _DeleteUserByAuthentication = ResultSet
            Exit Function
        End If
        If AdminUserName.Trim.Length = 0 Then
            'Password Value is empty or invalid
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_USERNAME
            row.Item("Description") = "Failure.Invalid Authenticate user name."
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _DeleteUserByAuthentication = ResultSet
            Exit Function
        End If
        If Password.Trim.Length = 0 Then
            'Password Value is empty or invalid
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_PASSWORD
            row.Item("Description") = "Failure.Invalid Password."
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _DeleteUserByAuthentication = ResultSet
            Exit Function
        End If
        Try
            Using MyCommon.LRTadoConn
                If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                    MyCommon.Open_LogixRT()
                End If
                If Not IsValidGUID(GUID, MethodName) Then
                    RetCode = StatusCodes.INVALID_GUID
                    RetMsg = "Failure.Invalid GUID"
                Else

                    Dim USalt As String = MyCommon.GetUserSalt(MyCommon, AdminUserName)
                    Dim strUserPassword As String = String.Empty
                    If Not String.IsNullOrEmpty(USalt) Then
                        strUserPassword = HashLib.SQL_LoginHash(Password, USalt)
                    Else
                        strUserPassword = MyCryptlib.SQL_LegacyLoginEncrypt(Password)
                    End If
                    
                    MyCommon.QueryStr = "Select AdminUserID from AdminUsers with (NoLock) where UserName=@AdminUserName and Password=@strUserPassword"
                    MyCommon.DBParameters.Add("@AdminUserName", SqlDbType.NVarChar).Value = AdminUserName
                    MyCommon.DBParameters.Add("@strUserPassword", SqlDbType.NVarChar).Value = strUserPassword
                    rstAuthenticate = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                    If rstAuthenticate.Rows.Count > 0 Then
                        MyCommon.QueryStr = "Select AdminUserID from AdminUsers with (NoLock) where UserName=@DeleteUserName"
                        MyCommon.DBParameters.Add("@DeleteUserName", SqlDbType.NVarChar).Value = DeleteUserName
                        rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                        If rst.Rows.Count > 0 Then
                            UserId = rst.Rows(0)(0)
                            If Logix.IsAdministrator(UserId, MyCommon) Then 'Logix.i(UserId, MyCommon)
                                RetCode = StatusCodes.INVALID_USERID
                                RetMsg = "Failure.Admin User cannot be deleted."
                            Else
                                MyCommon.QueryStr = "DELETE FROM AdminUsers with (RowLock) WHERE AdminUserID= " & UserId
                                MyCommon.LRT_Execute()
                                MyCommon.QueryStr = "DELETE from AdminUserRoles with (RowLock) where AdminUserID=" & UserId
                                MyCommon.LRT_Execute()
                                RetCode = StatusCodes.SUCCESS
                                RetMsg = "Success: User - " & DeleteUserName & "(" & UserId & ") deleted successfully."
                            End If
                        Else
                            RetCode = StatusCodes.INVALID_USERID
                            RetMsg = "Failure.Userid does not exist in database."
                        End If
                    Else
                        RetCode = StatusCodes.INVALID_AUTHENTICATION
                        RetMsg = "Failure.Invalid username/password."
                    End If
                End If
            End Using
        Catch ex As Exception
            RetCode = StatusCodes.APPLICATION_EXCEPTION
            RetMsg = "Failure.Application " & ex.ToString
        Finally
            If dtStatus IsNot Nothing AndAlso dtStatus.Rows.Count = 0 Then
                row = dtStatus.NewRow()
                row.Item("StatusCode") = RetCode
                row.Item("Description") = RetMsg
                dtStatus.Rows.Add(row)
            End If
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
        End Try
        Return ResultSet
    End Function

    Private Function _GetUserDetailsByID(ByVal GUID As String, ByVal UserId As Integer) As DataSet

        Dim ResultSet As New System.Data.DataSet("UserManagement")
        Dim MethodName As String = "GetUserDetails"
        Dim dtStatus As DataTable
        Dim row As System.Data.DataRow
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = ""
        Dim dtUserDetails As System.Data.DataTable = New DataTable()
        Dim MyCryptlib As New CMS.CryptLib

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        If isValidInputText(GUID) = False Then
            'GUID Value is empty
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_GUID
            row.Item("Description") = "Failure.Invalid GUID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _GetUserDetailsByID = ResultSet
            Exit Function
        End If
        If UserId = -1 Then
            'UserId  Value is empty or invalid
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_USERID
            row.Item("Description") = "Failure.Invalid UserId."
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _GetUserDetailsByID = ResultSet
            Exit Function
        End If
        Try
            Using MyCommon.LRTadoConn
                If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                    MyCommon.Open_LogixRT()
                End If
                If Not IsValidGUID(GUID, MethodName) Then
                    RetCode = StatusCodes.INVALID_GUID
                    RetMsg = "Failure.Invalid GUID"
                Else
                    MyCommon.QueryStr = "SELECT  AdminUserID, FirstName, LastName, UserName, Email, LastLogin, JobTitle, EmployeeId FROM AdminUsers WHERE AdminUserID = " & UserId
                    dtUserDetails = MyCommon.LRT_Select
                    If dtUserDetails.Rows.Count > 0 Then
                        For Each dtrow As DataRow In dtUserDetails.Rows
                            dtrow("Email") = MyCryptlib.SQL_StringDecrypt(dtrow("Email").ToString())
                        Next
                        RetMsg = "Success"
                    Else
                        RetCode = StatusCodes.INVALID_USERID
                        RetMsg = " User Id does not exist."
                    End If
                End If
            End Using
        Catch ex As Exception
            RetCode = StatusCodes.APPLICATION_EXCEPTION
            RetMsg = "Failure.Application " & ex.ToString
        Finally
            If dtStatus IsNot Nothing AndAlso dtStatus.Rows.Count = 0 Then
                row = dtStatus.NewRow()
                row.Item("StatusCode") = RetCode
                row.Item("Description") = RetMsg
                dtStatus.Rows.Add(row)
            End If
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            If RetCode = StatusCodes.SUCCESS Then
                dtUserDetails.TableName = "UserDetailsById"
                dtUserDetails.AcceptChanges()
                ResultSet.Tables.Add(dtUserDetails.Copy())
            End If
        End Try
        Return ResultSet
    End Function

    Private Function _GetUserDetailsByName(ByVal GUID As String, ByVal UserName As String) As DataSet

        Dim ResultSet As New System.Data.DataSet("UserManagement")
        Dim MethodName As String = "GetUserDetailsByName"
        Dim dtStatus As DataTable
        Dim row As System.Data.DataRow
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = ""
        Dim dtUserDetails As System.Data.DataTable = Nothing
        Dim MyCryptlib As New CMS.CryptLib

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        If isValidInputText(GUID) = False Then
            'GUID Value is empty
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_GUID
            row.Item("Description") = "Failure.Invalid GUID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _GetUserDetailsByName = ResultSet
            Exit Function
        End If
        If UserName.Trim.Length = 0 Or UserName.Length > 50 Then
            'UserName  Value is empty or invalid
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_USERNAME
            row.Item("Description") = "Failure.Invalid UserName."
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _GetUserDetailsByName = ResultSet
            Exit Function
        End If
        Try
            Using MyCommon.LRTadoConn
                If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                    MyCommon.Open_LogixRT()
                End If
                If Not IsValidGUID(GUID, MethodName) Then
                    RetCode = StatusCodes.INVALID_GUID
                    RetMsg = "Failure.Invalid GUID"
                Else
                    MyCommon.QueryStr = "SELECT  AdminUserID, FirstName, LastName, UserName, Email, LastLogin, JobTitle, EmployeeId FROM AdminUsers WHERE UserName=@UserName"
                    MyCommon.DBParameters.Add("@UserName", SqlDbType.NVarChar).Value = UserName
                    dtUserDetails = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                    If dtUserDetails.Rows.Count > 0 Then
                        For Each dtrow As DataRow In dtUserDetails.Rows
                            dtrow("Email") = MyCryptlib.SQL_StringDecrypt(dtrow("Email").ToString())
                        Next
                        RetMsg = "Success"
                    Else
                        RetCode = StatusCodes.INVALID_USERID
                        RetMsg = " Username does not exist."
                    End If
                End If
            End Using
        Catch ex As Exception
            RetCode = StatusCodes.APPLICATION_EXCEPTION
            RetMsg = "Failure.Application " & ex.ToString
        Finally
            If dtStatus IsNot Nothing AndAlso dtStatus.Rows.Count = 0 Then
                row = dtStatus.NewRow()
                row.Item("StatusCode") = RetCode
                row.Item("Description") = RetMsg
                dtStatus.Rows.Add(row)
            End If
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            If RetCode = StatusCodes.SUCCESS Then
                dtUserDetails.TableName = "UserDetails"
                dtUserDetails.AcceptChanges()
                ResultSet.Tables.Add(dtUserDetails.Copy())
            End If
        End Try
        Return ResultSet
    End Function

    Private Function _GetUserRolesByName(ByVal GUID As String, ByVal Username As String) As System.Data.DataSet

        Dim ResultSet As New System.Data.DataSet("UserManagement")
        Dim MethodName As String = "GetUserRolesByName"
        Dim dtStatus As DataTable
        Dim dtUserRoles As DataTable = New DataTable()
        Dim rst As DataTable
        Dim row As System.Data.DataRow
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = ""
        Dim bOpenedRTConnection As Boolean = False
        Dim UserId As Integer = 0

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        If isValidInputText(GUID) = False Then
            'GUID Value is empty
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_GUID
            row.Item("Description") = "Failure.Invalid GUID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _GetUserRolesByName = ResultSet
            Exit Function
        End If
        If Username.Trim.Length = 0 Then
            'Username Value is empty or invalid
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_USERNAME
            row.Item("Description") = "Failure.Invalid User Name."
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _GetUserRolesByName = ResultSet
            Exit Function
        End If
        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                bOpenedRTConnection = True
                MyCommon.Open_LogixRT()
            End If
            If Not IsValidGUID(GUID, MethodName) Then
                RetCode = StatusCodes.INVALID_GUID
                RetMsg = "Failure.Invalid GUID"
            Else
                MyCommon.QueryStr = "SELECT AdminUserID FROM AdminUsers with (NoLock) WHERE UserName=@UserName"
                MyCommon.DBParameters.Add("@UserName", SqlDbType.NVarChar).Value = Username
                rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                If rst.Rows.Count = 0 Then
                    RetCode = StatusCodes.INVALID_USERID
                    RetMsg = "Failure: Invalid UserID"
                Else
                    UserId = rst.Rows(0)(0)
                    dtUserRoles = _GetUserRoleDetails(UserId)
                    If dtUserRoles.Rows.Count <= 0 Then
                        RetCode = StatusCodes.NOTFOUND_RECORDS
                        RetMsg = "Failure: NOT Found Records"
                    Else
                        RetMsg = "Success"
                    End If
                End If

            End If
        Catch ex As Exception
            RetCode = StatusCodes.APPLICATION_EXCEPTION
            RetMsg = "Failure.Application " & ex.ToString
        Finally
            If dtStatus IsNot Nothing AndAlso dtStatus.Rows.Count = 0 Then
                row = dtStatus.NewRow()
                row.Item("StatusCode") = RetCode
                row.Item("Description") = RetMsg
                dtStatus.Rows.Add(row)
            End If
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            If RetCode = StatusCodes.SUCCESS Then
                dtUserRoles.TableName = "UserRoles"
                dtUserRoles.AcceptChanges()
                ResultSet.Tables.Add(dtUserRoles.Copy())
            End If
        End Try
        Return ResultSet
    End Function

    Private Function _GetUserRolesByID(ByVal GUID As String, ByVal UserId As Integer) As System.Data.DataSet

        Dim ResultSet As New System.Data.DataSet("UserManagement")
        Dim MethodName As String = "GetUserRolesByID"
        Dim dtStatus As DataTable
        Dim dtUserRoles As DataTable = New DataTable()
        Dim rst As DataTable
        Dim row As System.Data.DataRow
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = ""
        Dim bOpenedRTConnection As Boolean = False

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        If isValidInputText(GUID) = False Then
            'GUID Value is empty
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_GUID
            row.Item("Description") = "Failure.Invalid GUID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _GetUserRolesByID = ResultSet
            Exit Function
        End If
        If UserId <= 0 Then
            'Username Value is empty or invalid
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_USERID
            row.Item("Description") = "Failure.Invalid UserID."
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _GetUserRolesByID = ResultSet
            Exit Function
        End If
        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                bOpenedRTConnection = True
                MyCommon.Open_LogixRT()
            End If
            If Not IsValidGUID(GUID, MethodName) Then
                RetCode = StatusCodes.INVALID_GUID
                RetMsg = "Failure.Invalid GUID"
            Else
                MyCommon.QueryStr = "Select AdminUserID from AdminUsers with (NoLock) where AdminUserID=" & UserId.ToString
                rst = MyCommon.LRT_Select
                If rst.Rows.Count = 0 Then
                    RetCode = StatusCodes.INVALID_USERID
                    RetMsg = "Failure: Invalid UserID"
                Else
                    dtUserRoles = _GetUserRoleDetails(UserId)
                    If dtUserRoles.Rows.Count <= 0 Then
                        RetCode = StatusCodes.NOTFOUND_RECORDS
                        RetMsg = "Failure: NOT Found Records"
                    Else
                        RetMsg = "Success"
                    End If
                End If
            End If
        Catch ex As Exception
            RetCode = StatusCodes.APPLICATION_EXCEPTION
            RetMsg = "Failure.Application " & ex.ToString
        Finally
            If dtStatus IsNot Nothing AndAlso dtStatus.Rows.Count = 0 Then
                row = dtStatus.NewRow()
                row.Item("StatusCode") = RetCode
                row.Item("Description") = RetMsg
                dtStatus.Rows.Add(row)
            End If
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            If RetCode = StatusCodes.SUCCESS Then
                dtUserRoles.TableName = "UserRoles"
                dtUserRoles.AcceptChanges()
                ResultSet.Tables.Add(dtUserRoles.Copy())
            End If
        End Try
        Return ResultSet
    End Function

    Private Function _GetUserRoleDetails(ByVal UserId As String) As DataTable

        Dim dt As DataTable
        MyCommon.QueryStr = "select R.RoleID, R.RoleName from AdminRoles as R with (NoLock) where R.RoleID " & _
                            "in(select RoleID from AdminUserRoles where AdminUserID=" & UserId.ToString & ") ORDER BY RoleName"
        dt = MyCommon.LRT_Select
        Return dt
    End Function

    Private Function _GetUsersRoles(ByVal GUID As String) As DataSet

        Dim ResultSet As New System.Data.DataSet("UserManagement")
        Dim MethodName As String = "GetUserList"
        Dim dtStatus As DataTable
        Dim row As System.Data.DataRow
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = ""
        Dim dtUserRoles As System.Data.DataTable = Nothing

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        If isValidInputText(GUID) = False Then
            'GUID Value is empty
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_GUID
            row.Item("Description") = "Failure.Invalid GUID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _GetUsersRoles = ResultSet
            Exit Function
        End If

        Try
            Using MyCommon.LRTadoConn
                If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                    MyCommon.Open_LogixRT()
                End If
                If Not IsValidGUID(GUID, MethodName) Then
                    RetCode = StatusCodes.INVALID_GUID
                    RetMsg = "Failure.Invalid GUID"
                Else
                    MyCommon.QueryStr = "SELECT UserName, AUR.RoleID, RoleName FROM AdminRoles AR WITH (NoLock)" & _
                                            "INNER JOIN AdminUserRoles AUR WITH (NoLock) ON AUR.RoleID = AR.RoleID " & _
                                            "INNER JOIN AdminUsers AU with (NoLock) ON AU.AdminUserID = AUR.AdminUserID"
                    dtUserRoles = MyCommon.LRT_Select
                    If dtUserRoles.Rows.Count > 0 Then
                        RetMsg = "Success"
                    Else
                        RetCode = StatusCodes.INVALID_USERID
                        RetMsg = "Failure.No roles/user exists in the system."
                    End If
                End If
            End Using
        Catch ex As Exception
            RetCode = StatusCodes.APPLICATION_EXCEPTION
            RetMsg = "Failure.Application " & ex.ToString
        Finally
            If dtStatus IsNot Nothing AndAlso dtStatus.Rows.Count = 0 Then
                row = dtStatus.NewRow()
                row.Item("StatusCode") = RetCode
                row.Item("Description") = RetMsg
                dtStatus.Rows.Add(row)
            End If
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            If RetCode = StatusCodes.SUCCESS Then
                dtUserRoles.TableName = "UsersRoles"
                dtUserRoles.AcceptChanges()
                ResultSet.Tables.Add(dtUserRoles.Copy())
            End If

        End Try

        Return ResultSet

    End Function

    Private Function _GetRolePermissionsList(ByVal GUID As String, ByVal RoleId As Integer) As DataSet

        Dim ResultSet As New System.Data.DataSet("UserManagement")
        Dim MethodName As String = "GetUserList"
        Dim dtStatus As DataTable
        Dim row As System.Data.DataRow
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = ""
        Dim dtRoleDetails As System.Data.DataTable = Nothing

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        If isValidInputText(GUID) = False Then
            'GUID Value is empty
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_GUID
            row.Item("Description") = "Failure.Invalid GUID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _GetRolePermissionsList = ResultSet
            Exit Function
        End If

        Try
            Using MyCommon.LRTadoConn
                If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                    MyCommon.Open_LogixRT()
                End If
                If Not IsValidGUID(GUID, MethodName) Then
                    RetCode = StatusCodes.INVALID_GUID
                    RetMsg = "Failure.Invalid GUID"
                Else
                    If RoleId <> -1 Then
                        MyCommon.QueryStr = "SELECT RP.PermissionID, P.Description from RolePermissions RP with (NoLock) " & _
                                            "INNER JOIN Permissions P with (NoLock) ON P.PermissionID = RP.PermissionID where RoleID = " & RoleId
                    Else
                        MyCommon.QueryStr = "SELECT RP.PermissionID, P.Description from RolePermissions RP with (NoLock) " & _
                                            "INNER JOIN Permissions P with (NoLock) ON P.PermissionID = RP.PermissionID "
                    End If
                    dtRoleDetails = MyCommon.LRT_Select
                    If dtRoleDetails.Rows.Count > 0 Then
                        RetMsg = "Success"
                    Else
                        RetCode = StatusCodes.INVALID_USERID
                        RetMsg = "Failure.No Permission exist for this role."
                    End If
                End If
            End Using
        Catch ex As Exception
            RetCode = StatusCodes.APPLICATION_EXCEPTION
            RetMsg = "Failure.Application " & ex.ToString
        Finally
            If dtStatus IsNot Nothing AndAlso dtStatus.Rows.Count = 0 Then
                row = dtStatus.NewRow()
                row.Item("StatusCode") = RetCode
                row.Item("Description") = RetMsg
                dtStatus.Rows.Add(row)
            End If
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            If RetCode = StatusCodes.SUCCESS Then
                dtRoleDetails.TableName = "RolePermissions"
                dtRoleDetails.AcceptChanges()
                ResultSet.Tables.Add(dtRoleDetails.Copy())
            End If

        End Try

        Return ResultSet

    End Function

    Private Function _GetUserList(ByVal GUID As String) As DataSet

        Dim ResultSet As New System.Data.DataSet("UserManagement")
        Dim MethodName As String = "GetUserList"
        Dim dtStatus As DataTable
        Dim row As System.Data.DataRow
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = ""
        Dim dtUserDetails As System.Data.DataTable = Nothing
        Dim MyCryptlib As New CMS.CryptLib

        'Initialize the status table, which will report the success or failure of the operation
        dtStatus = New DataTable
        dtStatus.TableName = "Status"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        If isValidInputText(GUID) = False Then
            'GUID Value is empty
            row = dtStatus.NewRow()
            row.Item("StatusCode") = StatusCodes.INVALID_GUID
            row.Item("Description") = "Failure.Invalid GUID"
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            _GetUserList = ResultSet
            Exit Function
        End If
        Try
            Using MyCommon.LRTadoConn
                If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                    MyCommon.Open_LogixRT()
                End If
                If Not IsValidGUID(GUID, MethodName) Then
                    RetCode = StatusCodes.INVALID_GUID
                    RetMsg = "Failure.Invalid GUID"
                Else
                    MyCommon.QueryStr = "SELECT  AdminUserID, FirstName, LastName, UserName, Email, LastLogin, JobTitle FROM AdminUsers"
                    dtUserDetails = MyCommon.LRT_Select
                    If dtUserDetails.Rows.Count > 0 Then
                        For Each drow As DataRow In dtUserDetails.Rows
                            drow("Email") = MyCryptlib.SQL_StringDecrypt(drow("Email"))
                        Next
                        RetMsg = "Success"
                    Else
                        RetCode = StatusCodes.INVALID_USERID
                        RetMsg = "Failure.No Users exist in the database"
                    End If
                End If
            End Using
        Catch ex As Exception
            RetCode = StatusCodes.APPLICATION_EXCEPTION
            RetMsg = "Failure.Application " & ex.ToString
        Finally
            If dtStatus IsNot Nothing AndAlso dtStatus.Rows.Count = 0 Then
                row = dtStatus.NewRow()
                row.Item("StatusCode") = RetCode
                row.Item("Description") = RetMsg
                dtStatus.Rows.Add(row)
            End If
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
            If RetCode = StatusCodes.SUCCESS Then
                dtUserDetails.TableName = "UserDetails"
                dtUserDetails.AcceptChanges()
                ResultSet.Tables.Add(dtUserDetails.Copy())
            End If

        End Try
        Return ResultSet
    End Function
    ''Common Methods

    Private Function IsValidGUID(ByVal GUID As String, ByVal MethodName As String) As Boolean

        Dim IsValid As Boolean = False
        Dim ConnInc As New Copient.ConnectorInc
        Dim MsgBuf As New StringBuilder()
        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            IsValid = ConnInc.IsValidConnectorGUID(MyCommon, CONNECTOR_ID, GUID)
        Catch ex As Exception
            IsValid = False
        End Try
        ' Log the call

        Try
            MsgBuf.Append(IIf(IsValid, "Validated call to ", "Invalid call to "))
            MsgBuf.Append(MethodName)
            MsgBuf.Append(" from GUID: ")
            MsgBuf.Append(GUID)
            MsgBuf.Append(" and IP: " & HttpContext.Current.Request.UserHostAddress)
            Copient.Logger.Write_Log(UMLogFile, MsgBuf.ToString, True)
        Catch ex As Exception
            ' ignore
        End Try

        Return IsValid
    End Function

    Public Function isValidInputText(ByVal pvalue As String, Optional ByVal pvalueMandatory As Boolean = True) As Boolean

        Dim flag As Boolean
        Try
            If pvalueMandatory Then
                If (pvalue = "") Then
                    flag = False
                End If
            Else
                flag = True
            End If
            If (pvalue = "") Then
                Return flag
            End If
            If (pvalue.Contains("'") Or pvalue.Contains("""")) Then
                Return False
            End If
            flag = True
        Catch ex As Exception
            flag = False
        End Try
        Return flag
    End Function
    '''''

    Function EmailAddressCheck(ByVal emailAddress As String) As Boolean
        Return New EmailAddressAttribute().IsValid(emailAddress)
    End Function

    Private Function UBound(ByVal aryTextFile As String()) As Integer

        Dim UB As Integer
        UB = aryTextFile.GetUpperBound(0)
        '   Throw New NotImplementedException
        Return UB
    End Function

    Private Function IsValidRoleId(ByVal RoleId As Integer) As Boolean

        Dim dt As DataTable
        MyCommon.QueryStr = "Select RoleName from AdminRoles Where RoleID=@RoleID"
        MyCommon.DBParameters.Add("@RoleID", SqlDbType.Int).Value = RoleId
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        If dt.Rows.Count > 0 Then
            Return True
        Else
            Return False

        End If
    End Function

    Private Function UserRoleExists(ByVal u_UID As Integer, ByVal RoleID As Integer) As Boolean

        Dim AddRole As DataTable
        Dim RoleExists As Boolean = False
        MyCommon.QueryStr = "select RoleID from AdminUserRoles with (NoLock) where AdminUserID=" & u_UID & " and RoleID=" & RoleID & ";"
        AddRole = MyCommon.LRT_Select()
        If AddRole.Rows.Count > 0 Then
            RoleExists = True
        End If
        Return RoleExists
    End Function

End Class

