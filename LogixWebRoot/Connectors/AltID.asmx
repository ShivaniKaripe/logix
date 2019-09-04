<%@ WebService Language="VB" Class="Service" %>

Imports System
Imports System.Web.Services
Imports System.Data
Imports Copient.CryptLib

Imports Copient.commonShared

<WebService(Namespace:="http://www.copienttech.com/LogixCustomerService/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class Service
    Inherits System.Web.Services.WebService
    ' version:7.3.1.138972.Official Build (SUSDAY10202)
    ' Return error codes
    Public Const ERROR_NONE As Integer = 0

    Public Enum AltIDResponse As Integer
        SUCCESS = 0
        ALTIDINUSE = 1
        MEMBERNOTFOUND = 2
        INCORRECTARGUMENTS = 3
        BANNERNOTFOUND = 4
        MEMBERINOTHERBANNER = 5
        INVALIDCONFIGURATION = 6
        ERROR_APPLICATION = 9
    End Enum

    Private MyCommon As New Copient.CommonInc
    Private MyAltID As New Copient.AlternateID
    Private MyCryptlib As New Copient.CryptLib
    Private Function IsValidGUID(ByVal GUID As String) As Boolean
        Dim IsValid As Boolean = False
        Dim ConnInc As New Copient.ConnectorInc

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            IsValid = ConnInc.IsValidConnectorGUID(MyCommon, 1, GUID)
        Catch ex As Exception
            IsValid = False
        End Try

        Return IsValid
    End Function

    <WebMethod()> _
    Public Function CreateUpdate(ByVal GUID As String, ByVal MemberCard As String, ByVal AltID As String, ByVal Email As String, ByVal FirstName As String, ByVal LastName As String, ByVal BannerID As String) As Integer
        '    Dim CustomerPK As Long
        'Dim ErrorCode As Integer = ERROR_NONE
        Dim ErrorMsg As String = ""
        Dim dt As DataTable
        Dim AltIDTableColumn As String
        Dim CUDRESPONSE As Copient.AlternateID.CreateUpdateResponse
        Dim LogFile As String = "AltIDWebService." & MyCommon.Leading_Zero_Fill(Year(Today), 4) & MyCommon.Leading_Zero_Fill(Month(Today), 2) & MyCommon.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
        Dim LogFileRejection As String = "AltIDWebService.Rejection." & MyCommon.Leading_Zero_Fill(Year(Today), 4) & MyCommon.Leading_Zero_Fill(Month(Today), 2) & MyCommon.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
        Dim CustomerPK As Integer
        Dim BanID As Integer

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
            MyCommon.Write_Log(LogFile, "Entering CreateUpdate", True)
            '
            AltIDTableColumn = MyCommon.Fetch_SystemOption(60)

            ' first we need to determine if they send the right amount of stuff
            If (MemberCard.Length < 1 Or AltID.Length < 1 Or Not IsValidGUID(GUID)) Then
                CUDRESPONSE = Copient.AlternateID.CreateUpdateResponse.INCORRECTARGUMENTS
                MyCommon.Write_Log(LogFile, "Incorrect Arguments Specified", True)

            ElseIf (Not Integer.TryParse(BannerID, BanID)) Then
                CUDRESPONSE = Copient.AlternateID.CreateUpdateResponse.INCORRECTARGUMENTS
                MyCommon.Write_Log(LogFile, "BannerID should be numeric", True)

            ElseIf (AltIDTableColumn Is Nothing OrElse AltIDTableColumn.Trim().Length = 0) Then
                CUDRESPONSE = AltIDResponse.INVALIDCONFIGURATION
                MyCommon.Write_Log(LogFile, "Invalid configuration for Alternate Identifier.  Missing an AltID table column name in the value for system option 60. " & _
                                            "The CreateUpdate operation should only be set when using the legacy version of AltID that specifies the column name used for AltID. " & _
                                            "The new version of AltID uses the CardTypes table with an AltID card type to determine the alternate identifier.  When using the new  " & _
                                            "method of alternate identifier, the CreateUpdateCard operation of the AltID web service should be invoked instead of the CreateUpdate method.")
            ElseIf (MyCommon.AllowToProcessCustomerCard(MemberCard, CardTypes.CUSTOMER, Nothing) = False) Then
                CUDRESPONSE = Copient.AlternateID.CreateUpdateResponse.INCORRECTARGUMENTS
                MyCommon.Write_Log(LogFile, "Customer card Id should be numeric", True)

            ElseIf (MyCommon.AllowToProcessCustomerCard(AltID, CardTypes.ALTERNATEID, Nothing) = False) Then
                CUDRESPONSE = Copient.AlternateID.CreateUpdateResponse.INCORRECTARGUMENTS
                MyCommon.Write_Log(LogFile, "AltID should be numeric", True)

            Else

                MemberCard = MyCommon.Pad_ExtCardID(MemberCard, CardTypes.CUSTOMER)

                ' lets make sure the AltID isn't already in use by another card
                MyCommon.QueryStr = "select CustomerPK from CardIDs where CardTypeID=0 and ExtCardID = @ExtCardID"
                MyCommon.DBParameters.Add("@ExtCardID", SqlDbType.NVarChar).Value = MyCryptlib.SQL_StringEncrypt(MemberCard, True)
                dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixXS)

                'MyCommon.Write_Log(LogFile, MyCommon.QueryStr, True)

                If (dt.Rows.Count < 1) Then
                    CUDRESPONSE = Copient.AlternateID.CreateUpdateResponse.MEMBERNOTFOUND
                    MyCommon.Write_Log(LogFile, "Unable to find requested memberid " & Copient.MaskHelper.MaskCard(MemberCard, CardTypes.CUSTOMER) & " of type ID " & CardTypes.CUSTOMER, True)
                    MyCommon.Write_Log(LogFileRejection, "Unable to find requested memberid " & MemberCard & " of type ID " & CardTypes.CUSTOMER, True)
                Else

                    ' ok the id sent is not in use so we can store it
                    CustomerPK = dt.Rows(0).Item("CustomerPK")
                    MyCommon.Write_Log(LogFile, "Attempting to update record for CustomerPK " & CustomerPK & " AltID: " & Copient.MaskHelper.MaskCard(AltID, CardTypes.ALTERNATEID) & " BannerID: " & BannerID, True)

                    ' going to call the dll function here to do the creation
                    CUDRESPONSE = MyAltID.UpdateCustomerAltID(CustomerPK, AltID, BannerID)
                    '
                    MyCommon.Write_Log(LogFile, "ErrMessage: " & MyAltID.ErrorMessage, True)
                End If

            End If

        Catch ex As Exception

            '   CUDRESPONSE = Copient.AlternateID.CreateUpdateResponse.ERROR_APPLICATION
            ErrorMsg = "Add Customer encountered the following error: " & ex.ToString
            MyCommon.Write_Log(LogFile, ErrorMsg & " error: " & ex.ToString, True)

        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try

        ' ErrorCode = CUDRESPONSE
        If (CUDRESPONSE = Copient.AlternateID.CreateUpdateResponse.SUCCESS) Then
            MyCommon.Write_Log(LogFile, "AltID Record updated", True)
        ElseIf (CUDRESPONSE = Copient.AlternateID.CreateUpdateResponse.MEMBERNOTFOUND) Then
            ' the member wasn't found so lets create it
            'run the Stored Procedure to insert a record into Customers (returns the new PrimaryKey)

            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            MyCommon.QueryStr = "dbo.pa_CPE_IN_CreateCustomer"
            MyCommon.Open_LXSsp()
            MyCommon.LXSsp.Parameters.Add("@InitialCardID", SqlDbType.NVarChar, 400).Value = MyCryptlib.SQL_StringEncrypt(MemberCard, True)
            MyCommon.LXSsp.Parameters.Add("@CustomerTypeID", SqlDbType.Int).Value = 0
            MyCommon.LXSsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = -9
            MyCommon.LXSsp.Parameters.Add("@BannerID", SqlDbType.Int).Value = BannerID
            MyCommon.LXSsp.Parameters.Add("@InitialCardTypeID", SqlDbType.Int).Value = 0
            MyCommon.LXSsp.Parameters.Add("@PKID", SqlDbType.BigInt).Direction = ParameterDirection.Output
            MyCommon.LXSsp.Parameters.Add("@ExtCardIDOriginal", SqlDbType.NVarChar, 400).Value = MyCryptlib.SQL_StringEncrypt(MemberCard)
            MyCommon.LXSsp.ExecuteNonQuery()
            CustomerPK = MyCommon.LXSsp.Parameters("@PKID").Value
            MyCommon.Close_LXSsp()

            ' well we made it here so we have the new CustomerPK
            MyCommon.Write_Log(LogFile, "Attempting to update record for newly created CustomerPK " & CustomerPK & " AltID: " & Copient.MaskHelper.MaskCard(AltID, CardTypes.ALTERNATEID) & " BannerID: " & BannerID, True)
            ' going to call the dll function here to do the creation
            CUDRESPONSE = MyAltID.UpdateCustomerAltID(CustomerPK, AltID, BannerID)

            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()

        Else
            MyCommon.Write_Log(LogFile, "AltID Record Not Updated Error Code: " & CUDRESPONSE, True)
        End If

        Return CUDRESPONSE
    End Function

    <WebMethod()> _
    Public Function Request(ByVal GUID As String, ByVal MemberCard As String, ByVal Email As String, ByVal FirstName As String, ByVal LastName As String, ByVal BannerID As Integer) As String
        Dim dt As DataTable
        Dim AltIDTableColumn As String
        Dim AltIDReturn As String = ""
        Dim LogFile As String = "AltIDWebService." & MyCommon.Leading_Zero_Fill(Year(Today), 4) & MyCommon.Leading_Zero_Fill(Month(Today), 2) & MyCommon.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
        Dim LogFileRejection As String = "AltIDWebService.Rejection." & MyCommon.Leading_Zero_Fill(Year(Today), 4) & MyCommon.Leading_Zero_Fill(Month(Today), 2) & MyCommon.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"

        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
        If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
    
        AltIDTableColumn = MyCommon.Fetch_SystemOption(60)
        ' MyCommon.Write_Log(LogFile, "MemberCard before padding " & MemberCard, True)
        MemberCard = MyCommon.Pad_ExtCardID(MemberCard, Copient.commonShared.CardTypes.CUSTOMER)
    
        ' added
        If Email Is Nothing OrElse Email.Trim = "" OrElse Not String.IsNullOrEmpty(Email) Then
            MyCommon.QueryStr = "select * from CustomerExt where email=@Email"
            MyCommon.DBParameters.Add("@Email", SqlDbType.NVarChar).Value = MyCryptlib.SQL_StringEncrypt(Email)
            dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixXS)
            If (dt.Rows.Count = 0) Then
                'AltIDReturn = "Incorrect Arguments specified:EmailID not found."
                AltIDReturn = ""
            Else
                If FirstName Is Nothing OrElse FirstName.Trim = "" OrElse Not String.IsNullOrEmpty(FirstName) OrElse LastName Is Nothing OrElse LastName.Trim = "" OrElse Not String.IsNullOrEmpty(LastName) Then
                    MyCommon.QueryStr = "select * from Customers where FirstName=@FirstName and LastName=@LastName"
                    MyCommon.DBParameters.Add("@FirstName", SqlDbType.NVarChar).Value = FirstName
                    MyCommon.DBParameters.Add("@LastName", SqlDbType.NVarChar).Value = LastName
                    dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixXS)
           
                    If (dt.Rows.Count = 0) Then
                        'AltIDReturn = "Either Firstname/Lastname not found."
                        AltIDReturn = ""
                    Else
                        If (MemberCard.Length < 1 Or Not IsValidGUID(GUID)) Then
                            MyCommon.Write_Log(LogFile, "Incorrect Arguments Specified", True)
                        ElseIf (AltIDTableColumn Is Nothing OrElse AltIDTableColumn.Trim().Length = 0) Then
                            MyCommon.Write_Log(LogFile, "Invalid configuration for Alternate Identifier.  Missing an AltID table column name in the value for system option 60. " & _
                                                        "The CreateUpdate operation should only be set when using the legacy version of AltID that specifies the column name used for AltID. " & _
                                                        "The new version of AltID uses the CardTypes table with an AltID card type to determine the alternate identifier.  When using the new  " & _
                                                        "method of alternate identifier, the CreateUpdateCard operation of the AltID web service should be invoked instead of the CreateUpdate method.")
                        Else
      
                            ' lets make sure the AltID isn't already in use by another card
                            MyCommon.QueryStr = "select isnull(" & AltIDTableColumn & ",'') as AltID from Customers " & _
                                                "left join CustomerExt on Customers.CustomerPK=CustomerExt.CustomerPK " & _
                                                "where Customers.CustomerPK in (" & _
                                                " select CustomerPK from CardIDs with (NoLock) where CardTypeID=0 and ExtCardID=@ExtCardID" & _
                                                ")"
                            MyCommon.DBParameters.Add("@ExtCardID", SqlDbType.NVarChar).Value = MyCryptlib.SQL_StringEncrypt(MemberCard, True)
                            If BannerID <> 0 Then
                                MyCommon.QueryStr &= " and BannerID=@BannerID"
                                MyCommon.DBParameters.Add("@BannerID", SqlDbType.Int).Value = BannerID
                            End If
                            dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixXS)
      
                            If (dt.Rows.Count < 1) Then
                                MyCommon.Write_Log(LogFile, "Unable to find requested memberid " & Copient.MaskHelper.MaskCard(MemberCard, CardTypes.CUSTOMER) & " of type ID " & CardTypes.CUSTOMER, True)
                                MyCommon.Write_Log(LogFileRejection, "Unable to find requested memberid " & MemberCard & " of type ID " & CardTypes.CUSTOMER, True)
                            Else
                                ' ok the id sent is not in use so we can store it
                                ' Advise whats in Fetch_SystemOption(60)?
                                AltIDReturn = dt.Rows(0).Item("AltID")
                            End If
    
                        End If
                    End If
                End If
            End If
        End If
        If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()

        Return AltIDReturn
    End Function

    <WebMethod()> _
    Public Function BannerList(ByVal GUID As String) As DataTable
        Dim dt As DataTable
        'Dim LogFile As String = "AltIDWebService." & MyCommon.Leading_Zero_Fill(Year(Today), 4) & MyCommon.Leading_Zero_Fill(Month(Today), 2) & MyCommon.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
        dt = New DataTable("Table")
        If (Not IsValidGUID(GUID)) Then
            dt.Columns.Add("GUID", GetType(System.String))
            dt.Rows.Add("Invalid Guid")
            Return dt
        Else
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            MyCommon.QueryStr = "select BannerID, Name from Banners where Deleted=0 and AllBanners=0;"
            dt = MyCommon.LRT_Select
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then
                MyCommon.Close_LogixRT()
            End If
            Return dt
        End If
    End Function

    <WebMethod()> _
    Public Function CreateUpdateCard(ByVal GUID As String, ByVal MemberCard As String, ByVal AltID As String, ByVal BannerID As Integer) As Integer
        Dim MyLookup As New Copient.CustomerLookup()
        Dim ReturnCode As Copient.CustomerAbstract.RETURN_CODE
        'Dim ErrorCode As Integer = ERROR_NONE
        Dim ErrorMsg As String = ""
        Dim CustomerPK As Long = 0
        Dim CustomerBannerID As Integer = 0
        Dim LogFile As String = "AltIDWebService." & MyCommon.Leading_Zero_Fill(Year(Today), 4) & MyCommon.Leading_Zero_Fill(Month(Today), 2) & MyCommon.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
        Dim LogFileRejection As String = "AltIDWebService.Rejection." & MyCommon.Leading_Zero_Fill(Year(Today), 4) & MyCommon.Leading_Zero_Fill(Month(Today), 2) & MyCommon.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
        Dim CUDRESPONSE As AltIDResponse

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

            MyCommon.Write_Log(LogFile, "Beginning AltID CreateUpdateCard...", True)
            MemberCard = MyCommon.Pad_ExtCardID(MemberCard, 0)

            'Identify the banner
            MyCommon.QueryStr = "select Name from Banners with (NoLock) " & _
                                "where BannerID=@BannerID and Deleted=0"
            MyCommon.DBParameters.Add("@BannerID", SqlDbType.Int).Value = BannerID
            Dim dtBanner As DataTable = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)

            'Identify the customer
            MyCommon.QueryStr = "select CI.CustomerPK, IsNULL(C.BannerID, 0) as BannerID from CardIDs as CI with (NoLock) " & _
                                "left join Customers as C on C.CustomerPK=CI.CustomerPK " & _
                                "where CI.ExtCardID=@ExtCardID and CI.CardTypeID=0"
            MyCommon.DBParameters.Add("@ExtCardID", SqlDbType.NVarChar).Value = MyCryptlib.SQL_StringEncrypt(MemberCard, True)
            Dim dtCustomer As DataTable = MyCommon.ExecuteQuery(Copient.DataBases.LogixXS)

            If (Not IsValidGUID(GUID)) Then
                CUDRESPONSE = AltIDResponse.INCORRECTARGUMENTS
                MyCommon.Write_Log(LogFile, "Invalid GUID", True)
            ElseIf (MemberCard.Length = 0) Then
                CUDRESPONSE = AltIDResponse.INCORRECTARGUMENTS
                MyCommon.Write_Log(LogFile, "Invalid member card", True)
            ElseIf (AltID.Length = 0) Then
                CUDRESPONSE = AltIDResponse.INCORRECTARGUMENTS
                MyCommon.Write_Log(LogFile, "Invalid Alt ID", True)
            ElseIf (MyCommon.AllowToProcessCustomerCard(MemberCard, CardTypes.CUSTOMER, Nothing) = False) Then
                CUDRESPONSE = Copient.AlternateID.CreateUpdateResponse.INCORRECTARGUMENTS
                MyCommon.Write_Log(LogFile, "Customer card Id should be numeric", True)
            ElseIf (MyCommon.AllowToProcessCustomerCard(AltID, CardTypes.ALTERNATEID, Nothing) = False) Then
                CUDRESPONSE = Copient.AlternateID.CreateUpdateResponse.INCORRECTARGUMENTS
                MyCommon.Write_Log(LogFile, "AltID should be numeric", True)
            ElseIf (dtBanner.Rows.Count = 0 AndAlso BannerID <> 0) Then
                CUDRESPONSE = AltIDResponse.BANNERNOTFOUND
                MyCommon.Write_Log(LogFile, "Banner " & BannerID & " not found", True)
            Else
                If dtCustomer.Rows.Count > 0 Then
                    CustomerPK = MyCommon.NZ(dtCustomer.Rows(0).Item("CustomerPK"), 0)
                    CustomerBannerID = MyCommon.NZ(dtCustomer.Rows(0).Item("BannerID"), 0)
                End If
                If (CustomerPK > 0) AndAlso (BannerID <> CustomerBannerID) Then
                    'Customer exists, but not in the specified banner.
                    CUDRESPONSE = AltIDResponse.ERROR_APPLICATION
                    MyCommon.Write_Log(LogFile, "Member " & Copient.MaskHelper.MaskCard(MemberCard, CardTypes.CUSTOMER) & " for CustomerPK " & CustomerPK & " exists, but not in the specified banner.", True)
                    MyCommon.Write_Log(LogFileRejection, "Member " & MemberCard & " for CustomerPK " & CustomerPK & " exists, but not in the specified banner.", True)
                Else
                    AltID = MyCommon.Pad_ExtCardID(AltID, CardTypes.ALTERNATEID)
                    'Identify the existence and holder of the AltID card
                    MyCommon.QueryStr = "select CardPK, CustomerPK from CardIDs with (NoLock) " & _
                                        "where ExtCardID=@ExtCardID and CardTypeID=3"
                    MyCommon.DBParameters.Add("@ExtCardID", SqlDbType.NVarChar).Value = MyCryptlib.SQL_StringEncrypt(AltID, True)
                    Dim dtCard As DataTable = MyCommon.ExecuteQuery(Copient.DataBases.LogixXS)

                    If (dtCard.Rows.Count > 0) Then
                        'Specified AltID card already exists, so do nothing.
                        If MyCommon.NZ(dtCard.Rows(0).Item("CustomerPK"), 0) = CustomerPK Then
                            'It belongs to the customer.
                            CUDRESPONSE = AltIDResponse.ALTIDINUSE
                            MyCommon.Write_Log(LogFile, "AltID " & Copient.MaskHelper.MaskCard(AltID, CardTypes.ALTERNATEID) & " is already associated to member " & Copient.MaskHelper.MaskCard(MemberCard, CardTypes.CUSTOMER) & " for CustomerPK " & dtCard.Rows(0).Item("CustomerPK") & ".", True)
                            MyCommon.Write_Log(LogFileRejection, "AltID " & AltID & " is already associated to member " & MemberCard & " for CustomerPK " & dtCard.Rows(0).Item("CustomerPK") & ".", True)
                        Else
                            'It belongs to someone else
                            CUDRESPONSE = AltIDResponse.ALTIDINUSE
                            MyCommon.Write_Log(LogFile, "AltID " & Copient.MaskHelper.MaskCard(AltID, CardTypes.ALTERNATEID) & " for CustomerPK " & dtCard.Rows(0).Item("CustomerPK") & " is already associated to another member.", True)
                            MyCommon.Write_Log(LogFileRejection, "AltID " & AltID & " for CustomerPK " & dtCard.Rows(0).Item("CustomerPK") & " is already associated to another member.", True)
                        End If
                    Else
                        'Specified AltID card doesn't exist.
                        If dtCustomer.Rows.Count = 0 Then
                            'Create a new customer with the MemberCard number provided.
                            MyCommon.Write_Log(LogFile, "Member " & Copient.MaskHelper.MaskCard(MemberCard, CardTypes.CUSTOMER) & " of type ID " & CardTypes.CUSTOMER & " does not exist.  Creating...", True)
                            MyCommon.Write_Log(LogFileRejection, "Member " & MemberCard & " of type ID " & CardTypes.CUSTOMER & " does not exist.  Creating...", True)
                            MyCommon.QueryStr = "dbo.pa_CPE_IN_CreateCustomer"
                            MyCommon.Open_LXSsp()
                            MyCommon.LXSsp.Parameters.Add("@InitialCardID", SqlDbType.NVarChar, 400).Value = MyCryptlib.SQL_StringEncrypt(MemberCard, True)
                            MyCommon.LXSsp.Parameters.Add("@CustomerTypeID", SqlDbType.Int).Value = 0
                            MyCommon.LXSsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = -9
                            MyCommon.LXSsp.Parameters.Add("@BannerID", SqlDbType.Int).Value = BannerID
                            MyCommon.LXSsp.Parameters.Add("@InitialCardTypeID", SqlDbType.Int).Value = 0
                            MyCommon.LXSsp.Parameters.Add("@PKID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                            MyCommon.LXSsp.Parameters.Add("@ExtCardIDOriginal", SqlDbType.NVarChar, 400).Value = MyCryptlib.SQL_StringEncrypt(MemberCard)
                            MyCommon.LXSsp.ExecuteNonQuery()
                            CustomerPK = MyCommon.LXSsp.Parameters("@PKID").Value
                            MyCommon.Close_LXSsp()
                            MyCommon.Activity_Log2(25, 10, CustomerPK, 1, Copient.PhraseLib.Lookup("history.customer-add-customer", 1), CustomerPK)
                            MyCommon.Write_Log(LogFile, "Member " & Copient.MaskHelper.MaskCard(MemberCard, CardTypes.CUSTOMER) & " for CustomerPK " & CustomerPK & " created.", True)
                        End If
                        'Create a new AltID card for the customer.
                        If MyLookup.AddCardToCustomer(CustomerPK, AltID, 3, 1, ReturnCode) Then
                            'Success.
                            MyCommon.Write_Log(LogFile, "Member " & Copient.MaskHelper.MaskCard(MemberCard, CardTypes.CUSTOMER) & " now has AltID card " & Copient.MaskHelper.MaskCard(AltID, CardTypes.ALTERNATEID) & " for CustomerPK " & CustomerPK & ".", True)
                        Else
                            MyCommon.Write_Log(LogFile, "Member " & Copient.MaskHelper.MaskCard(MemberCard, CardTypes.CUSTOMER) & " failed to receive AltID card " & Copient.MaskHelper.MaskCard(AltID, CardTypes.ALTERNATEID) & " for CustomerPK " & CustomerPK & ".", True)
                            MyCommon.Write_Log(LogFileRejection, "Member " & MemberCard & " failed to receive AltID card " & AltID & " for CustomerPK " & CustomerPK & ".", True)
                            CUDRESPONSE = AltIDResponse.ERROR_APPLICATION
                        End If
                    End If
                End If
            End If

        Catch ex As Exception
            ErrorMsg = "AltID CreateUpdateCard method encountered the following error: " & ex.ToString
            CUDRESPONSE = AltIDResponse.ERROR_APPLICATION
            MyCommon.Write_Log(LogFile, ErrorMsg & " error: " & ex.ToString, True)
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try

        Return CUDRESPONSE
    End Function

    <WebMethod()> _
    Public Function RequestCard(ByVal GUID As String, ByVal MemberCard As String, ByVal BannerID As Integer) As String
        Dim CustomerPK As Long = 0
        Dim LogFile As String = "AltIDWebService." & MyCommon.Leading_Zero_Fill(Year(Today), 4) & MyCommon.Leading_Zero_Fill(Month(Today), 2) & MyCommon.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
        Dim LogFileRejection As String = "AltIDWebService.Rejection." & MyCommon.Leading_Zero_Fill(Year(Today), 4) & MyCommon.Leading_Zero_Fill(Month(Today), 2) & MyCommon.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
        Dim AltID As String = ""

        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
        If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

        MemberCard = MyCommon.Pad_ExtCardID(MemberCard, 0)

        'MemberCard = MemberCard.PadLeft(CardLen, "0")

        MyCommon.QueryStr = "select BannerID from Banners with (NoLock) " & _
                            "where BannerID=@BannerID and Deleted=0"
        MyCommon.DBParameters.Add("@BannerID", SqlDbType.Int).Value = BannerID
        Dim dtBanner As DataTable = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)

        MyCommon.QueryStr = "select CI.CustomerPK from CardIDs as CI with (NoLock) " & _
                            "left join Customers as C on C.CustomerPK=CI.CustomerPK " & _
                            "where CI.ExtCardID=@ExtCardID and CI.CardTypeID=0"
        MyCommon.DBParameters.Add("@ExtCardID", SqlDbType.NVarChar).Value = MyCryptlib.SQL_StringEncrypt(MemberCard, True)
        If BannerID <> 0 Then
            MyCommon.QueryStr &= " and C.BannerID=@BannerID"
            MyCommon.DBParameters.Add("@BannerID", SqlDbType.Int).Value = BannerID
        End If
        Dim dtCustomer As DataTable = MyCommon.ExecuteQuery(Copient.DataBases.LogixXS)

        If (Not IsValidGUID(GUID)) Then
            MyCommon.Write_Log(LogFile, "Incorrect arguments: invalid GUID", True)
        ElseIf (MemberCard.Length = 0) Then
            MyCommon.Write_Log(LogFile, "Incorrect arguments: invalid member card", True)
        ElseIf (dtBanner.Rows.Count = 0 AndAlso BannerID <> 0) Then
            MyCommon.Write_Log(LogFile, "Banner " & BannerID & " not found", True)
        ElseIf (dtCustomer.Rows.Count = 0) Then
            MyCommon.Write_Log(LogFile, "Member " & Copient.MaskHelper.MaskCard(MemberCard, CardTypes.CUSTOMER) & " of type ID " & CardTypes.CUSTOMER & " not found in banner " & BannerID & ".  Card was padded to " & MemberCard.Length.ToString() & " digits.", True)
            MyCommon.Write_Log(LogFileRejection, "Member " & MemberCard & " of type ID " & CardTypes.CUSTOMER & " not found in banner " & BannerID & ".  Card was padded to " & MemberCard.Length.ToString() & " digits.", True)
        Else
            CustomerPK = MyCommon.NZ(dtCustomer.Rows(0).Item("CustomerPK"), 0)
            MyCommon.QueryStr = "select ExtCardIDOriginal as ExtCardID from CardIDs with (NoLock) " & _
                                "where CustomerPK=@CustomerPK and CardTypeID=3"
            MyCommon.DBParameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
            Dim dtCard As DataTable = MyCommon.ExecuteQuery(Copient.DataBases.LogixXS)

            If (dtCard.Rows.Count > 0) Then
                AltID = MyCryptlib.SQL_StringDecrypt(dtCard.Rows(0).Item("ExtCardID").ToString())
            End If
        End If

        If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()

        Return AltID
    End Function

    <WebMethod()> _
    Public Function DeleteCard(ByVal GUID As String, ByVal AltID As String, ByVal BannerID As Integer) As Integer
        Dim DelRESPONSE As AltIDResponse
        Dim MyLookup As New Copient.CustomerLookup()
        Dim ReturnCode As Copient.CustomerAbstract.RETURN_CODE
        Dim ErrorMsg As String = ""
        Dim dt As DataTable
        Dim LogFile As String = "AltIDWebService." & MyCommon.Leading_Zero_Fill(Year(Today), 4) & MyCommon.Leading_Zero_Fill(Month(Today), 2) & MyCommon.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
        Dim LogFileRejection As String = "AltIDWebService.Rejection." & MyCommon.Leading_Zero_Fill(Year(Today), 4) & MyCommon.Leading_Zero_Fill(Month(Today), 2) & MyCommon.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
        Dim CustomerPK As Long = 0
        Dim CustomerBannerID As Integer = 0
        Dim CardPK As Long = 0

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
            MyCommon.Write_Log(LogFile, "Deleting Alternate ID " & Copient.MaskHelper.MaskCard(AltID, CardTypes.ALTERNATEID) & " of type ID " & CardTypes.ALTERNATEID, True)

            If (Not IsValidGUID(GUID)) Then
                DelRESPONSE = AltIDResponse.ERROR_APPLICATION
                MyCommon.Write_Log(LogFile, "Invalid GUID", True)
            ElseIf (AltID.Length < 1) Then
                DelRESPONSE = AltIDResponse.ERROR_APPLICATION
                MyCommon.Write_Log(LogFile, "Invalid Alternate ID", True)
            Else
                'Identify the banner
                MyCommon.QueryStr = "select Name from Banners with (NoLock) " & _
                                    "where BannerID=@BannerID and Deleted=0"
                MyCommon.DBParameters.Add("@BannerID", SqlDbType.Int).Value = BannerID
                dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)

                If (dt.Rows.Count = 0 AndAlso BannerID <> 0) Then
                    DelRESPONSE = AltIDResponse.BANNERNOTFOUND
                    MyCommon.Write_Log(LogFile, "Banner " & BannerID & " not found", True)
                Else
                    'Identify the customer

                    AltID = MyCommon.Pad_ExtCardID(AltID, 3)

                    MyCommon.QueryStr = "select CustomerPK, CardPK from CardIDs where CardTypeID=3 and ExtCardID=@ExtCardID"
                    MyCommon.DBParameters.Add("@ExtCardID", SqlDbType.NVarChar).Value = MyCryptlib.SQL_StringEncrypt(AltID, True)
                    dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixXS)

                    If (dt.Rows.Count) < 1 Then
                        DelRESPONSE = AltIDResponse.MEMBERNOTFOUND
                        MyCommon.Write_Log(LogFile, "Alternate ID " & Copient.MaskHelper.MaskCard(AltID, CardTypes.ALTERNATEID) & " of type ID " & CardTypes.ALTERNATEID & " is not found", True)
                        MyCommon.Write_Log(LogFileRejection, "Alternate ID " & AltID & " of type ID " & CardTypes.ALTERNATEID & " is not found", True)
                    Else
                        CustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)
                        CardPK = MyCommon.NZ(dt.Rows(0).Item("CardPK"), 0)
                        MyCommon.QueryStr = "select BannerID from Customers where CustomerPK =@CustomerPK"
                        MyCommon.DBParameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
                        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixXS)

                        If (dt.Rows.Count < 1) Then
                            DelRESPONSE = AltIDResponse.MEMBERNOTFOUND
                            MyCommon.Write_Log(LogFile, "Alternate ID " & Copient.MaskHelper.MaskCard(AltID, CardTypes.ALTERNATEID) & " for CustomerPK " & CustomerPK & " is not found", True)
                            MyCommon.Write_Log(LogFileRejection, "Alternate ID " & AltID & " for CustomerPK " & CustomerPK & " is not found", True)
                        Else
                            CustomerBannerID = MyCommon.NZ(dt.Rows(0).Item("BannerID"), 0)
                            If (CustomerBannerID <> BannerID) Then
                                'Customer exists, but not in the specified banner.
                                DelRESPONSE = AltIDResponse.MEMBERINOTHERBANNER
                                MyCommon.Write_Log(LogFile, "Alternate ID " & Copient.MaskHelper.MaskCard(AltID, CardTypes.ALTERNATEID) & " for CustomerPK " & CustomerPK & " exists, but not in the specified banner.", True)
                                MyCommon.Write_Log(LogFileRejection, "Alternate ID " & AltID & " for CustomerPK " & CustomerPK & " exists, but not in the specified banner.", True)
                            Else
                                'Delete AltID
                                If (MyLookup.RemoveCardFromCustomer(CustomerPK, CardPK, ReturnCode)) Then
                                    DelRESPONSE = AltIDResponse.SUCCESS
                                    MyCommon.Write_Log(LogFile, "Deleted AltID " & Copient.MaskHelper.MaskCard(AltID, CardTypes.ALTERNATEID) & " for CustomerPK " & CustomerPK, True)
                                End If
                            End If
                        End If
                    End If
                End If
            End If

        Catch ex As Exception
            ErrorMsg = "AltID Delete method encountered the following error: " & ex.ToString
            DelRESPONSE = AltIDResponse.ERROR_APPLICATION
            MyCommon.Write_Log(LogFile, ErrorMsg & " error: " & ex.ToString, True)
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try

        Return DelRESPONSE
    End Function
End Class