Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.Odbc
Imports System.IO
Imports System.Xml
Imports System.Text
Imports System.Security.Cryptography
Imports Copient.CommonInc
Imports Copient.PhraseLib
Imports Copient.CryptLib
Imports Copient.ExternalRewards
Imports Copient.ExternalCustomerAccounts
Imports Copient.AMSPreferenceLib
<WebService(Namespace:="http://www.copienttech.com/CmConnector/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class CMConnector
  Inherits System.Web.Services.WebService

  Private Enum DebugState
    BeginTime = -1
    CurrentTime = 0
    EndTime = 1
  End Enum

  Private Enum DuplicateStatus
    HasNotBeenChecked = -1
    IsNotDuplicate = 0
    IsDuplicate = 1
  End Enum

  Private Enum ErrorType
    SqlServer = 0
    General = 1
    Db2ExtCustomerIdNotFound = 101
    Db2General = 102
  End Enum

  Private Enum MessageType
    Info = 0
    Warning = 1
    AppError = 2
    SysError = 3
    Debug = 4
  End Enum

  Private Enum FetchTargetedListMode
    Normal = 0
    FobOnly = 1
    FobTargeted = 2
  End Enum

  Private Const sXmls As String = "http://www.ncr.com/rsd/cm/accounts/1.0"
  Private Const sVersion As String = "7.3.1.138972"
  Private Const iOptionGlobalStartTime As Integer = 31
  Private Const iOptionGlobalEndTime As Integer = 34
  Private Const sDateFormat As String = "yyyy-MM-ddTHH:mm:ss"
  Private Const sLogFileName As String = "CMConnector"
  Private Const sOkStatus As String = "<Response Status=""Ok"" />" & vbCrLf
  Private Const sOkStatusDistRequestTrue As String = "<Response Status=""Ok"" DistributionRequested=""True"" />" & vbCrLf
  Private Const sOkStatusDistRequestFalse As String = "<Response Status=""Ok"" DistributionRequested=""False"" />" & vbCrLf
  Private Const sDefaultDb2Connection As String = "DSN=DEVDB2NW;UID=asddev01;PWD=123456"
  Private Const iDb2Connection As Integer = 5
  Private Const iOptionOfferDistributionEnabled As Integer = 9
  Private Const sAppName As String = "CMConnector"
  Private Const scDisabled As String = "DISABLED"
  Private Const scInactive As String = "INACTIVE"
  Private Const scActive As String = "ACTIVE"
  Private Const scDashes As String = "------------------------------------"
  Private Const bUseStoreIdForLog As Boolean = False
  Private Const bDebugLogOn As Boolean = True
  Private Const iAutoHouseholdCustGrpOptionId As Integer = 24
  Private Const iPrimaryIdOptionId As Integer = 61
  Private Const iSecondaryIdOptionId As Integer = 62
  Private Const iMemberIdOptionId As Integer = 63
  Private Const iAlternateIdOptionId As Integer = 64
  Private Const iInterpretCardStatusOptionId As Integer = 73
  Private Const iEmployeeIdOptionId As Integer = 147

  Private sCurrentStoreId As String = ""
  Private i32CurrentLocalServerId As Int32 = -1
  Private i64CurrentLocationId As Int64 = -1
  Private bDuplicateTransactionXml As DuplicateStatus = DuplicateStatus.HasNotBeenChecked
  Private sXmlMd5Hash As String
  Private sInputForLog As String = ""
  Private sCurrentMethod As String = ""
  Private sInstallationName As String = ""
  Private MyCommon As Copient.CommonInc
  Private MyCryptLib As Copient.CryptLib
  Private PrefLib As Copient.AMSPreferenceLib
  Private eDefaultErrorType As ErrorType = ErrorType.General
  Private sDb2Connection As String = ""
  Private bDb2TestMode As Boolean = False
  Private aStatusMsgs() As String = {"Ready", "Notified", "Processed", "Error", "Skipped"}
  Private DebugStartTimes As New ArrayList()
  Private bBeginTransactionEX As Boolean = False
  Private sDefaultFirstName As String = ""
  Private sDefaultLastName As String = ""

  ' Card Id mapping variables
  Private iIdTypeForMemberId As Integer = 0
  Private iIdTypeForPrimaryId As Integer = -1
  Private iIdTypeForSecondaryId As Integer = -1
  Private iIdTypeForAlternateId As Integer = -1
  Private iIdTypeForEmployeeId As Integer = -1
  Private cardValidationResp As CardValidationResponse
  
  'DB2 String
  Private user As String
  Private euser As String
  Private pwd As String
  Private epwd As String
  Private upos As Integer
  Private ppos As Integer
  Private pend As Integer
  Private tempstr As String



  <WebMethod()> _
  Public Function AboutThisService() As String
    Dim sResponse As String

    sResponse = "CM Connector (Version " & sVersion & ")"

    Return sResponse
  End Function

  <WebMethod()> _
  Public Function GetStatus(ByVal iMode As Integer) As String
    Dim sResponse As String

    If iMode = 0 Then
      sResponse = sOkStatus
    Else
      sResponse = GetSqlStatus(iMode, "")
    End If

    Return sResponse
  End Function

  <WebMethod()> _
  Public Function GetStatusUpdateHealth(ByVal iMode As Integer, ByVal sStoreId As String) As String
    Dim sResponse As String

    sStoreId = sStoreId.Trim(" ")
    If iMode = 0 And sStoreId.Length = 0 Then
      sResponse = sOkStatus
    Else
      sResponse = GetSqlStatus(iMode, sStoreId)
    End If

    Return sResponse
  End Function

  <WebMethod()> _
  Public Function GetAccount(ByVal CustomerId As String, ByVal IdType As Integer, ByVal BusinessDate As Date, ByVal LaneId As String, ByVal TransactionNumber As Integer, ByVal StoreId As Int16) As String
    Dim sXml As String

    sXml = GetAccountInfo(CustomerId, IdType, BusinessDate, LaneId, TransactionNumber, StoreId.ToString())

    Return sXml
  End Function

  <WebMethod()> _
  Public Function GetAccountEx(ByVal CustomerId As String, ByVal IdType As Integer, ByVal BusinessDate As Date, ByVal LaneId As String, ByVal TransactionNumber As Integer, ByVal StoreId As String) As String
    Dim sXml As String
    ' *******************************************************************************************************
    ' Note: IdType is the ACS ID type, so this ACS ID type must be translated to a corresponding AMS ID type
    ' *******************************************************************************************************

    StoreId = StoreId.Trim(" ")
    If IdType = 5 Then
      sXml = GetCouponInfo(CustomerId, 999, BusinessDate, LaneId, TransactionNumber, StoreId)
    Else
      sXml = GetAccountInfo(CustomerId, IdType, BusinessDate, LaneId, TransactionNumber, StoreId)
    End If

    Return sXml
  End Function

  <WebMethod()> _
  Public Function PutTransaction(ByVal TransactionXml As String) As String
    Dim sResponse As String

    sResponse = ProcessTransactionXml(TransactionXml, "")

    Return sResponse
  End Function

  <WebMethod()> _
  Public Function PutTransactionUpdateHealth(ByVal TransactionXml As String, ByVal sStoreId As String) As String
    Dim sResponse As String

    sStoreId = sStoreId.Trim(" ")
    sResponse = ProcessTransactionXml(TransactionXml, sStoreId)

    Return sResponse
  End Function

  <WebMethod()> _
  Public Function RequestIPL(ByVal sStoreId As String) As String
    Dim sResponse As String

    sStoreId = sStoreId.Trim(" ")
    sResponse = SetIpl(sStoreId)

    Return sResponse
  End Function

  <WebMethod()> _
  Public Function GetDistributionList(ByVal sStoreId As String) As String
    Dim sResponse As String

    sStoreId = sStoreId.Trim(" ")
    sResponse = GetFileList(sStoreId)

    Return sResponse
  End Function

  <WebMethod()> _
  Public Function UpdateDistributionStatus(ByVal sStoreId As String, ByVal sStatusXml As String) As String
    Dim sResponse As String

    sStoreId = sStoreId.Trim(" ")
    sResponse = UpdateFileStatus(sStoreId, sStatusXml)

    Return sResponse
  End Function

  <WebMethod()> _
  Public Function GetLogixLocationInfo(ByVal sStoreId As String) As String
    Dim sResponse As String

    sStoreId = sStoreId.Trim(" ")
    sResponse = GetLocationInfo(sStoreId)

    Return sResponse
  End Function

  <WebMethod()> _
  Public Function GetAccountPending(ByVal CustomerId As String) As String
    Dim sXml As String

    sXml = GetPendingInfo(CustomerId)

    Return sXml
  End Function

  <WebMethod()> _
  Public Function DeletePending(ByVal CartId As String) As String
    Dim sResponse As String

    sResponse = DeletePendingInfo(CartId)

    Return sResponse
  End Function

  Private Function GetSqlStatus(ByVal iMode As Integer, ByVal sStoreId As String) As String
    Dim sResponse As String = sOkStatus
    Dim bTestLocation As Boolean = False

    Dim dt As DataTable

    ' iMode = 1 test GetAccount databases (XS & RT)
    ' iMode = 2 test PutTransaction databases (XS, RT, WH)
    ' iMode = 99 - generate Sql connection error (for client testing)

    Try
      sCurrentStoreId = sStoreId
      sCurrentMethod = "GetSqlStatus"
      sInputForLog = "*(Type=Input) (Method=GetSqlStatus)] - (Mode='" & iMode.ToString & "')"
      MyCommon = New Copient.CommonInc
      MyCommon.AppName = sAppName
      sStoreId = sStoreId.Trim(" ")

      If iMode = 0 Then
        WriteDebug("GetStatusUpdateHealth", DebugState.BeginTime)
      Else
        WriteDebug("GetStatus(" & iMode & ")", DebugState.BeginTime)
      End If

      If iMode = 1 Or iMode = 2 Then
        If MyCommon.LRTadoConn.State <> ConnectionState.Open Then
          MyCommon.Open_LogixRT()
          MyCommon.Load_System_Info()
          sInstallationName = MyCommon.InstallationName
        End If

        MyCommon.Open_LogixXS()
        MyCommon.QueryStr = "select 1;"
        dt = MyCommon.LXS_Select
        dt.Dispose()

        If iMode = 2 Then
          MyCommon.Open_LogixWH()
          MyCommon.QueryStr = "select 1;"
          dt = MyCommon.LWH_Select
          dt.Dispose()
        End If
      ElseIf iMode = 99 Then
        Throw New ApplicationException("Sql Server Error Test")
      End If

      If sStoreId.Length > 0 Then
        If MyCommon.LRTadoConn.State <> ConnectionState.Open Then
          MyCommon.Open_LogixRT()
          MyCommon.Load_System_Info()
          sInstallationName = MyCommon.InstallationName
        End If
        sResponse = UpdateLocationHealth(sStoreId, 0, bTestLocation)
      End If
    Catch exApp As ApplicationException
      sResponse = BuildErrorXml(exApp.Message, ErrorType.SqlServer, False)
    Catch ex As Exception
      sResponse = BuildErrorXml(ex.Message, ErrorType.SqlServer)
    Finally
      If MyCommon.LWHadoConn.State <> ConnectionState.Closed Then
        MyCommon.Close_LogixWH()
      End If
      If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then
        MyCommon.Close_LogixXS()
      End If
      WriteDebug("GetStatus", DebugState.EndTime)
      If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then
        MyCommon.Close_LogixRT()
      End If
    End Try
    Return sResponse
  End Function

  Private Function GetLocationInfo(ByVal sStoreId As String) As String
    Dim sResponse As String = sOkStatus
    Dim i64LocationId As Int64 = 0
    Dim i32LocalServerId As Int32 = 0
    Dim bDistributeFiles As Boolean
    Dim bTestLocation As Boolean
    Dim sTestLocation As String

    Try
      sCurrentStoreId = sStoreId
      sCurrentMethod = "GetLocationInfo"
      sInputForLog = "*(Type=Input) (Method=GetLocationInfo)] - (StoreId='" & sStoreId & "')"
      MyCommon = New Copient.CommonInc
      MyCommon.AppName = sAppName
      sStoreId = sStoreId.Trim(" ")

      eDefaultErrorType = ErrorType.SqlServer
      If MyCommon.LRTadoConn.State <> ConnectionState.Open Then
        MyCommon.Open_LogixRT()
        MyCommon.Load_System_Info()
        sInstallationName = MyCommon.InstallationName
      End If
      eDefaultErrorType = ErrorType.General

      MyCommon.QueryStr = "dbo.pa_LogixServ_HealthUpdate"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@ExtLocationCode", SqlDbType.NVarChar, 20).Value = sStoreId
      MyCommon.LRTsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Direction = ParameterDirection.Output
      MyCommon.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Direction = ParameterDirection.Output
      MyCommon.LRTsp.Parameters.Add("@DistributeFiles", SqlDbType.Bit).Direction = ParameterDirection.Output
      MyCommon.LRTsp.Parameters.Add("@TestLocation", SqlDbType.Bit).Direction = ParameterDirection.Output
      MyCommon.LRTsp.ExecuteNonQuery()
      i64LocationId = MyCommon.LRTsp.Parameters("@LocationID").Value
      i32LocalServerId = MyCommon.LRTsp.Parameters("@LocalServerID").Value
      bDistributeFiles = MyCommon.LRTsp.Parameters("@DistributeFiles").Value
      bTestLocation = MyCommon.LRTsp.Parameters("@TestLocation").Value
      MyCommon.Close_LRTsp()
      If i64LocationId > 0 Then
        If i32LocalServerId > 0 Then
          If bTestLocation Then
            sTestLocation = "True"
          Else
            sTestLocation = "False"
          End If
          sResponse = "<Response ServerSerial=""" & i32LocalServerId & """ LocationID=""" & i64LocationId & _
                      """ TestLocation=""" & sTestLocation & """ />" & vbCrLf
        Else
          Throw New ApplicationException("Error generating new LocalServer for ExtLocationCode:" & sStoreId)
        End If
      Else
        Throw New ApplicationException("Location with ExtLocationCode='" & sStoreId & "' does not exist!")
      End If
    Catch exApp As ApplicationException
      sResponse = BuildErrorXml(exApp.Message, eDefaultErrorType, False)
    Catch ex As Exception
      sResponse = BuildErrorXml(ex.Message, eDefaultErrorType, True)
    Finally
      If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then
        MyCommon.Close_LogixRT()
      End If
    End Try
    Return sResponse
  End Function

  Private Function Mask(ByVal sCustomerId As String) As String
    Const sMask As String = "*****************"
    Dim sMaskedId As String

    If sCustomerId.Length > 4 Then
      sMaskedId = sMask.Substring(0, sCustomerId.Length - 4) & sCustomerId.Substring(sCustomerId.Length - 4, 4)
    ElseIf sCustomerId.Length >= 1 Then
      sMaskedId = sMask.Substring(0, 11 - sCustomerId.Length) & sCustomerId.Substring(1, sCustomerId.Length - 1)
    Else
      sMaskedId = sCustomerId
    End If
    Return sMaskedId
  End Function

  Private Sub SetGlobalCardTypeIds()

    If MyCommon.LRTadoConn.State <> ConnectionState.Open Then
      MyCommon.Open_LogixRT()
    End If

    If Not Integer.TryParse(MyCommon.Fetch_CM_SystemOption(iMemberIdOptionId), iIdTypeForMemberId) Then iIdTypeForMemberId = 0
    If Not Integer.TryParse(MyCommon.Fetch_CM_SystemOption(iPrimaryIdOptionId), iIdTypeForPrimaryId) Then iIdTypeForPrimaryId = -1
    If Not Integer.TryParse(MyCommon.Fetch_CM_SystemOption(iSecondaryIdOptionId), iIdTypeForSecondaryId) Then iIdTypeForSecondaryId = -1
    If Not Integer.TryParse(MyCommon.Fetch_CM_SystemOption(iAlternateIdOptionId), iIdTypeForAlternateId) Then iIdTypeForAlternateId = -1
    If Not Integer.TryParse(MyCommon.Fetch_CM_SystemOption(iEmployeeIdOptionId), iIdTypeForEmployeeId) Then iIdTypeForEmployeeId = -1

  End Sub

  Private Function GetTrxCardTypeId() As Integer
    Dim iTrxCardTypeId As Integer

    SetGlobalCardTypeIds()

    If iIdTypeForPrimaryId > 0 Then
      iTrxCardTypeId = iIdTypeForPrimaryId
    Else
      iTrxCardTypeId = iIdTypeForMemberId
    End If
    Return iTrxCardTypeId
  End Function

  Private Function GetAccountInfo(ByVal CustomerId As String, ByVal IdType As Integer, ByVal BusinessDate As Date, ByVal LaneId As String, ByVal TransactionNumber As Integer, ByVal sStoreId As String) As String
    ' the following formats may seem backwards as for as hh:mm go, but see usage in SQL query
    Const sDateFormatBusStart As String = "yyyy-MM-ddT23:59:59"
    Const sDateFormatBusEnd As String = "yyyy-MM-ddT00:00:00"


    Dim sXml As String = ""
    Dim dt As DataTable
    Dim dtAccountInfo As DataTable = Nothing
    Dim dtPromoVars As DataTable = Nothing
    Dim dtExtPromoVars As DataTable = Nothing
    Dim dtPendingPromoVars As DataTable = Nothing
    Dim dtPromos As DataTable = Nothing
    Dim dtStoredValues As DataTable = Nothing
    Dim dtPromosAnyAll As DataTable = Nothing
    Dim dtExtBalances As DataTable = Nothing
    Dim dtSupplementals As DataTable = Nothing
    Dim lCustomerPk As Long
    Dim lHhPk As Long
    Dim lUseCustomerPk As Long
    Dim iMaxRows As Integer
    Dim sMaskedId As String
    Dim sBusDate As String
    Dim sBusDateStart As String
    Dim sBusDateEnd As String
    Dim sResponse As String = sOkStatus
    Dim bNewCustomer As Boolean = False
    Dim objTemp As Object
    Dim sarrAltIdTableCol() As String
        Dim sAltIdTableCol As String
        Dim sValAltId As String
    Dim bUsedAltId As Boolean = False
    Dim bTestLocation As Boolean = False
    Dim bTestCustomer As Boolean = False
    Dim bCustomerDisabled As Boolean = False
    Dim bTestCardOptionEnabled As Boolean = False
    Dim sDebugMsg As String
    Dim sMemberLevelInfo As String = ""
    Dim sarrVerifierTableCol() As String
    Dim sVerifierTableCol As String
    Dim sVerifier As String = ""
    Dim sTenderPointIds As String
    Dim bLifetime As Boolean = False
    Dim sUrl As String
    Dim sUserName As String
    Dim sPassword As String
    Dim ExternalCA As Copient.ExternalCustomerAccounts

    Dim sAcsIdType As String
    Dim bCardFound As Boolean

    Dim iCustInfoCardType As Integer

    Dim sMemberId As String = ""
    Dim sPrimaryId As String = ""
    Dim sSecondaryId As String = ""
    Dim sAlternateId As String = ""
    Dim sCustInfoCardId As String
    Dim bReadOnly As Boolean = False
    
    MyCryptLib = New Copient.CryptLib
    Dim bFobEligibilityEnabled As Boolean = False
    Dim bNewFobCustomer As Boolean = False
    Dim oFetchTargetedListMode As FetchTargetedListMode = FetchTargetedListMode.Normal

    Try
      MyCommon = New Copient.CommonInc
      MyCommon.AppName = sAppName

      sCurrentStoreId = sStoreId
      sCurrentMethod = "GetAccountInfo"
      sBusDate = Format(BusinessDate, sDateFormat)
      sBusDateStart = Format(BusinessDate, sDateFormatBusStart)
      sBusDateEnd = Format(BusinessDate, sDateFormatBusEnd)

      sMaskedId = Mask(CustomerId)

      sInputForLog = "*(Type=Input) (Method=GetAccountInfo)] - (CustomerId='" & sMaskedId & "') (IdType='" & IdType.ToString & _
      "') (BusinessDate='" & sBusDate & "') (LaneId='" & LaneId & "') (TransactionNumber='" & _
      TransactionNumber.ToString & "') (sStoreId='" & sStoreId & "')"

      bFobEligibilityEnabled = IIf(MyCommon.Fetch_CM_SystemOption(142) = "1", True, False)

      If IdType > 999 Then
        IdType = IdType - 1000
        bReadOnly = True
        WriteDebug("GetAccount (Read Only)", DebugState.BeginTime)
      Else
        WriteDebug("GetAccount", DebugState.BeginTime)
      End If

      If IdType = 10 Then
        sXml = BuildErrorXml("This is a SQL Server error test", ErrorType.SqlServer, False)
        Exit Try
      ElseIf IdType = 11 Then
        sXml = BuildErrorXml("This is a HARD error test", ErrorType.General, False)
        Exit Try
      ElseIf IdType = 12 Then
        sXml = BuildErrorXml("This is a DB2 NOT FOUND error test", ErrorType.Db2ExtCustomerIdNotFound, False)
        Exit Try
      ElseIf IdType = 13 Then
        sXml = BuildErrorXml("This is a DB2 HARD/OFFLINE error test", ErrorType.Db2General, False)
        Exit Try
      End If
      eDefaultErrorType = ErrorType.SqlServer
      If MyCommon.LRTadoConn.State <> ConnectionState.Open Then
        MyCommon.Open_LogixRT()
        MyCommon.Load_System_Info()
        sInstallationName = MyCommon.InstallationName
      End If
      MyCommon.Open_LogixXS()
      eDefaultErrorType = ErrorType.General

      If MyCommon.IsIntegrationInstalled(Integrations.PREFERENCE_MANAGER) Then
        MyCommon.Open_PrefManRT()
        PrefLib = New Copient.AMSPreferenceLib(MyCommon)
        WriteDebug("Preference Manager loaded", DebugState.CurrentTime)
      End If

      UpdateLocationHealth(sStoreId, 1, bTestLocation)

      ' check if External customer Account is enabled
      If MyCommon.NZ(MyCommon.Fetch_CM_SystemOption(49), "") = "1" Then
        sCurrentMethod = "GetExternalCustomerAccount"
        sUrl = MyCommon.NZ(MyCommon.Fetch_CM_SystemOption(50), "")
        If sUrl <> "" Then
          sUserName = MyCommon.NZ(MyCommon.Fetch_CM_SystemOption(53), "")
          If sUserName <> "" Then
            sPassword = MyCommon.NZ(MyCommon.Fetch_CM_SystemOption(54), "")
          Else
            sPassword = ""
          End If
          WriteDebug("Get External Customer Information", DebugState.BeginTime)
          Try
            ExternalCA = New Copient.ExternalCustomerAccounts()
            ExternalCA.sUrl = sUrl
            ExternalCA.sUserName = sUserName
            ExternalCA.sPassword = sPassword
            sXml = ExternalCA.GetExternalCustomerAccount(CustomerId, IdType, BusinessDate, LaneId, TransactionNumber.ToString, sStoreId)
          Catch ex As Exception
            Throw ex
          Finally
            WriteDebug("Get External Customer Information", DebugState.EndTime)
          End Try
        Else
          sXml = BuildErrorXml("URL to External Customer Account web service is not defined in CM Settings", ErrorType.General, False)
        End If
        Exit Try
      Else
        sCurrentMethod = "GetAccountInfo"
      End If

      ' set global card types using Cm options in order to map ACS types to Logix types
      SetGlobalCardTypeIds()

      ' obtain CustomerPK for this Card
      Select Case IdType
        Case 0
          ' Membership Card (Barcode)
          bCardFound = GetCustomerPK(CustomerId, iIdTypeForMemberId, lCustomerPk, lHhPk, bTestCustomer)
          If Not bCardFound Then
            If Not MyCommon.AllowToProcessCustomerCard(CustomerId, iIdTypeForMemberId, cardValidationResp) Then
              sXml = BuildErrorXml(MyCommon.CardValidationResponseMessage(CustomerId, iIdTypeForMemberId, cardValidationResp), ErrorType.General, False)
              Exit Try
            End If
            bCardFound = AddCustomer(CustomerId, iIdTypeForMemberId, lCustomerPk, lHhPk, bTestCustomer)
            If bCardFound Then
              bNewCustomer = True
            Else
              sXml = ConvertMemberNotFoundToXml(CustomerId, IdType)
              Exit Try
            End If
          End If
          sAcsIdType = "ID"
          iCustInfoCardType = iIdTypeForMemberId
          sCustInfoCardId = CustomerId
          sMemberId = CustomerId
        Case 3
          ' external database lookup
          Dim sExtCardId As String
          eDefaultErrorType = ErrorType.Db2General

          sExtCardId = GetDb2ExtCustomerId(CustomerId, IdType, BusinessDate, LaneId, TransactionNumber, Int16.Parse(sStoreId))
          eDefaultErrorType = ErrorType.General
          sCurrentMethod = "GetAccountInfo"
          If sExtCardId = "" OrElse sExtCardId = "0" Then
            sXml = BuildErrorXml("ID not found!", ErrorType.Db2ExtCustomerIdNotFound, False)
            Exit Try
          Else
            bCardFound = GetCustomerPK(sExtCardId, 0, lCustomerPk, lHhPk, bTestCustomer)
            If Not bCardFound Then
              If Not MyCommon.AllowToProcessCustomerCard(sExtCardId, 0, cardValidationResp) Then
                sXml = BuildErrorXml(MyCommon.CardValidationResponseMessage(sExtCardId, 0, cardValidationResp), ErrorType.General, False)
                Exit Try
              End If
              bCardFound = AddCustomer(sExtCardId, 0, lCustomerPk, lHhPk, bTestCustomer)
              If bCardFound Then
                bNewCustomer = True
              Else
                sXml = ConvertMemberNotFoundToXml(CustomerId, 0)
                Exit Try
              End If
            End If
          End If
          sAcsIdType = "ExtID"
          iCustInfoCardType = 0
          sCustInfoCardId = sExtCardId
          sMemberId = sExtCardId
        Case 4
          ' Alternate ID, so must check if using card type or customer field
          If iIdTypeForAlternateId > 0 Then
            ' use card type for alternate ID
            bCardFound = GetCustomerPK(CustomerId, iIdTypeForAlternateId, lCustomerPk, lHhPk, bTestCustomer)
            If Not bCardFound Then
              sXml = ConvertMemberNotFoundToXml(CustomerId, IdType)
              Exit Try
            End If
            iCustInfoCardType = iIdTypeForAlternateId
            sCustInfoCardId = CustomerId
            sAlternateId = CustomerId
          Else
            ' use customer field for alternate ID
            sAltIdTableCol = MyCommon.Fetch_SystemOption(60)
            If sAltIdTableCol.Length > 0 Then
              sarrAltIdTableCol = sAltIdTableCol.Split(".")
              If sarrAltIdTableCol.Length = 2 Then

                                sValAltId = CustomerId

                                If sAltIdTableCol = "CustomerExt.PhoneAsEntered" Or sAltIdTableCol = "CustomerExt.PhoneDigitsOnly" Or sAltIdTableCol = "CustomerExt.email" Or sAltIdTableCol = "CustomerExt.DateOfBirth" Or sAltIdTableCol = "CustomerExt.DOB" Or sAltIdTableCol = "CustomerExt.MobilePhoneAsEntered" Or sAltIdTableCol = "CustomerExt.MobilePhoneDigitsOnly" Or sAltIdTableCol = "CustomerExt.TaxExemptID" Then
                                    sValAltId = MyCryptLib.SQL_StringEncrypt(CustomerId)
                                End If
                                MyCommon.QueryStr = "select Customers.CustomerPK, CustomerTypeID, HHPK, TestCard from Customers with (NoLock)" & _
                                    " left join CustomerExt with (NoLock) on CustomerExt.CustomerPK=Customers.CustomerPK" & _
                                    " where CustomerTypeID=0 and " & sAltIdTableCol & "='" & sValAltId & "';"
                dt = MyCommon.LXS_Select()
                If (dt.Rows.Count > 0) Then
                  sAlternateId = CustomerId
                  lCustomerPk = dt.Rows(0).Item("CustomerPK")
                  sMemberId = GetCardId(lCustomerPk, 0)
                  If sMemberId <> "" Then
                    lHhPk = dt.Rows(0).Item("HHPK")
                    bNewCustomer = False
                    bTestCustomer = MyCommon.NZ(dt.Rows(0).Item("TestCard"), False)
                    iCustInfoCardType = 0
                    sCustInfoCardId = sMemberId
                    bUsedAltId = True
                  Else
                    If bFobEligibilityEnabled And (iIdTypeForSecondaryId = 7) Then
                      sSecondaryId = GetCardId(lCustomerPk, 7)
                      If sSecondaryId <> "" Then
                        iCustInfoCardType = 7
                        sCustInfoCardId = sSecondaryId
                        lHhPk = dt.Rows(0).Item("HHPK")
                        bNewCustomer = False
                        bTestCustomer = MyCommon.NZ(dt.Rows(0).Item("TestCard"), False)
                        bUsedAltId = True
                      Else
                        sXml = ConvertMemberNotFoundToXml(CustomerId, IdType)
                        Exit Try
                      End If
                    Else
                      sXml = ConvertMemberNotFoundToXml(CustomerId, IdType)
                      Exit Try
                    End If
                  End If
                Else
                  sXml = ConvertMemberNotFoundToXml(CustomerId, IdType)
                  Exit Try
                End If
              Else
                sXml = BuildErrorXml("Invalid Alternate ID defined in options file!", ErrorType.General, False)
                Exit Try
              End If
            Else
              sXml = BuildErrorXml("Alternate ID is NOT defined in options file!", ErrorType.General, False)
              Exit Try
            End If
          End If
          sAcsIdType = "Alt ID"
        Case 6
          ' Secondary Card 
          If iIdTypeForSecondaryId > 0 Then
            bCardFound = GetCustomerPK(CustomerId, iIdTypeForSecondaryId, lCustomerPk, lHhPk, bTestCustomer)
            If Not bCardFound Then
              If bFobEligibilityEnabled And (iIdTypeForSecondaryId = 7) Then
                ' customer not found for Fob, so create new FOB only customer
                bCardFound = AddCustomer(CustomerId, iIdTypeForSecondaryId, lCustomerPk, lHhPk, bTestCustomer)
                If bCardFound Then
                  bNewFobCustomer = True
                  WriteDebug("Added new FOB only customer", DebugState.CurrentTime)
                Else
                  sXml = ConvertMemberNotFoundToXml(CustomerId, 0)
                  Exit Try
                End If
              Else
                ' customer not found for SecondaryID
                sXml = ConvertMemberNotFoundToXml(CustomerId, IdType)
                Exit Try
              End If
            End If
            sAcsIdType = "2nd ID"
            iCustInfoCardType = iIdTypeForSecondaryId
            sCustInfoCardId = CustomerId
            sSecondaryId = CustomerId
          Else
            sXml = BuildErrorXml("Secondary ID is NOT defined in options file!", ErrorType.General, False)
            Exit Try
          End If
        Case 8
          bCardFound = GetCustomerPK(CustomerId, iIdTypeForEmployeeId, lCustomerPk, lHhPk, bTestCustomer)
          If Not bCardFound Then
            sXml = ConvertMemberNotFoundToXml(CustomerId, IdType)
            Exit Try
          End If
          iCustInfoCardType = iIdTypeForEmployeeId
          sCustInfoCardId = CustomerId
          sAcsIdType = "Employee ID"
          sMemberId = CustomerId
        Case Else
          sXml = BuildErrorXml("Invalid ACS Customer ID type: " & IdType, ErrorType.General, False)
          Exit Try
      End Select

      If lHhPk = 0 Then
        lUseCustomerPk = lCustomerPk
        sDebugMsg = "got CustomerPK (" & lCustomerPk & ") via " & sAcsIdType & " (" & sMaskedId & ")"
      Else
        lUseCustomerPk = lHhPk
        sDebugMsg = "got CustomerPK (" & lCustomerPk & ") via " & sAcsIdType & " (" & sMaskedId & "), using HouseholdPK (" & lHhPk & ")"
      End If
      WriteDebug(sDebugMsg, DebugState.CurrentTime)

      ' Query for customer Info
      sCustInfoCardId = MyCommon.Pad_ExtCardID(sCustInfoCardId, iCustInfoCardType)
      MyCommon.QueryStr = "dbo.pa_LogixServ_CustInfo"
      MyCommon.Open_LXSsp()
            MyCommon.LXSsp.Parameters.Add("@ExtCardID", SqlDbType.NVarChar, 400).Value = MyCryptLib.SQL_StringEncrypt(sCustInfoCardId)
      MyCommon.LXSsp.Parameters.Add("@CardTypeID", SqlDbType.Int).Value = iCustInfoCardType
      dtAccountInfo = MyCommon.LXSsp_select
      MyCommon.Close_LXSsp()
      dtAccountInfo.TableName = "AccountInfo"
      WriteDebug("got CustomerInfo", DebugState.CurrentTime)

      if (MyCommon.NZ(MyCommon.Fetch_CM_SystemOption(141), "0") = "1") Then
        MyCommon.QueryStr = "select CustomerSupplemental.Value, CustomerSupplementalFields.ExtFieldID from CustomerSupplemental with (NoLock)" & _
                            " left join CustomerSupplementalFields on CustomerSupplemental.FieldID = CustomerSupplementalFields.FieldID" & _
                            " where CustomerPK = " & lUseCustomerPk
        dtSupplementals = MyCommon.LXS_Select()
      End If


      ' Get IDs for cards other than one presented
      If sPrimaryId.Length = 0 Then
        If iIdTypeForPrimaryId > -1 Then
          If iIdTypeForPrimaryId = iIdTypeForMemberId Then
            sPrimaryId = ""
          Else
            sPrimaryId = GetCardId(lCustomerPk, iIdTypeForPrimaryId)
            If sPrimaryId = "" Then
              ' Primary ID not required in FOB scenario
              If Not bFobEligibilityEnabled Then
                sXml = BuildErrorXml("Primary ID not found for " & sAcsIdType & ": " & CustomerId, ErrorType.General, False)
                Exit Try
              End If
            End If
          End If
        Else
          sPrimaryId = ""
        End If
      End If

      If sSecondaryId.Length = 0 Then
        If iIdTypeForSecondaryId > -1 Then
          If iIdTypeForSecondaryId = iIdTypeForMemberId Then
            sSecondaryId = ""
          Else
            sSecondaryId = GetCardId(lCustomerPk, iIdTypeForSecondaryId)
          End If
        Else
          sSecondaryId = ""
        End If
      End If

      If sMemberId.Length = 0 Then
        If iIdTypeForMemberId > -1 Then
          sMemberId = GetCardId(lCustomerPk, iIdTypeForMemberId)
        Else
          sMemberId = ""
        End If
      End If

      If sAlternateId.Length = 0 Then
        If iIdTypeForAlternateId > 0 Then
          sAlternateId = GetCardId(lCustomerPk, iIdTypeForAlternateId)
        Else
          sAltIdTableCol = MyCommon.Fetch_SystemOption(60)
          If sAltIdTableCol.Length > 0 Then
            sarrAltIdTableCol = sAltIdTableCol.Split(".")
            If sarrAltIdTableCol.Length = 2 Then
              MyCommon.QueryStr = "select " & sAltIdTableCol & " as AltId from Customers with (NoLock)" & _
                                  " left join CustomerExt with (NoLock) on CustomerExt.CustomerPK=Customers.CustomerPK" & _
                                  " where CustomerTypeID=0 and Customers.CustomerPK=" & lCustomerPk & ";"
              dt = MyCommon.LXS_Select()
              If (dt.Rows.Count > 0) Then
                sAlternateId = MyCommon.NZ(dt.Rows(0).Item("AltId"), "")
                sAlternateId = sAlternateId.Trim()
              End If
            Else
              WriteLog("Invalid Alternate ID defined in options file!", MessageType.Warning)
            End If
          End If
        End If
      End If

      ' Get the current value for Alternate ID Verifier to be returned as part of customer info
      If sVerifier.Length = 0 Then
        sVerifierTableCol = MyCommon.Fetch_SystemOption(61)
        If sVerifierTableCol.Length > 0 Then
          sarrVerifierTableCol = sVerifierTableCol.Split(".")
          If sarrVerifierTableCol.Length = 2 Then
            MyCommon.QueryStr = "select " & sVerifierTableCol & " as Verifier from Customers with (NoLock)" & _
                                " left join CustomerExt with (NoLock) on CustomerExt.CustomerPK=Customers.CustomerPK" & _
                                " where CustomerTypeID=0 and Customers.CustomerPK=" & lCustomerPk & ";"
            dt = MyCommon.LXS_Select()
            If (dt.Rows.Count > 0) Then
              sVerifier = MyCommon.NZ(dt.Rows(0).Item("Verifier"), "")
              sVerifier = sVerifier.Trim()
              If sarrVerifierTableCol(1).ToLower = "password" Then
                sVerifier = MyCryptLib.SQL_StringDecrypt(sVerifier)
              End If
            End If
          Else
            WriteLog("Invalid Alternate ID Verifier defined in options file!", MessageType.Warning)
          End If
        End If
      End If

      If bFobEligibilityEnabled Then
        ' Special FOB stuff
        Select Case IdType
          Case 6
            ' Presented FOB
            If bNewFobCustomer Then
              ' Fob nor found, but created customer on fly
              oFetchTargetedListMode = FetchTargetedListMode.FobOnly
              ' return FOB for ID value in GetAccount response XML
              sPrimaryId = sSecondaryId
            Else
              If sMemberId.Length > 0 Then
                ' found mPerks
                oFetchTargetedListMode = FetchTargetedListMode.Normal
              Else
                ' mPerks not found
                oFetchTargetedListMode = FetchTargetedListMode.FobTargeted
                ' return FOB for ID value in GetAccount response XML
                sPrimaryId = sSecondaryId
              End If
            End If
          Case 4
            ' Presented Alternate ID
            If sMemberId.Length > 0 Then
              ' found mPerks
              oFetchTargetedListMode = FetchTargetedListMode.Normal
            Else
              ' mPerks not found
              If sSecondaryId.Length > 0 Then
                ' found FOB
                oFetchTargetedListMode = FetchTargetedListMode.FobTargeted
                ' return FOB for ID value in GetAccount response XML
                sPrimaryId = sSecondaryId
              Else
                ' FOB not found
                sXml = ConvertMemberNotFoundToXml(CustomerId, IdType)
                Exit Try
              End If
            End If
          Case Else
            oFetchTargetedListMode = FetchTargetedListMode.Normal
        End Select
      End If

      ' Determine if customer should be disabled or not
      ' Test Cards enabled?
      If MyCommon.Fetch_SystemOption(88) = "1" Then
        bTestCardOptionEnabled = True
        If (bTestLocation And bTestCustomer) Or ((Not bTestLocation) And (Not bTestCustomer)) Then
          bCustomerDisabled = False
        Else
          bCustomerDisabled = True
        End If
      Else
        bTestCardOptionEnabled = False
        bCustomerDisabled = False
      End If

      If bCustomerDisabled Then
        Dim sTemp As String = String.Format("Test card and test location mismatch: CustomerPK '{0}' Store '{1}' TermID '{2}' TranID '{3}'", lCustomerPk, sStoreId, LaneId, TransactionNumber)
        WriteLog(sTemp, MessageType.Warning)
      Else
        ' Query Promo Variables
        sTenderPointIds = MyCommon.NZ(MyCommon.Fetch_CM_SystemOption(37), "")
        If sTenderPointIds = "" Then
          bLifetime = False
          MyCommon.QueryStr = "dbo.pa_LogixServ_CustPromoVars"
        Else
          bLifetime = True
          ' this procedure returns an additional "Lifetime" column which is
          ' null for all promo vars except points which have a corresponding
          ' entry in the CM_Points_Lifetime table
          MyCommon.QueryStr = "dbo.pa_LogixServ_CustPromoVars_Lifetime"
        End If
        MyCommon.Open_LXSsp()
        MyCommon.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = lUseCustomerPk
        dtPromoVars = MyCommon.LXSsp_select
        MyCommon.Close_LXSsp()
        dtPromoVars.TableName = "PromoVars"
        WriteDebug("got CustomerPromoVars", DebugState.CurrentTime)

        ' Get all External Points Programs balances
        If MyCommon.Fetch_SystemOption(80) = "1" Then
          WriteDebug("Get External Points", DebugState.BeginTime)
          Try
            Dim ExternalPP As Copient.ExternalRewards
            sDb2Connection = MyCommon.Fetch_CM_SystemOption(iDb2Connection)
            tempstr = sDb2Connection
            upos = InStr(tempstr, "UID=", CompareMethod.Text)
            ppos = InStr(tempstr, ";PWD=", CompareMethod.Text)
            pend = InStr(tempstr, ";host", CompareMethod.Text)
            euser = tempstr.Substring(upos+3, ppos-upos-4)
            epwd = tempstr.Substring(ppos+4, pend-ppos-5)
            user =  MyCryptLib.SQL_StringDecrypt(euser)
            pwd =  MyCryptLib.SQL_StringDecrypt(epwd)
            tempstr = tempstr.Replace(euser, user)
            sDb2Connection = tempstr.Replace(epwd, pwd)
            ExternalPP = New Copient.ExternalRewards("", "", "", sDb2Connection)
            ' add promo variables for external points programs
            ExternalPP.appendExtPromoVarBalances(sMemberId, False, dtPromoVars, MyCommon)
          Catch
            Throw
          Finally
            WriteDebug("Get External Points", DebugState.EndTime)
          End Try
        End If

        ' If Pending Points is enabled, return pending point values
        If MyCommon.Fetch_SystemOption(251) = "1" Then
          MyCommon.QueryStr = "select ProgramID, ApplyEarnedPendingPoints from PointsPrograms with (nolock) where Deleted=0;"
          dt = MyCommon.LRT_Select
          MyCommon.QueryStr = "dbo.pa_LogixServ_CustPromoVars_Pending"
          MyCommon.Open_LXSsp()
          MyCommon.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = lUseCustomerPk
          MyCommon.LXSsp.Parameters.Add("@PendingFlagTable", SqlDbType.Structured).Value = dt
          dtPendingPromoVars = MyCommon.LXSsp_select
          MyCommon.Close_LXSsp()

          ' add pending promo variables to promo variable table
          If (Not dtPendingPromoVars Is Nothing) AndAlso (dtPendingPromoVars.Rows.Count > 0) Then
            dtPromoVars.Columns.Add("CartID", System.Type.GetType(dtPendingPromoVars.Columns("CartID").DataType.ToString))
            For Each dr In dtPendingPromoVars.Rows
               dtPromoVars.Rows.Add(dr.Item("ID"), dr.Item("Val"), dr.Item("ExternalID"), dr.Item("CartID"))
            Next
          End If
          WriteDebug("got Pending CustomerPromoVars", DebugState.CurrentTime)
        End If

        ' Query Stored Values
        objTemp = MyCommon.Fetch_CM_SystemOption(16)
        If Not (Integer.TryParse(objTemp.ToString, iMaxRows)) Then iMaxRows = 5
        If iMaxRows < 5 Then
          iMaxRows = 5
        ElseIf iMaxRows > 100 Then
          iMaxRows = 5
        End If
        MyCommon.QueryStr = "dbo.pa_LogixServ_CustStoredValues"
        MyCommon.Open_LXSsp()
        MyCommon.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = lUseCustomerPk
        MyCommon.LXSsp.Parameters.Add("@MaxRows", SqlDbType.BigInt).Value = iMaxRows
        dtStoredValues = MyCommon.LXSsp_select
        MyCommon.Close_LXSsp()
        dtStoredValues.TableName = "StoredValues"
        WriteDebug("got CustomerStoredValues", DebugState.CurrentTime)

        ' Member
        Dim bAutoHouseholdCustGrpEnabled As Boolean = MyCommon.Fetch_CM_SystemOption(iAutoHouseholdCustGrpOptionId)

        ' group list is returned as XML now
        Dim customer_groups As XmlDocument = getCustomerGroups(lCustomerPk, lHhPk, bAutoHouseholdCustGrpEnabled)

        WriteDebug("got CustomerGroupList (XML)", DebugState.CurrentTime)

        If (bAutoHouseholdCustGrpEnabled) Then
          sMemberLevelInfo = GetMembershipLevelForAutoHouseHold(customer_groups)
        Else
          sMemberLevelInfo = GetMembershipLevel(lCustomerPk)
      End If

        If bFobEligibilityEnabled Then
          Select Case oFetchTargetedListMode
            Case FetchTargetedListMode.FobOnly
              ' New FOB card
              dtPromos = getPromotionsFobOnly(sBusDateStart, sBusDateEnd, lCustomerPk)
              WriteDebug("got CustomerOffers (FOB only)", DebugState.CurrentTime)
            Case FetchTargetedListMode.FobTargeted
              ' Fob found, but no mPerks
              Using xnr As New XmlNodeReader(customer_groups)
                dtPromos = getPromotionsFobTargeted(xnr, sBusDateStart, sBusDateEnd, lCustomerPk)
              End Using
              WriteDebug("got CustomerOffers (FOB targeted - XML)", DebugState.CurrentTime)
            Case Else
              ' mPerks found, so use normal
                            Dim xnr1 As New XmlNodeReader(customer_groups)
                            dtPromos = getPromotions(xnr1, bNewCustomer, sBusDateStart, sBusDateEnd, bTestCardOptionEnabled And bTestLocation, lCustomerPk)

                            WriteDebug("got CustomerOffers (XML)", DebugState.CurrentTime)
                    End Select
        Else
          Using xnr As New XmlNodeReader(customer_groups)
            dtPromos = getPromotions(xnr, bNewCustomer, sBusDateStart, sBusDateEnd, bTestCardOptionEnabled And bTestLocation, lCustomerPk)
          End Using
          WriteDebug("got CustomerOffers (XML)", DebugState.CurrentTime)
        End If

      End If

      sXml = ConvertAccountInfoToXml(dtAccountInfo, dtPromoVars, dtPromos, dtStoredValues, lUseCustomerPk, sStoreId, LaneId, TransactionNumber.ToString, sMemberId, sPrimaryId, sSecondaryId, sAlternateId, sVerifier, sMemberLevelInfo, bUsedAltId, bCustomerDisabled, bLifetime, bReadOnly, lCustomerPk, dtSupplementals)

      ValidateXmlLength(sXml, CustomerId, BusinessDate, LaneId, TransactionNumber, sStoreId)

      If bReadOnly Then
        MyCommon.QueryStr = "update Customers with (RowLock) set UpdateCount=UpdateCount+1 where CustomerPK=" & lUseCustomerPk & ";"
        MyCommon.LXS_Execute()
      End If
    Catch exApp As ApplicationException
      sXml = BuildErrorXml(exApp.Message, ErrorType.Db2General, False)
    Catch ex As Exception
      sXml = BuildErrorXml(ex.Message, eDefaultErrorType)
    Finally
      If MyCommon.PMRTadoConn.State <> ConnectionState.Closed Then
        MyCommon.Close_PrefManRT()
      End If
      If MyCommon.LWHadoConn.State <> ConnectionState.Closed Then
        MyCommon.Close_LogixWH()
      End If
      If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then
        MyCommon.Close_LogixXS()
      End If
      WriteDebug("GetAccount", DebugState.EndTime)
      If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then
        MyCommon.Close_LogixRT()
      End If
    End Try
    Return sXml
  End Function


  Private Sub ValidateXmlLength(ByVal sXml As String,
                                  ByVal CustomerId As String,
                                  ByVal BusinessDate As Date, ByVal LaneId As String,
                                  ByVal TransactionNumber As Integer, ByVal sStoreId As String)
    Try
      Dim maxLength As Integer
      Dim sSeparator As String = vbCrLf
      maxLength = Convert.ToInt32(MyCommon.NZ(MyCommon.Fetch_CM_SystemOption(128), "0"))
      If maxLength > 0 And sXml.Length > maxLength Then
        Dim sWarning As String = "GetAcount XML contains " & sXml.Length & " characters which exceeds the warning length of " & maxLength & "!"
        WriteDebug(sWarning, DebugState.CurrentTime)

        sWarning &= String.Format("{5}Date : {0:d} {5}Card number : {1} {5}Transaction Number : {2} {5}Store Number : {3} {5}Lane ID : {4}",
                                  BusinessDate, CustomerId, TransactionNumber, sStoreId, LaneId, sSeparator)


        'email this ErrorMessage
        Dim EmailAddresses As String = MyCommon.Get_Error_Emails(1)

        Dim SystemEmailAddress As String = MyCommon.Fetch_SystemOption(3)
        Dim InstallationName As String = MyCommon.Fetch_SystemOption(2)
        MyCommon.Queue_Email(EmailAddresses, SystemEmailAddress, String.Format("{0} WARNING! - {1}", sAppName, sInstallationName), sWarning)

      End If
    Catch ex As Exception

    End Try
  End Sub


  Private Function GetCustomerPK(ByVal sCardId As String, ByVal iCardTypeId As Integer, ByRef lCustomerPK As Long, ByRef lHhPk As Long, ByRef bTestCustomer As Boolean) As Boolean
    Dim bStatus As Boolean
    Dim sId As String
    Dim dt As DataTable

    sId = sCardId.Trim
    sId = MyCommon.Pad_ExtCardID(sId, iCardTypeId)
        MyCryptLib = New Copient.CryptLib
        MyCommon.QueryStr = "dbo.pa_LogixServ_FetchCustPK"
    MyCommon.Open_LXSsp()
        MyCommon.LXSsp.Parameters.Add("@ExtCardID", SqlDbType.NVarChar, 400).Value = MyCryptLib.SQL_StringEncrypt(sId)
    MyCommon.LXSsp.Parameters.Add("@CardTypeID", SqlDbType.Int).Value = iCardTypeId
    dt = MyCommon.LXSsp_select()
    MyCommon.Close_LXSsp()
    If dt.Rows.Count > 0 Then
      lCustomerPK = dt.Rows(0).Item("CustomerPK")
      lHhPk = dt.Rows(0).Item("HHPK")
      bTestCustomer = MyCommon.NZ(dt.Rows(0).Item("TestCard"), False)
      bStatus = True
    Else
      lCustomerPK = 0
      lHhPk = 0
      bTestCustomer = False
      bStatus = False
    End If

    Return bStatus
  End Function

  Private Function GetCustomerPK_CardIdOnly(ByVal sCardId As String, ByRef iCardTypeId As Integer, ByRef lCustomerPK As Long, ByRef lHhPk As Long, ByRef bTestCustomer As Boolean) As Boolean
    Dim bStatus As Boolean
    Dim sId As String
    Dim dt As DataTable

    sId = sCardId.Trim
    sId = MyCommon.Pad_ExtCardID(sId, iCardTypeId)

    MyCommon.QueryStr = "dbo.pa_LogixServ_FetchCustPK_CardIdOnly"
    MyCommon.Open_LXSsp()
        MyCommon.LXSsp.Parameters.Add("@ExtCardID", SqlDbType.NVarChar, 400).Value = MyCryptLib.SQL_StringEncrypt(sId)
    dt = MyCommon.LXSsp_select()
    MyCommon.Close_LXSsp()
    If dt.Rows.Count > 0 Then
      lCustomerPK = dt.Rows(0).Item("CustomerPK")
      lHhPk = dt.Rows(0).Item("HHPK")
      bTestCustomer = MyCommon.NZ(dt.Rows(0).Item("TestCard"), False)
      iCardTypeId = dt.Rows(0).Item("CardTypeID")
      bStatus = True
    Else
      lCustomerPK = 0
      lHhPk = 0
      bTestCustomer = False
      iCardTypeId = -1
      bStatus = False
    End If

    Return bStatus
  End Function

  Private Function AddCustomer(ByVal sCardId As String, ByVal iCardType As Integer, ByRef lCustomerPK As Long, ByRef lHhPk As Long, ByRef bTestCustomer As Boolean) As Boolean
    Dim bStatus As Boolean
    Dim sId As String

    ' check if should disable add member on fly
    If MyCommon.Fetch_CM_SystemOption(65) = "1" Then
      ' disabled
      lCustomerPK = 0
      bStatus = False
    Else
      ' Add New Member on the fly
      sId = sCardId.Trim
      sId = MyCommon.Pad_ExtCardID(sId, iCardType)
      sDefaultFirstName = MyCommon.Fetch_CM_SystemOption(20)
      sDefaultLastName = MyCommon.Fetch_CM_SystemOption(21)

      MyCommon.QueryStr = "dbo.pa_ServiceCustomers_Insert"
      MyCommon.Open_LXSsp()
            MyCommon.LXSsp.Parameters.Add("@ExtCardID", SqlDbType.NVarChar, 400).Value = MyCryptLib.SQL_StringEncrypt(sId)
      MyCommon.LXSsp.Parameters.Add("@CardTypeID", SqlDbType.Int).Value = iCardType
      MyCommon.LXSsp.Parameters.Add("@CustomerTypeID", SqlDbType.Int).Value = 0
      MyCommon.LXSsp.Parameters.Add("@FirstName", SqlDbType.NVarChar, 50).Value = sDefaultFirstName
      MyCommon.LXSsp.Parameters.Add("@LastName", SqlDbType.NVarChar, 50).Value = sDefaultLastName
      MyCommon.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Direction = ParameterDirection.Output
      MyCommon.LXSsp.ExecuteNonQuery()
      lCustomerPK = MyCommon.LXSsp.Parameters("@CustomerPK").Value
      MyCommon.Close_LXSsp()
      bStatus = True
    End If
    lHhPk = 0
    bTestCustomer = False

    Return bStatus
  End Function

  Private Function GetCardId(ByVal lCustomerPK As Long, ByVal iCardTypeId As Integer) As String
    Dim sCardId As String = ""
    Dim dt As DataTable

    MyCommon.QueryStr = "dbo.pa_LogixServ_FetchCardID"
    MyCommon.Open_LXSsp()
    MyCommon.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = lCustomerPK
    MyCommon.LXSsp.Parameters.Add("@CardTypeID", SqlDbType.Int).Value = iCardTypeId
    dt = MyCommon.LXSsp_select()
    MyCommon.Close_LXSsp()
    If dt.Rows.Count > 0 Then
      sCardId = MyCommon.NZ(dt.Rows(0).Item(0), "")
    End If

    Return sCardId
  End Function



  Private Function GetCouponInfo(ByVal CouponId As String, ByVal IdType As Integer, ByVal BusinessDate As Date, ByVal LaneId As String, ByVal TransactionNumber As Integer, ByVal sStoreId As String) As String
    ' the following formats may seem backwards as for as hh:mm go, but see usage in SQL query
    Const sDateFormatBusStart As String = "yyyy-MM-ddT23:59:59"
    Const sDateFormatBusEnd As String = "yyyy-MM-ddT00:00:00"

    Dim sXml As String = ""
    Dim dt As DataTable
    Dim sBusDate As String
    Dim sBusDateStart As String
    Dim sBusDateEnd As String
    Dim sResponse As String = sOkStatus
    Dim sCouponId As String
    Dim bTestLocation As Boolean = False

    Try
      sCurrentStoreId = sStoreId
      sCurrentMethod = "GetCouponInfo"
      sBusDate = Format(BusinessDate, sDateFormat)
      sBusDateStart = Format(BusinessDate, sDateFormatBusStart)
      sBusDateEnd = Format(BusinessDate, sDateFormatBusEnd)


      sInputForLog = "*(Type=Input) (Method=GetCouponInfo)] - (CouponId='" & CouponId & "') (IdType='" & IdType.ToString & _
      "') (BusinessDate='" & sBusDate & "') (LaneId='" & LaneId & "') (TransactionNumber='" & _
      TransactionNumber.ToString & "') (sStoreId='" & sStoreId & "')"

      MyCommon = New Copient.CommonInc
      MyCommon.AppName = sAppName

      WriteDebug("GetCouponInfo", DebugState.BeginTime)

      eDefaultErrorType = ErrorType.SqlServer
      If MyCommon.LRTadoConn.State <> ConnectionState.Open Then
        MyCommon.Open_LogixRT()
        MyCommon.Load_System_Info()
        sInstallationName = MyCommon.InstallationName
      End If
      MyCommon.Open_LogixXS()
      eDefaultErrorType = ErrorType.General

      UpdateLocationHealth(sStoreId, 1, bTestLocation)
      sCurrentMethod = "GetCouponInfo"

      If IdType <> 999 Then
        sXml = BuildErrorXml("Invalid ID type: " & IdType, ErrorType.General, False)
        Exit Try
      End If
      sCouponId = CouponId.Trim

      ' Query Stored Values
      MyCommon.QueryStr = "select StoredValueId,LocalID,ServerSerial,SVProgramID,QtyEarned,QtyUsed," & _
                          "Value,EarnedDate,ExpireDate,ExternalID,StatusFlag,TotalValueEarned,RedeemedValue," & _
                          "BreakageValue from StoredValue with (NoLock) where CustomerPk=0 and ExternalId='" & sCouponId & "';"
      dt = MyCommon.LXS_Select
      If dt.Rows.Count = 0 Then
        ' Query Stored Values
        MyCommon.QueryStr = "select top 1 0 as StoredValueId,LocalID,ServerSerial,SVProgramID,QtyEarned,QtyUsed," & _
                            "Value,EarnedDate,ExpireDate,ExternalID,StatusFlag,TotalValueEarned,RedeemedValue," & _
                            "BreakageValue from SVHistory with (NoLock) where CustomerPk=0 and ExternalId='" & sCouponId & _
                            "' order by LastUpdate desc;"
        dt = MyCommon.LXS_Select
      End If
      If dt.Rows.Count > 0 Then
        sXml = ConvertCouponInfoToXml(dt, sStoreId, LaneId, TransactionNumber.ToString)
      Else
        sXml = ConvertMemberNotFoundToXml(sCouponId, IdType)
      End If
    Catch exApp As ApplicationException
      sXml = BuildErrorXml(exApp.Message, ErrorType.Db2General, False)
    Catch ex As Exception
      sXml = BuildErrorXml(ex.Message, eDefaultErrorType)
    Finally
      If MyCommon.LWHadoConn.State <> ConnectionState.Closed Then
        MyCommon.Close_LogixWH()
      End If
      If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then
        MyCommon.Close_LogixXS()
      End If
      WriteDebug("GetCouponInfo", DebugState.EndTime)
      If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then
        MyCommon.Close_LogixRT()
      End If
    End Try
    Return sXml
  End Function

  Private Function GetDb2ExtCustomerId(ByVal CustomerId As String, ByVal IdType As Integer, ByVal BusinessDate As Date, ByVal LaneId As String, ByVal TransactionNumber As Integer, ByVal StoreId As Int16) As String
    Dim sExtCustomerId As String = ""
    Dim iStatus As Integer
    Dim sErrorMsg As String
    Dim sMaskedId As String
    Const sMask As String = "00000000000000000"

    ' ODBC to DB2
    Dim oCn As OdbcConnection = Nothing
    Dim oCmd As OdbcCommand
    
    MyCryptLib = New Copient.CryptLib

    sCurrentMethod = "GetDb2ExtCustomerId"

    If LaneId.Length > 3 Then
      LaneId = LaneId.Substring(0, 3)
    End If

    If CustomerId.Length > 10 Then
      sMaskedId = CustomerId.Substring(0, 6) & sMask.Substring(0, CustomerId.Length - 10) & CustomerId.Substring(CustomerId.Length - 4, 4)
    Else
      sMaskedId = CustomerId
    End If

    WriteDebug("GetDb2ExtCustomerId", DebugState.BeginTime)
    Try
      If sDb2Connection.Length = 0 Then
        bDb2TestMode = False
        sDb2Connection = MyCommon.Fetch_CM_SystemOption(iDb2Connection)
        tempstr = sDb2Connection
        upos = InStr(tempstr, "UID=", CompareMethod.Text)
        ppos = InStr(tempstr, ";PWD=", CompareMethod.Text)
        pend = InStr(tempstr, ";host", CompareMethod.Text)
        euser = tempstr.Substring(upos+3, ppos-upos-4)
        epwd = tempstr.Substring(ppos+4, pend-ppos-5)
        user =  MyCryptLib.SQL_StringDecrypt(euser)
        pwd =  MyCryptLib.SQL_StringDecrypt(epwd)
        tempstr = tempstr.Replace(euser, user)
        sDb2Connection = tempstr.Replace(epwd, pwd)
        If sDb2Connection.Length = 0 Then
          sDb2Connection = sDefaultDb2Connection
        End If
        If sDb2Connection.IndexOf("DEVDB2NW") > -1 Then
          bDb2TestMode = True
        End If
      End If

      oCn = New OdbcConnection(sDb2Connection)
      oCn.Open()

      If bDb2TestMode Then
        ' SqlServer version (Copient test)
        oCmd = New OdbcCommand("{call DEVPROC.EFTF065 (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)}", oCn)
      Else
        ' DB2 version (Meijer production)
        oCmd = New OdbcCommand("{call EFTF065 (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)}", oCn)
      End If

      oCmd.Parameters.Add("MSKD_GUEST_PYMT_ID", OdbcType.Char, 26)
      oCmd.Parameters.Item("MSKD_GUEST_PYMT_ID").Value = sMaskedId

      oCmd.Parameters.Add("GUEST_PYMT_ID", OdbcType.Char, 26)
      oCmd.Parameters.Item("GUEST_PYMT_ID").Value = CustomerId

      oCmd.Parameters.Add("EFT_TN_BUS_DT", OdbcType.Char, 10)
      oCmd.Parameters.Item("EFT_TN_BUS_DT").Value = BusinessDate.ToString("yyyy-MM-dd")

      oCmd.Parameters.Add("UT_ID", OdbcType.SmallInt)
      oCmd.Parameters.Item("UT_ID").DbType = DbType.Int16
      oCmd.Parameters.Item("UT_ID").Value = StoreId

      oCmd.Parameters.Add("TL_ID", OdbcType.Char, 3)
      oCmd.Parameters.Item("TL_ID").Value = LaneId

      oCmd.Parameters.Add("TN_NBR_ID", OdbcType.Int)
      oCmd.Parameters.Item("TN_NBR_ID").Value = TransactionNumber

      oCmd.Parameters.Add("GST_TN_ID_TYP_ID", OdbcType.Char, 1)
      oCmd.Parameters.Item("GST_TN_ID_TYP_ID").Value = IdType.ToString.Substring(0, 1)

      oCmd.Parameters.Add("CALLING_PRC_ID", OdbcType.Char, 11)
      oCmd.Parameters.Item("CALLING_PRC_ID").Value = "Logix"

      oCmd.Parameters.Add("GUEST_HSHLD_ID", OdbcType.Double)
      oCmd.Parameters.Item("GUEST_HSHLD_ID").Precision = 15
      oCmd.Parameters.Item("GUEST_HSHLD_ID").Direction = ParameterDirection.Output

      oCmd.Parameters.Add("ERR_CODE", OdbcType.Int)
      oCmd.Parameters.Item("ERR_CODE").Direction = ParameterDirection.Output

      oCmd.Parameters.Add("ERR_MSG", OdbcType.Char, 80)
      oCmd.Parameters.Item("ERR_MSG").Direction = ParameterDirection.Output

      oCmd.ExecuteNonQuery()

      iStatus = oCmd.Parameters.Item("ERR_CODE").Value
      If iStatus > 0 Then
        sErrorMsg = oCmd.Parameters.Item("ERR_MSG").Value
        Throw New ApplicationException("Db2 returned error: " & iStatus & " - " & sErrorMsg)
        sExtCustomerId = ""
      Else
        sExtCustomerId = oCmd.Parameters.Item("GUEST_HSHLD_ID").Value
      End If
    Catch exApp As ApplicationException
      Throw
    Catch ex As Exception
      Throw
    Finally
      If Not oCn Is Nothing Then
        oCn.Close()
      End If
      WriteDebug("GetDb2ExtCustomerId", DebugState.EndTime)
    End Try

    Return sExtCustomerId
  End Function

  Private Function ConvertAccountInfoToXml(ByRef dtAccountInfo As DataTable, ByRef dtPromoVars As DataTable, ByRef dtPromos As DataTable, _
                                           ByRef dtStoredValues As DataTable, ByVal lCustomerPk As Long, ByVal sStoreNum As String, _
                                           ByVal sTerminalId As String, ByVal sTrxNum As String, ByVal sMemberId As String, _
                                           ByVal sPrimaryId As String, ByVal sSecondaryId As String, _
                                           ByVal sAltId As String, ByVal sVerifier As String, ByVal sMemberLevelInfo As String, _
                                           ByVal bUsedAltId As Boolean, ByVal bCustomerDisabled As Boolean, _
                                           ByVal bLifetime As Boolean, ByVal bReadOnly As Boolean, ByVal OrgCustomerPK As Long, ByVal dtSupplementals as DataTable) As String
    Dim sw As StringWriter = Nothing
    Dim Writer As XmlWriter = Nothing
    Dim Settings As XmlWriterSettings
    Dim dr As DataRow
    Dim decTemp As Decimal
    Dim sXml As String
    Dim sTemp As String
    Dim dtTemp As Date
    Dim iPrecision As Integer = 100
    Dim iLockStatus As Integer
    Dim sLocked As String = "false"
    Dim sPrepay As String = "false"
    Dim bEmployee As Boolean
    Dim iQtyEarned As Integer
    Dim iQtyUsed As Integer
    Dim iQtyBalance As Integer
    Dim decTotal As Decimal
    Dim decTotalEarned As Decimal
    Dim decTotalUsed As Decimal
    Dim bUseCustomerLocking As Boolean = False
    Dim sComment As String
    Dim sTaxID As String
    Dim sDriverLicense As String
    Dim dDateOpened As Date = Nothing
    Dim sDateOpened As Date = Nothing
    Dim bARFlag As Boolean
    Dim iCompoundFlag As Integer
    Dim iFinanceFlag As Integer
    Dim decCreditLimit As Decimal
    Dim decAPRPercent As Decimal
    Dim sARFlag As String = "true"
    Dim sAddress As String
    Dim sCity As String
    Dim sState As String
    Dim sZip As String
    Dim sPhone As String
    Dim sMemberShipArray() As String
    Dim sProgramId As String
    Dim bAccumulateProgram As Boolean = False
    Dim bPointsProgram As Boolean = False
    Dim bFuelPartnerProgram As Boolean = False
    Dim dValue As Decimal
    Dim lUpdateCount As Long = 0
    Dim iCustomerStatusId As Integer
    Dim sCardStatus As String
    Dim sDigitalReceipt As String
    Dim bPaperReceipt As Boolean
    Dim sLinkedID As String = ""

    sCurrentMethod = "ConvertAccountInfoToXml"
    Try
      sDefaultFirstName = MyCommon.Fetch_CM_SystemOption(20)
      sDefaultLastName = MyCommon.Fetch_CM_SystemOption(21)

      If Not bReadOnly Then
        ' Always lock customer if there are SV data
        If (Not dtStoredValues Is Nothing) AndAlso dtStoredValues.Rows.Count > 0 Then
          bUseCustomerLocking = True
        Else
          ' Check if there are PromoVar data
          If (Not dtPromoVars Is Nothing) AndAlso dtPromoVars.Rows.Count > 0 Then
            'Check to see if lock customer data in general
            If MyCommon.Fetch_CM_SystemOption(36) = "1" Then
              bUseCustomerLocking = True
            End If
          End If
        End If
      End If

      If bUseCustomerLocking Then
        iLockStatus = CustomerLockSet(lCustomerPk, sStoreNum, sTerminalId, sTrxNum, lUpdateCount, OrgCustomerPK)
        sCurrentMethod = "ConvertAccountInfoToXml"
        If iLockStatus > 0 Then
          sLocked = "true"
          If iLockStatus = 2 Then
            sPrepay = "true"
          End If
        End If
      End If

      sw = New StringWriter()
      Settings = New XmlWriterSettings()
      Settings.Indent = True
      Writer = XmlWriter.Create(sw, Settings)

      Writer.WriteStartDocument()
      Writer.WriteStartElement("Account", sXmls)
      Writer.WriteAttributeString("Type", "LOYALTY")
      Writer.WriteStartElement("LoyaltyAccount")
      If Not dtAccountInfo Is Nothing AndAlso dtAccountInfo.Rows.Count > 0 Then
        If sPrimaryId.Length = 0 Then
          Writer.WriteAttributeString("ID", sMemberId)
        Else
          Writer.WriteAttributeString("ID", sPrimaryId)
        End If
        dr = dtAccountInfo.Rows(0)
        If MyCommon.NZ(dr.Item("Employee"), False) Then
          bEmployee = True
          If bUsedAltId Then
            Dim sAllowEmpToUseAltId As String = MyCommon.Fetch_CM_SystemOption(17)
            If sAllowEmpToUseAltId = "0" Then
              bEmployee = False
            End If
          End If
        Else
          bEmployee = False
        End If
        If bEmployee Then
          Writer.WriteAttributeString("Employee", "1")
        Else
          Writer.WriteAttributeString("Employee", "0")
        End If

        ' "lUpdateCount <> 0" indicates that an EXISTING customer lock was reset in function "CustomerLockSet".
        ' If lock was NOT reset, then use original UpdateCount returned via intitial customer info lookup
        If lUpdateCount = 0 Then
          ' Note: UpdateCount comes from Household row if customer belongs to Household
          '       see sp pa_LogixServ_CustInfo in XS database
          lUpdateCount = MyCommon.NZ(dr.Item("UpdateCount"), 0)
        End If
        Writer.WriteAttributeString("UpdateCount", lUpdateCount)
        sComment = MyCommon.NZ(dr.Item("Comment"), "")
        If sComment.Length > 0 Then
          Writer.WriteAttributeString("Comment", sComment)
        End If

        If sPrimaryId.Length > 0 Then
          Writer.WriteAttributeString("MemberCardID", sMemberId)
        End If

        If sSecondaryId.Length > 0 Then
          Writer.WriteAttributeString("SecondaryCardID", sSecondaryId)
        End If

        If sAltId.Length > 0 Then
          Writer.WriteAttributeString("AlternateID", sAltId)
        End If

        If sVerifier.Length > 0 Then
          Writer.WriteAttributeString("Verifier", sVerifier)
        End If

        ' Find Linked Card if enabled
        If MyCommon.NZ(MyCommon.Fetch_CM_SystemOption(143), "") = "1" Then
          sLinkedId = MyCommon.NZ(dr.Item("LinkedID"), "")
          If sLinkedId.Length > 0 Then
            Writer.WriteAttributeString("LinkedID", sLinkedId)
          End If
        End If

        Writer.WriteStartElement("Status")
        If bCustomerDisabled Then
          Writer.WriteAttributeString("Val", scDisabled)
        Else
          If Not Integer.TryParse(MyCommon.NZ(dr.Item("CustomerStatusID"), ""), iCustomerStatusId) Then iCustomerStatusId = 0
          ' Is customer ACTIVE?
          If iCustomerStatusId = 1 Then
            ' Yes, so look at card status
            sCardStatus = MyCommon.NZ(dr.Item("CardStatus"), "")
            ' Interpret card status other than ACTIVE as INACTIVE?
            If MyCommon.Fetch_CM_SystemOption(iInterpretCardStatusOptionId) = "1" Then
              ' Yes, so is it active?
              If sCardStatus = scActive Then
                Writer.WriteAttributeString("Val", scActive)
              Else
                Writer.WriteAttributeString("Val", scInactive)
              End If
            Else
              'No, so return actual card status
              Writer.WriteAttributeString("Val", sCardStatus)
            End If
          Else
            ' No, so return INACTIVE status
            Writer.WriteAttributeString("Val", scInactive)
          End If
        End If
        Writer.WriteAttributeString("Locked", sLocked)
        Writer.WriteAttributeString("Prepay", sPrepay)
        Writer.WriteEndElement() ' Status

        If sMemberLevelInfo.Length > 1 Then
          sMemberShipArray = sMemberLevelInfo.Split(",")
          If sMemberShipArray(0).Length > 0 Then
            Writer.WriteStartElement("Level")
            Writer.WriteAttributeString("Val", sMemberShipArray(0))
            Writer.WriteAttributeString("Desc", sMemberShipArray(1))
            Writer.WriteEndElement() ' Level
          End If
        End If

        Writer.WriteStartElement("Customer")
        Writer.WriteAttributeString("FirstName", MyCommon.NZ(dr.Item("FirstName"), sDefaultFirstName))
        Writer.WriteAttributeString("LastName", MyCommon.NZ(dr.Item("LastName"), sDefaultLastName))
        sTaxID = MyCommon.NZ(dr.Item("TaxID"), "")
        If (sTaxID.Length > 0) Then
          Writer.WriteAttributeString("TaxId", sTaxID)
        End If
        sDriverLicense = MyCommon.NZ(dr.Item("DriverLicense"), "")
        If (sDriverLicense.Length > 0) Then
                    Writer.WriteAttributeString("DriverLicense", MyCryptLib.SQL_StringDecrypt(sDriverLicense))
        End If
        sAddress = MyCommon.NZ(dr.Item("Address"), "")
        If (sAddress.Length > 0) Then
          Writer.WriteAttributeString("Address", sAddress)
        End If
        sCity = MyCommon.NZ(dr.Item("City"), "")
        If (sCity.Length > 0) Then
          Writer.WriteAttributeString("City", sCity)
        End If
        sState = MyCommon.NZ(dr.Item("State"), "")
        If (sState.Length > 0) Then
          Writer.WriteAttributeString("State", sState)
        End If
        sZip = MyCommon.NZ(dr.Item("Zip"), "")
        If (sZip.Length > 0) Then
          Writer.WriteAttributeString("Zip", sZip)
        End If
        sPhone = MyCommon.NZ(dr.Item("Phone"), "")
        If (sPhone.Length > 0) Then
                    Writer.WriteAttributeString("Phone", MyCryptLib.SQL_StringDecrypt(sPhone))
        End If
        bARFlag = MyCommon.NZ(dr.Item("ARFlag"), False)
        If bARFlag Then
          Writer.WriteAttributeString("ARFlag", sARFlag)
        End If
        sDateOpened = Date.Parse(MyCommon.NZ(dr.Item("DateOpened"), "12/12/1900"))
        If (sDateOpened > "12/12/1900") Then
          Writer.WriteAttributeString("DateOpened", Format(sDateOpened, sDateFormat))
        End If
        iCompoundFlag = MyCommon.NZ(dr.Item("CompoundFlag"), 0)
        If (iCompoundFlag = 1) Then
          Writer.WriteAttributeString("CompoundFlag", iCompoundFlag)
        End If
        iFinanceFlag = MyCommon.NZ(dr.Item("FinanceFlag"), 0)
        If (iFinanceFlag = 1) Then
          Writer.WriteAttributeString("FinanceFlag", iFinanceFlag)
        End If
        decCreditLimit = MyCommon.NZ(dr.Item("CreditLimit"), 0)
        If (decCreditLimit > 0) Then
          Writer.WriteAttributeString("CreditLimit", decCreditLimit)
        End If
        decAPRPercent = MyCommon.NZ(dr.Item("APRPercent"), 0)
        If (decAPRPercent > 0) Then
          Writer.WriteAttributeString("APRPercent", decAPRPercent)
        End If
        If (MyCommon.Fetch_SystemOption(212) <> 0) Then
          sDigitalReceipt = MyCommon.NZ(dr.Item("DigitalReceipt"), 0)
          If (sDigitalReceipt.Length > 0) Then
            Writer.WriteAttributeString("DigitalReceipt", sDigitalReceipt)
          End If
        End If

        If (MyCommon.Fetch_SystemOption(310) <> 0) Then
          If MyCommon.NZ(dr.Item("PaperReceipt"), False) Then
            bPaperReceipt = True
          Else
            bPaperReceipt = False
          End If
          If bPaperReceipt Then
            Writer.WriteAttributeString("PaperReceipt", "1")
          Else
            Writer.WriteAttributeString("PaperReceipt", "0")
          End If
        End If

        Writer.WriteEndElement() ' Customer
        If (MyCommon.NZ(MyCommon.Fetch_CM_SystemOption(141), "0") = "1") Then
          Writer.WriteStartElement("Supplementals")
          If Not dtSupplementals Is Nothing AndAlso dtSupplementals.Rows.Count > 0 Then
            For Each dr In dtSupplementals.Rows
              Writer.WriteStartElement("Supplemental")
              Writer.WriteAttributeString("ExtID", MyCommon.NZ(dr.Item("ExtFieldID"), "0"))
              Writer.WriteAttributeString("Val", MyCommon.NZ(dr.Item("Value"), "0"))
              Writer.WriteEndElement() ' Supplemental
            Next
          End If
          Writer.WriteEndElement() ' Supplementals
        End If
      Else
        Writer.WriteAttributeString("ID", "0")
      End If

      Writer.WriteStartElement("Promos")
      If Not dtPromos Is Nothing AndAlso dtPromos.Rows.Count > 0 Then
        For Each dr In dtPromos.Rows
          Writer.WriteStartElement("Promo")
          Writer.WriteAttributeString("ID", MyCommon.NZ(dr.Item("OfferId"), "0"))
          Writer.WriteEndElement() ' Promo
        Next
      End If
      Writer.WriteEndElement() ' Promos

      Writer.WriteStartElement("PromoVars")
      If Not dtPromoVars Is Nothing AndAlso dtPromoVars.Rows.Count > 0 Then
        For Each dr In dtPromoVars.Rows
          decTemp = MyCommon.NZ(dr.Item("Val"), 0)
          Writer.WriteStartElement("PromoVar")
          Writer.WriteAttributeString("ID", MyCommon.NZ(dr.Item("Id"), "0"))
          Writer.WriteAttributeString("Val", decTemp.ToString("#0"))
          If bLifetime Then
            sTemp = MyCommon.NZ(dr.Item("Lifetime"), "")
            If sTemp <> "" Then
              If Not Decimal.TryParse(sTemp, decTemp) Then decTemp = 0.0
              Writer.WriteAttributeString("Lifetime", decTemp.ToString("#0"))
            End If
          End If

          ' Add pending cart ID to pending promo var values
          If dr.Table.Columns.Contains("CartID") Then
            sTemp = MyCommon.NZ(dr.Item("CartID"), "")
            If sTemp <> "" Then
              Writer.WriteAttributeString("PendingCartID", sTemp)
            End If
          End If

          Writer.WriteEndElement() ' PromoVar
        Next
      End If
      Writer.WriteEndElement() ' PromoVars

      Writer.WriteStartElement("StoredValues")
      If Not dtStoredValues Is Nothing AndAlso dtStoredValues.Rows.Count > 0 Then
        For Each dr In dtStoredValues.Rows
          iQtyEarned = MyCommon.NZ(dr.Item("QtyEarned"), 0)
          iQtyUsed = MyCommon.NZ(dr.Item("QtyUsed"), 0)
          iQtyBalance = iQtyEarned - iQtyUsed
          If iQtyBalance > 0 Then
            Writer.WriteStartElement("StoredValue")
            sProgramId = MyCommon.NZ(dr.Item("SVProgramID"), "0")
            Writer.WriteAttributeString("ProgramID", sProgramId)
            If MyCommon.NZ(dr.Item("StatusFlag"), 0) = 1 Then
              sTemp = "Issued"
            Else
              sTemp = "Unknown"
            End If
            Writer.WriteAttributeString("Status", sTemp)
            Writer.WriteAttributeString("LocalID", MyCommon.NZ(dr.Item("LocalID"), "0"))
            Writer.WriteAttributeString("ServerSerial", MyCommon.NZ(dr.Item("ServerSerial"), "0"))
            Writer.WriteAttributeString("ExternalID", MyCommon.NZ(dr.Item("ExternalID"), "0"))
            dtTemp = Date.Parse(MyCommon.NZ(dr.Item("ExpireDate"), "01/01/2001"))
            Writer.WriteAttributeString("Expiration", Format(dtTemp, sDateFormat))
            GetSVProgramDetails(sProgramId, bAccumulateProgram, bPointsProgram, bFuelPartnerProgram, dValue)
            decTemp = dValue
            If Not bPointsProgram Then
              decTemp *= iPrecision
            End If
            Writer.WriteAttributeString("UnitValue", decTemp.ToString("#0"))
            Writer.WriteAttributeString("QuantityEarned", iQtyBalance.ToString("#0"))

            ' compute Total Balance Value
            ' The POS expects ComputedTotalAmount = QtyBalance * Unit Value
            decTotalEarned = MyCommon.NZ(dr.Item("TotalValueEarned"), 0)
            decTotalUsed = MyCommon.NZ(dr.Item("RedeemedValue"), 0)
            decTotal = decTotalEarned - decTotalUsed
            If Not bPointsProgram Then
              decTotal *= iPrecision
            End If
            Writer.WriteAttributeString("ComputedTotalAmount", decTotal.ToString("#0"))

            ' always return 0 for these, since they are NOT used by POS in GetAccount
            ' however, they are used in PutTransaction
            decTemp = 0.0
            Writer.WriteAttributeString("ComputedRedeemedAmount", decTemp.ToString("#0"))
            Writer.WriteAttributeString("ComputedBreakageAmount", decTemp.ToString("#0"))
            Writer.WriteEndElement() ' StoredValue
          End If
        Next
      End If
      Writer.WriteEndElement() ' StoredValues
      Writer.WriteEndElement() ' LoyaltyAccount
      Writer.WriteEndElement() ' Account
      Writer.WriteEndDocument()
      Writer.Flush()

      sXml = sw.ToString
    Catch
      Throw
    Finally
      If Not Writer Is Nothing Then
        Writer.Close()
      End If
      If Not sw Is Nothing Then
        sw.Close()
        sw.Dispose()
      End If
    End Try
    Return sXml
  End Function

  Private Function ConvertCouponInfoToXml(ByRef dtStoredValues As DataTable, ByVal sStoreNum As String, _
                                          ByVal sTerminalId As String, ByVal sTrxNum As String) As String
    Dim sw As StringWriter = Nothing
    Dim Writer As XmlWriter = Nothing
    Dim Settings As XmlWriterSettings
    Dim dr As DataRow
    Dim decTemp As Decimal
    Dim sXml As String
    Dim sTemp As String
    Dim sLocked As String = "false"
    Dim sPrepay As String = "false"
    Dim dtTemp As Date
    Dim iPrecision As Integer = 100
    Dim iLockStatus As Integer
    Dim lStoredValueId As Long
    Dim iStatus As Integer

    sCurrentMethod = "ConvertCouponInfoToXml"
    Try
      ' Lock Coupon if there is SV data
      If dtStoredValues.Rows.Count = 0 Then
        sXml = BuildErrorXml("Coupon not found!", ErrorType.General, False)
        Exit Try
      ElseIf dtStoredValues.Rows.Count > 1 Then
        sXml = BuildErrorXml("CouponId is not unique!", ErrorType.General, False)
        Exit Try
      Else
        lStoredValueId = dtStoredValues.Rows(0).Item("StoredValueId")
      End If

      sw = New StringWriter()
      Settings = New XmlWriterSettings()
      Settings.Indent = True
      Writer = XmlWriter.Create(sw, Settings)

      Writer.WriteStartDocument()
      Writer.WriteStartElement("Account", sXmls)
      Writer.WriteAttributeString("Type", "LOYALTY")
      Writer.WriteStartElement("LoyaltyAccount")
      Writer.WriteAttributeString("ID", "")
      Writer.WriteAttributeString("UpdateCount", "0")
      Writer.WriteStartElement("Status")
      Writer.WriteAttributeString("Val", "ACTIVE")
      Writer.WriteAttributeString("Locked", sLocked)
      Writer.WriteAttributeString("Prepay", sPrepay)
      Writer.WriteEndElement() ' Status

      Writer.WriteStartElement("StoredValues")
      If Not dtStoredValues Is Nothing AndAlso dtStoredValues.Rows.Count > 0 Then
        For Each dr In dtStoredValues.Rows
          Writer.WriteStartElement("StoredValue")
          Writer.WriteAttributeString("ProgramID", dr.Item("SVProgramID"))
          iStatus = dr.Item("StatusFlag")
          Select Case iStatus
            Case 1
              iLockStatus = StoredValueLockSet(lStoredValueId, sStoreNum, sTerminalId, sTrxNum)
              If iLockStatus > 0 Then
                sTemp = "Locked"
              Else
                sTemp = "Issued"
              End If
            Case 2
              sTemp = "Revoked"
            Case 3
              sTemp = "Expired"
            Case 4
              sTemp = "Redeemed"
            Case Else
              sTemp = "Unknown"
              WriteLog("Unknown status: '" & iStatus & "' for chit: " & dr.Item("ExternalID") & " (Store: " & ")", MessageType.Warning)
          End Select
          Writer.WriteAttributeString("Status", sTemp)
          Writer.WriteAttributeString("LocalID", dr.Item("LocalID"))
          Writer.WriteAttributeString("ServerSerial", dr.Item("ServerSerial"))
          Writer.WriteAttributeString("ExternalID", dr.Item("ExternalID"))
          dtTemp = Date.Parse(dr.Item("ExpireDate"))
          Writer.WriteAttributeString("Expiration", Format(dtTemp, sDateFormat))
          decTemp = dr.Item("Value")
          decTemp *= iPrecision
          Writer.WriteAttributeString("UnitValue", decTemp.ToString("#0"))
          Writer.WriteAttributeString("QuantityEarned", dr.Item("QtyEarned"))
          decTemp = dr.Item("TotalValueEarned")
          decTemp *= iPrecision
          Writer.WriteAttributeString("ComputedTotalAmount", decTemp.ToString("#0"))
          decTemp = dr.Item("RedeemedValue")
          decTemp *= iPrecision
          Writer.WriteAttributeString("ComputedRedeemedAmount", decTemp.ToString("#0"))
          decTemp = dr.Item("BreakageValue")
          decTemp *= iPrecision
          Writer.WriteAttributeString("ComputedBreakageAmount", decTemp.ToString("#0"))
          Writer.WriteEndElement() ' StoredValue
        Next
      End If
      Writer.WriteEndElement() ' StoredValues
      Writer.WriteEndElement() ' LoyaltyAccount
      Writer.WriteEndElement() ' Account
      Writer.WriteEndDocument()
      Writer.Flush()

      sXml = sw.ToString
    Catch ex As Exception
      sXml = BuildErrorXml(ex.Message, eDefaultErrorType)
    Finally
      If Not Writer Is Nothing Then
        Writer.Close()
      End If
      If Not sw Is Nothing Then
        sw.Close()
        sw.Dispose()
      End If
    End Try
    Return sXml
  End Function

  Private Function ConvertMemberNotFoundToXml(ByVal sId As String, ByVal iIdType As Integer) As String
    Dim sXml As String
    Dim dt As DataTable
    Dim sCardDesc As String

    sXml = "<?xml version=""1.0"" encoding=""utf-16""?><Account Type=""LOYALTY"" xmlns=""http://www.ncr.com/rsd/cm/accounts/1.0"">" & _
           "<LoyaltyAccount ID=""" & sId & _
           """ UpdateCount=""0""><Status Val=""NOT_ON_FILE""/></LoyaltyAccount></Account>"

    If iIdType = 999 Then
      sCardDesc = "Coupon"
    Else
      MyCommon.QueryStr = "select Description, PhraseID, ExtCardTypeID from CardTypes with (NoLock) where CardTypeId=" & iIdType & ";"
      dt = MyCommon.LXS_Select
      If dt.Rows.Count > 0 Then
        sCardDesc = MyCommon.NZ(dt.Rows(0).Item("Description"), "Unknown Card Type")
      Else
        sCardDesc = "Unknown Card Type"
      End If
    End If
    WriteLog(sCardDesc & " '" & sId & "' is not on file!", MessageType.Warning)

    Return sXml
  End Function

  Private Function ConvertPendingInfoToXml(ByVal sMemberId As String, _
                                           ByRef dtPendingPoints As DataTable, _
                                           ByRef dtPendingDistribution As DataTable, _
                                           ByRef dtPendingCPERewardDistribution As DataTable, _
                                           ByRef dtPendingRewardLimits As DataTable) As String
    Dim sw As StringWriter = Nothing
    Dim Writer As XmlWriter = Nothing
    Dim Settings As XmlWriterSettings
    Dim dr As DataRow
    Dim sXml As String

    sCurrentMethod = "ConvertPendingInfoToXml"
    Try
      sw = New StringWriter()
      Settings = New XmlWriterSettings()
      Settings.Indent = True
      Writer = XmlWriter.Create(sw, Settings)

      Writer.WriteStartDocument()
      Writer.WriteStartElement("Customer", sXmls)
      Writer.WriteAttributeString("ID", sMemberId)

      Writer.WriteStartElement("PointsPending")
      If Not dtPendingPoints Is Nothing AndAlso dtPendingPoints.Rows.Count > 0 Then
        For Each dr In dtPendingPoints.Rows
          Writer.WriteStartElement("Point")
          Writer.WriteAttributeString("PromoVarID", MyCommon.NZ(dr.Item("PromoVarID"), "0"))
          Writer.WriteAttributeString("ProgramID", MyCommon.NZ(dr.Item("ProgramID"), "0"))
          Writer.WriteAttributeString("EarnedAmount", MyCommon.NZ(dr.Item("EarnedAmount"), "0"))
          Writer.WriteAttributeString("RedeemedAmount", MyCommon.NZ(dr.Item("RedeemedAmount"), "0"))
          Writer.WriteAttributeString("CartID", MyCommon.NZ(dr.Item("CartID"), "0"))
          Writer.WriteAttributeString("ExtLocationCode", MyCommon.NZ(dr.Item("ExtLocationCode"), "0"))
          Writer.WriteAttributeString("POSTimeStamp", MyCommon.NZ(dr.Item("POSTimeStamp"), "0"))
          Writer.WriteEndElement() ' Point
        Next
      End If
      Writer.WriteEndElement() ' PointsPending

      Writer.WriteStartElement("DistributionVariablesPending")
      If Not dtPendingDistribution Is Nothing AndAlso dtPendingDistribution.Rows.Count > 0 Then
        For Each dr In dtPendingDistribution.Rows
          Writer.WriteStartElement("DistributionVariable")
          Writer.WriteAttributeString("PromoVarID", MyCommon.NZ(dr.Item("PromoVarID"), "0"))
          Writer.WriteAttributeString("Amount", MyCommon.NZ(dr.Item("Amount"), "0"))
          Writer.WriteAttributeString("CartID", MyCommon.NZ(dr.Item("CartID"), "0"))
          Writer.WriteAttributeString("ExtLocationCode", MyCommon.NZ(dr.Item("ExtLocationCode"), "0"))
          Writer.WriteAttributeString("POSTimeStamp", MyCommon.NZ(dr.Item("POSTimeStamp"), "0"))
          Writer.WriteEndElement() ' DistributionVariable
        Next
      End If
      Writer.WriteEndElement() ' DistributionVariablesPending

      Writer.WriteStartElement("CPE_RewardDistributionPending")
      If Not dtPendingCPERewardDistribution Is Nothing AndAlso dtPendingCPERewardDistribution.Rows.Count > 0 Then
        For Each dr In dtPendingCPERewardDistribution.Rows
          Writer.WriteStartElement("RewardDistribution")
          Writer.WriteAttributeString("IncentiveID", MyCommon.NZ(dr.Item("IncentiveID"), "0"))
          Writer.WriteAttributeString("RewardOptionID", MyCommon.NZ(dr.Item("RewardOptionID"), "0"))
          Writer.WriteAttributeString("CartID", MyCommon.NZ(dr.Item("CartID"), "0"))
          Writer.WriteAttributeString("ExtLocationCode", MyCommon.NZ(dr.Item("ExtLocationCode"), "0"))
          Writer.WriteAttributeString("POSTimeStamp", MyCommon.NZ(dr.Item("POSTimeStamp"), "0"))
          Writer.WriteEndElement() ' RewardDistribution
        Next
      End If
      Writer.WriteEndElement() ' CPE_RewardDistributionPending

      Writer.WriteStartElement("RewardLimits")
      If Not dtPendingRewardLimits Is Nothing AndAlso dtPendingRewardLimits.Rows.Count > 0 Then
        For Each dr In dtPendingRewardLimits.Rows
          Writer.WriteStartElement("RewardLimit")
          Writer.WriteAttributeString("PromoVarID", MyCommon.NZ(dr.Item("PromoVarID"), "0"))
          Writer.WriteAttributeString("Amount", MyCommon.NZ(dr.Item("Amount"), "0"))
          Writer.WriteAttributeString("CartID", MyCommon.NZ(dr.Item("CartID"), "0"))
          Writer.WriteAttributeString("ExtLocationCode", MyCommon.NZ(dr.Item("ExtLocationCode"), "0"))
          Writer.WriteAttributeString("POSTimeStamp", MyCommon.NZ(dr.Item("POSTimeStamp"), "0"))
          Writer.WriteEndElement() ' RewardLimit
        Next
      End If
      Writer.WriteEndElement() ' RewardLimits

      Writer.WriteEndElement() ' Customer
      Writer.WriteEndDocument()
      Writer.Flush()

      sXml = sw.ToString
    Catch
      Throw
    Finally
      If Not Writer Is Nothing Then
        Writer.Close()
      End If
      If Not sw Is Nothing Then
        sw.Close()
        sw.Dispose()
      End If
    End Try
    Return sXml
  End Function

  Private Function ProcessTransactionXml(ByVal sTransactionXml As String, ByVal sStoreId As String) As String
    Dim sStatus As String = sOkStatus
    Dim sHealthStatus As String = sOkStatus
    Dim sType As String
    Dim bBeginTransactionRT As Boolean = False
    Dim bBeginTransactionXS As Boolean = False
    Dim bBeginTransactionWH As Boolean = False
    Dim bTestLocation As Boolean = False
    Dim sr As StringReader = Nothing
    Dim Settings As XmlReaderSettings
    Dim xr As XmlReader = Nothing
    Dim xrSubTree As XmlReader = Nothing
    Dim sUrl As String
    Dim sUserName As String
    Dim sPassword As String
    Dim ExternalCA As Copient.ExternalCustomerAccounts

    ' Note: Originally StoreId (Store Number) was not supplied to this function!
    '       As a result, the store number had to be extracted form the xml in order
    '       to check the computed md5 hash. So the check for duplicates is deferred until
    '       a store number has been extracted. This deferred check remains for
    '       backward compatibility since Store Id may be empty if called via PutTransaction.
    '       This extracted store number was also used to update Logix database.
    ' 
    ' Note: 1) ACS has a limitation that the TRUE store number must be an unsigned short
    '       2) ACS uses this TRUE store number to generate the transaction XML
    '       3) Some ACS/Logix customers require 20 alpha numeric characters for store id
    '
    '       To handle 1) - 3) above, the StoreId, if supplied, is used as the store number
    '       when updating Logix database rather than the store number embedded in the XML.
    Try
      sCurrentStoreId = sStoreId
      sCurrentMethod = "ProcessTransactionXml"
      sInputForLog = "*(Type=Input) (Method=ProcessTransactionXml)] -  (" & sTransactionXml & ")"
      MyCommon = New Copient.CommonInc
      MyCommon.AppName = sAppName
      sStoreId = sStoreId.Trim(" ")

      WriteDebug("PutTransaction", DebugState.BeginTime)

      sXmlMd5Hash = GetStringMd5(sTransactionXml)
      bDuplicateTransactionXml = DuplicateStatus.HasNotBeenChecked

      eDefaultErrorType = ErrorType.SqlServer
      If MyCommon.LRTadoConn.State <> ConnectionState.Open Then
        MyCommon.Open_LogixRT()
        MyCommon.Load_System_Info()
        sInstallationName = MyCommon.InstallationName
      End If
      MyCommon.Open_LogixXS()
      MyCommon.Open_LogixWH()
      eDefaultErrorType = ErrorType.General

      If sStoreId.Length > 0 Then
        sHealthStatus = UpdateLocationHealth(sStoreId, 2, bTestLocation)
        If Not sHealthStatus.Contains("Ok") Then
          sHealthStatus = sOkStatus
        End If
      End If

      ' check if External customer Account is enabled
      If MyCommon.NZ(MyCommon.Fetch_CM_SystemOption(49), "") = "1" Then
        CheckForDuplicateMd5(sStoreId)
        sCurrentMethod = "GetExternalCustomerAccount"
        sUrl = MyCommon.NZ(MyCommon.Fetch_CM_SystemOption(50), "")
        If sUrl <> "" Then
          sUserName = MyCommon.NZ(MyCommon.Fetch_CM_SystemOption(53), "")
          If sUserName <> "" Then
            sPassword = MyCommon.NZ(MyCommon.Fetch_CM_SystemOption(54), "")
          Else
            sPassword = ""
          End If
          WriteDebug("UpdateExternalCustomerAccount", DebugState.BeginTime)
          Try
            ExternalCA = New Copient.ExternalCustomerAccounts()
            ExternalCA.sUrl = sUrl
            ExternalCA.sUserName = sUserName
            ExternalCA.sPassword = sPassword
            sStatus = ExternalCA.UpdateExternalCustomerAccount(sTransactionXml, sStoreId)
            If sStatus.Contains("Ok") Then
              sStatus = sHealthStatus
            End If
          Catch ex As Exception
            Throw ex
          Finally
            WriteDebug("UpdateExternalCustomerAccount", DebugState.EndTime)
          End Try
        Else
          sStatus = BuildErrorXml("URL to External Customer Account web service is not defined in CM Settings", ErrorType.General, False)
        End If
        Exit Try
      Else
        sCurrentMethod = "ProcessTransactionXml"
      End If

      MyCommon.QueryStr = "begin transaction"
      MyCommon.LXS_Execute()
      bBeginTransactionXS = True

      MyCommon.QueryStr = "begin transaction"
      MyCommon.LWH_Execute()
      bBeginTransactionWH = True

      sr = New StringReader(sTransactionXml)

      Settings = New XmlReaderSettings()
      xr = XmlReader.Create(sr, Settings)
      xr.ReadToFollowing("Transactions")
      If xr.EOF Then
        Throw New ApplicationException("No 'Transactions' root element")
      End If

      Do
        xr.ReadToFollowing("Transaction")
        If xr.EOF Then
          Throw New ApplicationException("No 'Transaction' element")
        End If
        Do
          If xr.AttributeCount = 0 Then
            Throw New ApplicationException("No attributes in 'Transaction' element.")
          End If

          sType = xr.GetAttribute("Type")
          If sType Is Nothing Then
            Throw New ApplicationException("No 'Type' attribute in 'Transaction' element.")
          End If

          xrSubTree = xr.ReadSubtree()
          If sType = "LOYALTY" Then
            ProcessLoyaltyTransaction(xrSubTree, sStoreId)
          ElseIf sType = "STORED_VALUE" Then
            ProcessStoredValueTransaction(xrSubTree, sStoreId)
          ElseIf sType = "CUST_GROUP" Then
            ProcessCustGroupTransaction(xrSubTree, sStoreId)
          ElseIf sType = "PROMO_MOVT" Then
            ProcessPromoMovtTransaction(xrSubTree, sStoreId)
          ElseIf sType = "PROMO_MOVT_SUMMARY" Then
            PromoSummaryTransaction(xrSubTree, sStoreId)
          ElseIf sType = "PROMO_MOVT_BULK" Then
            ProcessPromoBulkTransaction(sTransactionXml, sStoreId)
          ElseIf sType = "TEST_SQL_ERROR" Then
            eDefaultErrorType = ErrorType.SqlServer
            Throw New ApplicationException("This is a SQL Server error test")
          Else
            Throw New ApplicationException("Invalid Transaction Type: " & sType)
          End If
          xrSubTree.Close()
          xrSubTree = Nothing

        Loop While xr.ReadToNextSibling("Transaction")
      Loop While xr.ReadToNextSibling("Transactions")

      UpdateLocationMd5()
      If bBeginTransactionRT Then
        MyCommon.QueryStr = "commit transaction"
        MyCommon.LRT_Execute()
      End If
      If bBeginTransactionXS Then
        MyCommon.QueryStr = "commit transaction"
        MyCommon.LXS_Execute()
      End If
      If bBeginTransactionWH Then
        MyCommon.QueryStr = "commit transaction"
        MyCommon.LWH_Execute()
      End If
      If bBeginTransactionEX Then
        MyCommon.QueryStr = "commit transaction"
        MyCommon.LEX_Execute()
      End If
      sStatus = sHealthStatus
    Catch exApp As ApplicationException
      If bBeginTransactionRT Then
        MyCommon.QueryStr = "rollback transaction"
        MyCommon.LRT_Execute()
      End If
      If bBeginTransactionXS Then
        MyCommon.QueryStr = "rollback transaction"
        MyCommon.LXS_Execute()
      End If
      If bBeginTransactionWH Then
        MyCommon.QueryStr = "rollback transaction"
        MyCommon.LWH_Execute()
      End If
      If bBeginTransactionEX Then
        MyCommon.QueryStr = "rollback transaction"
        MyCommon.LEX_Execute()
      End If
      If bDuplicateTransactionXml = DuplicateStatus.IsDuplicate Then
        sStatus = sHealthStatus
        WriteLog(exApp.Message, MessageType.Warning)
      Else
        sStatus = BuildErrorXml(exApp.Message, eDefaultErrorType, False)
      End If
    Catch exXml As XmlException
      If bBeginTransactionRT Then
        MyCommon.QueryStr = "rollback transaction"
        MyCommon.LRT_Execute()
      End If
      If bBeginTransactionXS Then
        MyCommon.QueryStr = "rollback transaction"
        MyCommon.LXS_Execute()
      End If
      If bBeginTransactionWH Then
        MyCommon.QueryStr = "rollback transaction"
        MyCommon.LWH_Execute()
      End If
      If bBeginTransactionEX Then
        MyCommon.QueryStr = "rollback transaction"
        MyCommon.LEX_Execute()
      End If
      sStatus = BuildErrorXml(exXml.Message, eDefaultErrorType, False)
    Catch ex As Exception
      If bBeginTransactionRT Then
        MyCommon.QueryStr = "rollback transaction"
        MyCommon.LRT_Execute()
      End If
      If bBeginTransactionXS Then
        MyCommon.QueryStr = "rollback transaction"
        MyCommon.LXS_Execute()
      End If
      If bBeginTransactionWH Then
        MyCommon.QueryStr = "rollback transaction"
        MyCommon.LWH_Execute()
      End If
      If bBeginTransactionEX Then
        MyCommon.QueryStr = "rollback transaction"
        MyCommon.LEX_Execute()
      End If
      sStatus = BuildErrorXml(ex.Message, eDefaultErrorType)
    Finally
      If Not xrSubTree Is Nothing Then
        xrSubTree.Close()
        xrSubTree = Nothing
      End If
      If Not xr Is Nothing Then
        xr.Close()
        xr = Nothing
      End If
      If Not sr Is Nothing Then
        sr.Close()
        sr.Dispose()
      End If
      If MyCommon.LEXadoConn.State <> ConnectionState.Closed Then
        MyCommon.Close_LogixEX()
      End If
      If MyCommon.LWHadoConn.State <> ConnectionState.Closed Then
        MyCommon.Close_LogixWH()
      End If
      If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then
        MyCommon.Close_LogixXS()
      End If
      WriteDebug("PutTransaction", DebugState.EndTime)
      If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then
        MyCommon.Close_LogixRT()
      End If
    End Try
    Return sStatus
  End Function

  Private Sub ProcessLoyaltyTransaction(ByRef xr As XmlReader, ByVal sStoreId As String)
    Dim sId As String
    Dim sPromoVarId As String
    Dim sValue As String
    Dim sDateTime As String
    Dim sStoreNum As String
    Dim sTerminalNum As String
    Dim sTransNum As String
    Dim sLogixTransNum As String
    Dim sPrepay As String
    Dim sLocationId As String = "0"
    Dim sLocalServerId As String = "0"
    Dim sarrAltIdTableCol() As String
    Dim sAltIdTableCol As String
    Dim sAltIdTable As String
    Dim sAltIdCol As String
        Dim sAltId As String
        Dim sValAltId As String
    Dim sCartID As String
    Dim lCustomerPk As Long
    Dim lHhPk As Long
    Dim lUseCustomerPk As Long
    Dim bNewCustomer As Boolean
    Dim dt As DataTable = Nothing
    Dim xrSubTree As XmlReader = Nothing
    Dim xrSubTree1 As XmlReader = Nothing
    Dim sPreviousId As String = ""
    Dim bUseGeneralCustomerLocking As Boolean
    Dim sTenderPointIds As String
    Dim sEarned As String
    Dim bLifetime As Boolean = False
    Dim iTrxCardTypeId As Integer
    Dim bTestCustomer As Boolean = False
    Dim bDuplicate As Boolean
    Dim bCardFound As Boolean
    Dim bFobEligibilityEnabled As Boolean = False
    Dim bLinkedIDEnabled As Boolean = False
    Dim sSecondaryId As String
    Dim lOldCustomerPK As Long = 0
    Dim lOldHhPk As Long = 0
    Dim cardValidationResp As CardValidationResponse
    Dim sAwardName As String = ""
    Dim sPromoId As String = ""
    Dim sContent As String = ""

    sCurrentMethod = "ProcessLoyaltyTransaction"
    WriteDebug(sCurrentMethod, DebugState.BeginTime)

    Try
      iTrxCardTypeId = GetTrxCardTypeId()

      sTenderPointIds = MyCommon.NZ(MyCommon.Fetch_CM_SystemOption(37), "")
      If sTenderPointIds = "" Then
        bLifetime = False
      Else
        bLifetime = True
      End If

      xr.ReadToFollowing("LoyaltyTransaction")
      If xr.EOF Then
        Throw New ApplicationException("No 'LoyaltyTransaction' element")
      End If

      ' check to see if general customer locking is on
      If MyCommon.Fetch_CM_SystemOption(36) = "1" Then
        ' yes, so unlock when processing Loyalty transaction type
        bUseGeneralCustomerLocking = True
      Else
        ' no, so unlock when processing Stored Values
        bUseGeneralCustomerLocking = False
      End If

      Do
        If xr.AttributeCount = 0 Then
          Throw New ApplicationException("No attributes in 'LoyaltyTransaction' element.")
        End If

        sDateTime = GetLocalDateTime(xr.GetAttribute("DateTime"))

        If sStoreId.Length > 0 Then
          ' use store number from sending store (20 Alpha Numeric)
          sStoreNum = sStoreId
        Else
          ' use store number embedded in xml (unsigned short)
          sStoreNum = xr.GetAttribute("StoreNum")
        End If

        If sStoreNum Is Nothing OrElse sStoreNum.Length = 0 Then
          sStoreNum = "0"
        Else
          CheckForDuplicateMd5(sStoreNum)
          sCurrentMethod = "ProcessLoyaltyTransaction"
        End If

        sTerminalNum = xr.GetAttribute("TerminalNum")
        If sTerminalNum Is Nothing OrElse sTerminalNum.Length = 0 Then
          sTerminalNum = "0"
        End If

        sTransNum = xr.GetAttribute("TransNum")
        If sTransNum Is Nothing OrElse sTerminalNum.Length = 0 Then
          sTransNum = "0"
        End If
        If sTransNum.Length >128 then
            Throw New ApplicationException("Transaction Number length of'" & sTransNum & "' should not be greater than 128 characters")
        End If
        sLogixTransNum = xr.GetAttribute("LogixTransNum")
        If sLogixTransNum Is Nothing OrElse sLogixTransNum.Length = 0 Then
          sLogixTransNum = "1"
        End If

        sPrepay = xr.GetAttribute("Prepay")
        If sPrepay Is Nothing OrElse sPrepay.Length = 0 Then
          sPrepay = "False"
        End If

        sId = xr.GetAttribute("ID")
        If sId Is Nothing OrElse sId.Length = 0 Then
          Throw New ApplicationException("No valid 'ID' attribute in 'LoyaltyTransaction' element.")
        End If

        bLinkedIDEnabled = IIf(MyCommon.Fetch_CM_SystemOption(143) = "1", True, False)
        bFobEligibilityEnabled = IIf(MyCommon.Fetch_CM_SystemOption(142) = "1", True, False)
        If bFobEligibilityEnabled or bLinkedIDEnabled Then
          ' FOB path - Get Customer PK & CardTypeID for this ID
          bCardFound = GetCustomerPK_CardIdOnly(sId, iTrxCardTypeId, lCustomerPk, lHhPk, bTestCustomer)
          If Not bCardFound Then
            Throw New ApplicationException("Card NOT found for Card ID '" & sId & "'!")
            Exit Try
          End If
        Else
          ' Normal path - Get Customer PK if Customer is a member
        sId = MyCommon.Pad_ExtCardID(sId, iTrxCardTypeId)
        bCardFound = GetCustomerPK(sId, iTrxCardTypeId, lCustomerPk, lHhPk, bTestCustomer)
        If Not bCardFound Then
          If Not MyCommon.AllowToProcessCustomerCard(sId, iTrxCardTypeId, cardValidationResp) Then
            Throw New ApplicationException(MyCommon.CardValidationResponseMessage(sId, iTrxCardTypeId, cardValidationResp))
            Exit Try
          End If
          bCardFound = AddCustomer(sId, iTrxCardTypeId, lCustomerPk, lHhPk, bTestCustomer)
          If bCardFound Then
            bNewCustomer = True
          Else
            Throw New ApplicationException("Card NOT found for Card ID '" & sId & "' of type '" & iTrxCardTypeId & "'!")
            Exit Try
          End If
        End If
       End If

        ' see if householded
        If lHhPk = 0 Then
          lUseCustomerPk = lCustomerPk
        Else
          lUseCustomerPk = lHhPk
        End If

        MyCommon.QueryStr = "dbo.pa_CM_Gen_CheckExtLocationCode"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@ExtLocationCode", SqlDbType.NVarChar).Value = sStoreNum
        dt = MyCommon.LRTsp_select
        MyCommon.Close_LRTsp()
        If dt.Rows.Count > 0 Then
          sLocationId = MyCommon.NZ(dt.Rows(0).Item("LocationID"), "0")
          sLocalServerId = MyCommon.NZ(dt.Rows(0).Item("LocalServerID"), "0")
        End If
        dt = Nothing

        If sLocationId = "0" OrElse sLocalServerId = "0" Then
          Throw New ApplicationException("LocationId or LocalServerID not found for ExtLocationCode: " & sStoreNum)
        End If

                sAltId = xr.GetAttribute("AlternateID")
                sValAltId = sAltId

        If Not sAltId Is Nothing Then
          sAltIdTableCol = MyCommon.Fetch_SystemOption(60)
          If sAltIdTableCol.Length > 0 Then
            sarrAltIdTableCol = sAltIdTableCol.Split(".")
            If sarrAltIdTableCol.Length = 2 Then
              sAltIdTable = sarrAltIdTableCol(0).ToLower.Trim
                            sAltIdCol = sarrAltIdTableCol(1).ToLower.Trim
                            If sAltIdTableCol = "CustomerExt.PhoneAsEntered" Or sAltIdTableCol = "CustomerExt.PhoneDigitsOnly" Or sAltIdTableCol = "CustomerExt.email" Or sAltIdTableCol = "CustomerExt.DateOfBirth" Or sAltIdTableCol = "CustomerExt.DOB" Or sAltIdTableCol = "CustomerExt.MobilePhoneAsEntered" Or sAltIdTableCol = "CustomerExt.MobilePhoneDigitsOnly" Or sAltIdTableCol = "CustomerExt.TaxExemptID" Then
                                sValAltId = MyCryptLib.SQL_StringEncrypt(sAltId)
                            End If
              Try
                                MyCommon.QueryStr = "select Customers.CustomerPK as CustomerPK from Customers with (NoLock)" & _
                                                    " left join CustomerExt with (NoLock) on CustomerExt.CustomerPK=Customers.CustomerPK" & _
                                                    " where CustomerTypeID=0 and " & sAltIdTableCol & "='" & sValAltId & "';"
                dt = MyCommon.LXS_Select()
                If (dt.Rows.Count > 0) Then
                  lOldCustomerPK = dt.Rows(0).Item("CustomerPK")
                  If lCustomerPk <> lOldCustomerPK Then
                    WriteLog("Cannot update Alternate ID because Alternate ID '" & sAltId & "' is in use by another customer (CustomerPK: " & lOldCustomerPK & ")!", MessageType.Warning)
                  End If
                Else
                  If sAltIdTableCol = "CustomerExt.PhoneDigitsOnly" And sAltId.Length = 10 Then
                    Dim sPhone As String = "(" & Mid(sAltId, 1, 3) & ") " & Mid(sAltId, 4, 3) & "-" & Mid(sAltId, 7, 4)
                                        MyCommon.QueryStr = "update " & sAltIdTable & " with (RowLock) set PhoneAsEntered = '" & MyCryptLib.SQL_StringEncrypt(sPhone) & _
                                        "', PhoneDigitsOnly = '" & sValAltId & "' where CustomerPK=" & lCustomerPk & ";"
                    MyCommon.LXS_Execute()
                    If MyCommon.RowsAffected = 0 Then
                                            MyCommon.QueryStr = "insert into " & sAltIdTable & " with (RowLock) (CustomerPk,PhoneAsEntered,PhoneDigitsOnly) values (" & lCustomerPk & ",'" & MyCryptLib.SQL_StringEncrypt(sPhone) & "','" & sValAltId & "');"
                      MyCommon.LXS_Execute()
                    End If
                  Else
                                        MyCommon.QueryStr = "update " & sAltIdTable & " with (RowLock) set " & sAltIdCol & " = '" & sValAltId & _
                                        "' where CustomerPK=" & lCustomerPk & ";"
                    MyCommon.LXS_Execute()
                    If MyCommon.RowsAffected = 0 Then
                                            MyCommon.QueryStr = "insert into " & sAltIdTable & " with (RowLock) (CustomerPk," & sAltIdCol & ") values (" & lCustomerPk & ",'" & sValAltId & "');"
                      MyCommon.LXS_Execute()
                    End If
                  End If
                  WriteDebug("Updated Alternate ID (" & sAltId & ") for customer " & Mask(sId) & ".", DebugState.CurrentTime)
                End If
              Catch ex As Exception
                WriteLog("Update of Alternate ID (" & sAltId & ") for customer " & Mask(sId) & " failed!", MessageType.Warning)
                WriteLog("Error: " & ex.Message, MessageType.Warning)
              End Try
            Else
              WriteLog("Cannot update Alternate ID because Invalid Alternate ID defined in options file!", MessageType.Warning)
            End If
          Else
            WriteLog("Cannot update Alternate ID because Alternate ID is NOT defined in options file!", MessageType.Warning)
          End If
        End If

        sSecondaryId = xr.GetAttribute("SecondaryID")
        If Not sSecondaryId Is Nothing Then
          Try
            sSecondaryId = MyCommon.Pad_ExtCardID(sSecondaryId, iIdTypeForSecondaryId)
            bCardFound = GetCustomerPK(sSecondaryId, iIdTypeForSecondaryId, lOldCustomerPK, lOldHhPk, False)
            If bCardFound Then
              If lOldCustomerPK = lCustomerPk Then
                ' already assigned to this customer
                Exit Try
              Else
                ' different customer has this card
                If bFobEligibilityEnabled And (iIdTypeForSecondaryId = 7) Then
                  MyCommon.QueryStr = "select CardPK from CardIDs with (NoLock)" & _
                                      " where CardTypeID=0 and CardStatusId=1 and CustomerPK=" & lOldCustomerPK & ";"
                  dt = MyCommon.LXS_Select()
                  If (dt.Rows.Count > 0) Then
                    ' other cusromer has an mPerks card
                    WriteLog("Cannot update FOB '" & sSecondaryId & "' because it is in use by another customer (CustomerPK: " & lOldCustomerPK & ") who has an active mPerks card!", MessageType.Warning)
                    Exit Try
                  Else
                    ' no active mPerks card, so steal the FOB
                    ' set status of existing Secondary cards to INACTIVE
                    MyCommon.QueryStr = "update CardIds with (RowLock) set CardStatusID=2 where CardStatusID=1 and CardTypeID=7 and CustomerPK=" & lCustomerPk & ";"
                    MyCommon.LXS_Execute()
                    ' tranfer this secondary card from original customer to this customer
                                        MyCommon.QueryStr = "update CardIDs with (RowLock) set CardStatusID=1, CustomerPK=" & lCustomerPk & _
                                                            " where CardTypeID=7 and CustomerPK=" & lOldCustomerPK & " and ExtCardID='" & MyCryptLib.SQL_StringEncrypt(sSecondaryId) & "';"
                    MyCommon.LXS_Execute()
                    WriteDebug("Transferred FOB '" & sSecondaryId & "' from CustomerPK '" & lOldCustomerPK & "' to CustomerPK '" & lCustomerPk & "'.", DebugState.CurrentTime)
                    Exit Try
                  End If
                Else
                  WriteLog("Update of Secondary ID (" & sSecondaryId & ") for customer " & Mask(sId) & " failed!", MessageType.Warning)
                  WriteLog("  It is in use by another customer (CustomerPK:  " & lOldCustomerPK & ").", MessageType.Warning)
                  Exit Try
                End If
              End If
            Else
              ' set status of existing Secondary cards to INACTIVE
              MyCommon.QueryStr = "update CardIds with (RowLock) set CardStatusID=2 where CardStatusID=1 and CardTypeID=" & iIdTypeForSecondaryId & " and CustomerPK=" & lCustomerPk & ";"
              MyCommon.LXS_Execute()
              ' add card to customer
                            MyCommon.QueryStr = "insert into CardIds with (RowLock) (CustomerPK, ExtCardID, CardStatusID, CardTypeID) values (" & lCustomerPk & ",'" & MyCryptLib.SQL_StringEncrypt(sSecondaryId) & "',1," & iIdTypeForSecondaryId & ");"
              MyCommon.LXS_Execute()
              WriteDebug("Added Secondary ID '" & sSecondaryId & "' to CustomerPK '" & lCustomerPk & "'.", DebugState.CurrentTime)
            End If
          Catch ex As Exception
            WriteLog("Update of Secondary ID (" & sSecondaryId & ") for customer " & Mask(sId) & " failed!", MessageType.Warning)
            WriteLog("Error: " & ex.Message, MessageType.Warning)
          End Try
        End If

        sCartID = xr.GetAttribute("PendingCartID")
        If Not sCartID Is Nothing Then
          Dim lTargetetLocationId As Long
          If MyCommon.Fetch_SystemOption(219) = "1" Then
            ' share data between CM & UE
            ' get Broker LocationID
            MyCommon.QueryStr = "select LocationID from Locations with (NoLock) where Deleted=0 and EngineID=9 and LocationTypeID=2;"
            dt = MyCommon.LRT_Select
            If (dt.Rows.Count > 0) Then
              lTargetetLocationId = dt.Rows(0).Item(0)
            End If
            If lTargetetLocationId = 0 Then
              WriteLog("Invalid LocationID was found for UE Broker, so the deletion of pending data for CartID '" & sCartID & "' was not sent to the broker!", MessageType.AppError)
            End If
          End If

          If lTargetetLocationId > 0 Then
            MyCommon.QueryStr = "dbo.pt_PendingDeleteByCartID_UpdateBroker"
          Else
            MyCommon.QueryStr = "dbo.pt_PendingDeleteByCartID"
          End If
          MyCommon.Open_LXSsp()
          MyCommon.LXSsp.Parameters.Add("@TableNum", SqlDbType.VarChar, 4).Value = ""
          MyCommon.LXSsp.Parameters.Add("@Operation", SqlDbType.VarChar, 2).Value = ""
          MyCommon.LXSsp.Parameters.Add("@Col1", SqlDbType.VarChar, 36).Value = sCartID
          MyCommon.LXSsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = 0
          MyCommon.LXSsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = lTargetetLocationId
          MyCommon.LXSsp.Parameters.Add("@WaitingACK", SqlDbType.Int).Value = 0
          MyCommon.LXSsp.ExecuteNonQuery()
          MyCommon.Close_LXSsp()
        End If

        MyCommon.QueryStr = "dbo.pt_TransHistory_Insert_Duplicate_Check"
        MyCommon.Open_LWHsp()
        MyCommon.LWHsp.Parameters.Add("@LogixTransNum", SqlDbType.VarChar, 36).Value = sLogixTransNum
        MyCommon.LWHsp.Parameters.Add("@CustomerPrimaryExtID", SqlDbType.NVarChar, 26).Value = sId
        MyCommon.LWHsp.Parameters.Add("@CustomerTypeID", SqlDbType.Int).Value = 0
        MyCommon.LWHsp.Parameters.Add("@ExtLocationCode", SqlDbType.NVarChar, 20).Value = sStoreNum
        MyCommon.LWHsp.Parameters.Add("@TransDate", SqlDbType.DateTime).Value = sDateTime
        MyCommon.LWHsp.Parameters.Add("@TerminalNum", SqlDbType.NVarChar, 4).Value = sTerminalNum
        MyCommon.LWHsp.Parameters.Add("@POSTransNum", SqlDbType.NVarChar,128).Value = sTransNum
        MyCommon.LWHsp.Parameters.Add("@IsDuplicate", SqlDbType.Bit).Direction = ParameterDirection.Output
        MyCommon.LWHsp.ExecuteNonQuery()
        bDuplicate = MyCommon.LWHsp.Parameters("@IsDuplicate").Value
        MyCommon.Close_LWHsp()

        If bDuplicate Then
          WriteLog("Skipping duplicate Loyalty transaction (Store: " & sStoreNum & ", Terminal: " & sTerminalNum & ", PosTransaction: " & sTransNum & _
                   ", CardID: " & sId & ", TransDate: " & sDateTime & ")", MessageType.Info)
        Else
        'Check for Promo that rewards Points
          xrSubTree = xr.ReadSubtree()
	xrSubtree.Read() 'Returns LoyaltyTransaction
	xrSubtree.Read() 'Returns next element: either Promos or PromoVars, if they exist
	sContent = xrSubtree.Name

	If (sContent).Contains("Promos") Then
              xrSubTree1 = xr.ReadSubtree()
              xrSubTree1.ReadToFollowing("Add")
              If Not xrSubTree1.EOF Then
                  If xrSubTree1.AttributeCount = 0 Then
                    Throw New ApplicationException("(LoyaltyTransaction) No attributes in 'Add' element.")
                  End If

                  sPromoId = xrSubTree1.GetAttribute("List")
                  If sPromoId Is Nothing OrElse sPromoId.Length = 0 Then
		    Throw New ApplicationException("(LoyaltyTransaction) No valid 'List' attribute in 'Add' element.")
                  Else
                    If (sPromoId like "*,") Then 
                      sPromoId = sPromoId.substring(0, (sPromoId.Length-1))
                    End If
                  End If

              End If
              xrSubTree1.Close()
              xrSubTree1 = Nothing

	  xrSubTree.Read()
	  sContent = xrSubtree.Name
          End If
          
	  
          If sContent.Contains("PromoVars") Then
            Do
              xrSubTree1 = xr.ReadSubtree()
              xrSubTree1.ReadToFollowing("Mod")
              If Not xrSubTree1.EOF Then
                Do
                  If xrSubTree1.AttributeCount = 0 Then
                    Throw New ApplicationException("(LoyaltyTransaction) No attributes in 'Mod' element.")
                  End If

                  sPromoVarId = xrSubTree1.GetAttribute("ID")
                  If sPromoVarId Is Nothing OrElse sPromoVarId.Length = 0 Then
                    Throw New ApplicationException("(LoyaltyTransaction) No valid 'ID' attribute in 'Mod' element.")
                  End If

                  sValue = xrSubTree1.GetAttribute("Val")
                  If sValue Is Nothing Then
                    Throw New ApplicationException("(LoyaltyTransaction) No 'Val' attribute in 'Mod' element.")
                  End If
                  If sValue.Length = 0 Then
                    sValue = "0"
                  End If

                  If bLifetime Then
                    sEarned = xrSubTree1.GetAttribute("Earned")
                    If sEarned Is Nothing Then
                      sEarned = ""
                    End If
                  Else
                    sEarned = ""
                  End If
                  
                If (sPromoId.Length>0) Then                   
                  MyCommon.QueryStr = "SELECT DISTINCT O.Name, O.OfferID FROM RewardPoints AS RP WITH (NoLock) " & _
                        "INNER JOIN OfferRewards AS OFFR WITH (NoLock) ON RP.RewardPointsID = OFFR.LinkID " & _
                        "AND (OFFR.RewardTypeID = 2 OR OFFR.RewardTypeID = 13) AND OFFR.Deleted=0 " & _
                        "INNER JOIN Offers AS O WITH (NoLock) ON OFFR.OfferID = O.OfferID AND O.Deleted = 0 WHERE RP.ProgramID = " & sPromoVarID & _
                        " and O.OfferID in (" & sPromoId & ");"

                  dt = MyCommon.LRT_Select
                  If dt.Rows.Count > 0 Then
                    sAwardName = MyCommon.NZ(dt.Rows(0).Item("Name"), "")
                  End If
		
		  End If
		  WriteLog("AwardName: " & sAwardName, MessageType.Info)
                  
                  UpdatePromoVars(lUseCustomerPk, sId, sPromoVarId, sValue, sEarned, sDateTime, sStoreNum, sTerminalNum, sTransNum, sLogixTransNum, sLocationId, sLocalServerId, sAwardName)
                Loop While xrSubTree1.ReadToNextSibling("Mod")
              End If
              xrSubTree1.Close()
              xrSubTree1 = Nothing
            Loop While xrSubTree.ReadToNextSibling("PromoVars")
          End If
          xrSubTree.Close()
          xrSubTree = Nothing

          If sId <> "0" Then
            If bUseGeneralCustomerLocking Then
              CustomerLockRelease(sPrepay, sId, iTrxCardTypeId, sStoreNum, sTerminalNum)
            End If
            ' only update customer once per "batch"
            If sId <> sPreviousId Then
              If Not MyCommon.AllowToProcessCustomerCard(sId, iTrxCardTypeId, cardValidationResp) Then
                Throw New ApplicationException(MyCommon.CardValidationResponseMessage(sId, iTrxCardTypeId, cardValidationResp))
                Exit Try
              End If
              UpdateCustomerCount(sId, iTrxCardTypeId)
              sPreviousId = sId
            End If
          End If
        End If
      Loop While xr.ReadToNextSibling("LoyaltyTransaction")
    Catch
      Throw
    Finally
      If Not dt Is Nothing Then
        dt.Dispose()
        dt = Nothing
      End If
      If Not xrSubTree Is Nothing Then
        xrSubTree.Close()
        xrSubTree = Nothing
      End If
      If Not xrSubTree1 Is Nothing Then
        xrSubTree1.Close()
        xrSubTree1 = Nothing
      End If
      WriteDebug("ProcessLoyaltyTransaction", DebugState.EndTime)
    End Try
  End Sub

  Private Sub ProcessStoredValueTransaction(ByRef xr As XmlReader, ByVal sStoreId As String)
    Dim sCustomerId As String
    Dim sDateTime As String
    Dim sStoreNum As String
    Dim sServerSerial As String
    Dim sTerminalNum As String
    Dim sTransNum As String
    Dim sLogixTransNum As String
    Dim sPrepay As String
    Dim sProgramId As String
    Dim sExtId As String
    Dim sStatus As String
    Dim sLocalID As String
    Dim sExpiration As String
    Dim sUnitValue As String
    Dim sQtyEarned As String
    Dim sQtyUsed As String
    Dim sTotalAmount As String
    Dim sBreakageAmount As String
    Dim sRedeemedAmount As String
    Dim sUnitLimit As String
    Dim sOfferId As String = "0"
    Dim decTemp As Decimal
    Dim sLocationId As String = "0"
    Dim sLocalServerId As String = "0"
    Dim sPreviousCustId As String = ""
    Dim dt As DataTable = Nothing
    Dim xrSubTree As XmlReader = Nothing
    Dim bNeedToBumpCounter As Boolean = False
    Dim bFuelPartnerProgram As Boolean = False
    Dim bUseSVCustomerLocking As Boolean
    Dim bAccumulateProgram As Boolean = False
    Dim bPointsProgram As Boolean = False
    Dim dValue As Decimal
    Dim iTrxCardTypeId As Integer

    sCurrentMethod = "ProcessStoredValueTransaction"
    WriteDebug(sCurrentMethod, DebugState.BeginTime)

    Try
      xr.ReadToFollowing("SVTransaction")
      If xr.EOF Then
        Throw New ApplicationException("No 'SVTransaction ' element")
      End If

      iTrxCardTypeId = GetTrxCardTypeId()

      ' check to see if general customer locking is on
      If MyCommon.Fetch_CM_SystemOption(36) = "1" Then
        ' yes, so unlock when processing Loyalty transaction type
        bUseSVCustomerLocking = False
      Else
        ' no, so unlock when processing Stored Values
        bUseSVCustomerLocking = True
      End If

      Do
        If xr.AttributeCount = 0 Then
          Throw New ApplicationException("No attributes in 'SVTransaction ' element.")
        End If

        If sStoreId.Length > 0 Then
          ' use store number from sending store (20 Alpha Numeric)
          sStoreNum = sStoreId
        Else
          ' use store number embedded in xml (unsigned short)
          sStoreNum = xr.GetAttribute("StoreNum")
        End If
        If sStoreNum Is Nothing OrElse sStoreNum.Length = 0 Then
          sStoreNum = "0"
        Else
          CheckForDuplicateMd5(sStoreNum)
          sCurrentMethod = "ProcessStoredValueTransaction"
        End If

        sCustomerId = xr.GetAttribute("ID")
        If sCustomerId.Length = 0 Then
          sCustomerId = "0"
        End If

        sDateTime = GetLocalDateTime(xr.GetAttribute("DateTime"))

        sTerminalNum = xr.GetAttribute("TerminalNum")
        If sTerminalNum Is Nothing OrElse sTerminalNum.Length = 0 Then
          sTerminalNum = "0"
        End If

        sTransNum = xr.GetAttribute("TransNum")
        If sTransNum Is Nothing OrElse sTransNum.Length = 0 Then
          sTransNum = "0"
        End If
        If sTransNum.Length >128 then
            Throw New ApplicationException("Transaction Number length of'" & sTransNum & "' should not be greater than 128 characters")
        End If
        sLogixTransNum = xr.GetAttribute("LogixTransNum")
        If sLogixTransNum Is Nothing OrElse sLogixTransNum.Length = 0 Then
          sLogixTransNum = "1"
        End If

        sPrepay = xr.GetAttribute("Prepay")
        If sPrepay Is Nothing OrElse sPrepay.Length = 0 Then
          sPrepay = "False"
        End If

        MyCommon.QueryStr = "dbo.pa_CM_Gen_CheckExtLocationCode"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@ExtLocationCode", SqlDbType.NVarChar).Value = sStoreNum
        dt = MyCommon.LRTsp_select
        MyCommon.Close_LRTsp()
        If dt.Rows.Count > 0 Then
          sLocationId = MyCommon.NZ(dt.Rows(0).Item("LocationID"), "0")
          sLocalServerId = MyCommon.NZ(dt.Rows(0).Item("LocalServerID"), "0")
        End If
        dt = Nothing

        If sLocationId = "0" OrElse sLocalServerId = "0" Then
          Throw New ApplicationException("LocationId or LocalServerID not found for ExtLocationCode: " & sStoreNum)
        End If

        xrSubTree = xr.ReadSubtree()
        xrSubTree.ReadToFollowing("StoredValue")
        If xrSubTree.EOF Then
          Throw New ApplicationException("(SVTransaction) No 'StoredValue' element")
        End If

        Do
          If xrSubTree.AttributeCount = 0 Then
            Throw New ApplicationException("No attributes in 'StoredValue' element.")
          End If

          sStatus = xrSubTree.GetAttribute("Status")
          If sStatus Is Nothing OrElse sStatus.Length = 0 Then
            Throw New ApplicationException("(SVTransaction) No valid 'Status' attribute in 'StoredValue' element.")
          End If

          sExtId = xrSubTree.GetAttribute("ExternalID")
          If sExtId Is Nothing OrElse sExtId.Length = 0 Then
            Throw New ApplicationException("(SVTransaction) No valid 'ExternalID' attribute in 'StoredValue' element.")
          End If

          ' skip processing for Released or Claimed
          If sStatus <> "Released" AndAlso sStatus <> "Claimed" Then
            sLocalID = xrSubTree.GetAttribute("LocalID")
            If sLocalID Is Nothing OrElse sLocalID.Length = 0 Then
              Throw New ApplicationException("(SVTransaction) No valid 'LocalID' attribute in 'StoredValue' element.")
            End If

            sProgramId = xrSubTree.GetAttribute("ProgramID")
            If sProgramId Is Nothing OrElse sProgramId.Length = 0 Then
              Throw New ApplicationException("(SVTransaction) No valid 'ProgramID' attribute in 'StoredValue' element.")
            End If

            sExpiration = xrSubTree.GetAttribute("Expiration")
            If sExpiration Is Nothing OrElse sExpiration.Length = 0 Then
              Throw New ApplicationException("(SVTransaction) No valid 'Expiration' attribute in 'StoredValue' element.")
            End If
            sExpiration = GetLocalDateTime(sExpiration)



            sUnitLimit = xrSubTree.GetAttribute("UnitLimit")
            If sUnitLimit Is Nothing OrElse sUnitLimit.Length = 0 Then
              sUnitLimit = "0"
            End If

            GetSVProgramDetails(sProgramId, bAccumulateProgram, bPointsProgram, bFuelPartnerProgram, dValue)
            sUnitValue = dValue.ToString

            If sStatus = "Issued" OrElse sStatus = "Reissued" Then
              sQtyEarned = xrSubTree.GetAttribute("QuantityEarned")
              If sQtyEarned Is Nothing OrElse sQtyEarned.Length = 0 Then
                Throw New ApplicationException("(SVTransaction) No valid 'QuantityEarned' attribute in 'StoredValue' element.")
              End If
            Else
              sQtyEarned = "0"
            End If

            If sStatus = "Issued" OrElse sStatus = "Reissued" Then
              sQtyUsed = "0"
            Else
              sQtyUsed = xrSubTree.GetAttribute("QuantityUsed")
              If sQtyUsed Is Nothing OrElse sQtyUsed.Length = 0 Then
                ' currently CM is NOT returning this value
                sQtyUsed = xrSubTree.GetAttribute("QuantityEarned")
                If sQtyUsed Is Nothing OrElse sQtyUsed.Length = 0 Then
                  Throw New ApplicationException("(SVTransaction) No valid 'QuantityEarned' attribute in 'StoredValue' element for CM redemption.")
                End If
              End If
            End If

            sTotalAmount = xrSubTree.GetAttribute("ComputedTotalAmount")
            If sTotalAmount Is Nothing OrElse sTotalAmount.Length = 0 Then
              sTotalAmount = "0.00"
            End If
            decTemp = Decimal.Parse(sTotalAmount)
            If Not bPointsProgram Then
              decTemp *= 0.01
            End If
            sTotalAmount = decTemp.ToString

            sBreakageAmount = xrSubTree.GetAttribute("ComputedBreakageAmount")
            If sBreakageAmount Is Nothing OrElse sBreakageAmount.Length = 0 Then
              sBreakageAmount = "0.00"
            End If
            decTemp = Decimal.Parse(sBreakageAmount)
            If Not bPointsProgram Then
              decTemp *= 0.01
            End If
            sBreakageAmount = decTemp.ToString

            sRedeemedAmount = xrSubTree.GetAttribute("ComputedRedeemedAmount")
            If sRedeemedAmount Is Nothing OrElse sRedeemedAmount.Length = 0 Then
              sRedeemedAmount = "0.00"
            End If
            decTemp = Decimal.Parse(sRedeemedAmount)
            If Not bPointsProgram Then
              decTemp *= 0.01
            End If
            sRedeemedAmount = decTemp.ToString

            If sStatus = "Issued" Then
              sServerSerial = "0"
            Else
              sServerSerial = xrSubTree.GetAttribute("ServerSerial")
              If sServerSerial Is Nothing OrElse sServerSerial.Length = 0 Then
                Throw New ApplicationException("(SVTransaction) No valid 'sServerSerial' attribute in 'StoredValue' element.")
              End If
            End If

            If bFuelPartnerProgram Then
              ' update Fuel Partner SV table in EX database
              ' CmFuelPartnerAgent will subsequently update Stored Value table
              UpdateFuelPartnerSV(sStatus, sCustomerId, sStoreNum, sLocalID, sServerSerial, sExtId, _
                               sUnitValue, sExpiration, sQtyEarned, sQtyUsed, sTotalAmount, sBreakageAmount, _
                               sRedeemedAmount, sDateTime, sProgramId, sOfferId, sLocationId, sLocalServerId, sLogixTransNum)
            Else
              If Not sCustomerId = "0" Then
                If Not MyCommon.AllowToProcessCustomerCard(sCustomerId, iTrxCardTypeId, cardValidationResp) Then
                  Throw New ApplicationException(MyCommon.CardValidationResponseMessage(sCustomerId, iTrxCardTypeId, cardValidationResp))
                  Exit Try
                End If
              End If
              UpdateStoreValue(sStatus, sCustomerId, iTrxCardTypeId, sStoreNum, sLocalID, sServerSerial, sExtId, _
                               sUnitValue, sExpiration, sQtyEarned, sQtyUsed, sTotalAmount, sBreakageAmount, _
                               sRedeemedAmount, sDateTime, sProgramId, sOfferId, sLocationId, sLocalServerId, _
                               sLogixTransNum, bAccumulateProgram, bPointsProgram)
              bNeedToBumpCounter = True
            End If
          ElseIf sStatus = "Released" AndAlso MyCommon.Fetch_CM_SystemOption(51) = "True" Then
            'Also update the counter if the status is "Released" And the customer is Brookshire's
            bNeedToBumpCounter = True
          End If
          If sCustomerId = "0" Then
            StoredValueLockRelease(sExtId)
          End If
        Loop While xrSubTree.ReadToNextSibling("StoredValue")
        xrSubTree.Close()
        xrSubTree = Nothing
        If sCustomerId <> "0" Then
          If bUseSVCustomerLocking Then
            CustomerLockRelease(sPrepay, sCustomerId, iTrxCardTypeId, sStoreNum, sTerminalNum)
          End If
          ' only update customer once per "batch"
          If bNeedToBumpCounter AndAlso sCustomerId <> sPreviousCustId Then
            If Not MyCommon.AllowToProcessCustomerCard(sCustomerId, iTrxCardTypeId, cardValidationResp) Then
              Throw New ApplicationException(MyCommon.CardValidationResponseMessage(sCustomerId, iTrxCardTypeId, cardValidationResp))
              Exit Try
            End If
            UpdateCustomerCount(sCustomerId, iTrxCardTypeId)
            sPreviousCustId = sCustomerId
          End If
        End If
      Loop While xr.ReadToNextSibling("SVTransaction")
    Catch
      Throw
    Finally
      If Not dt Is Nothing Then
        dt.Dispose()
        dt = Nothing
      End If
      If Not xrSubTree Is Nothing Then
        xrSubTree.Close()
        xrSubTree = Nothing
      End If
      WriteDebug("ProcessStoredValueTransaction", DebugState.EndTime)
    End Try
  End Sub

  Private Sub GetSVProgramDetails(ByVal sProgramId As String, ByRef bAccumulateProgram As Boolean, ByRef bPointsProgram As Boolean, ByRef bFuelPartnerProgram As Boolean, ByRef dValue As Decimal)
    Dim dt As DataTable
    Dim iSvTypeID As Integer

    bAccumulateProgram = False
    bPointsProgram = False
    bFuelPartnerProgram = False
    dValue = 1.0

    MyCommon.QueryStr = "select SVTypeID, SVExpireType, SVExpirePeriodType, ExpirePeriod, Value, FuelPartner " & _
                        "from StoredValuePrograms with (NoLock) " & _
                        "where SVProgramID=" & sProgramId & " and Deleted=0;"
    dt = MyCommon.LRT_Select
    If (dt.Rows.Count > 0) Then
      dValue = MyCommon.NZ(dt.Rows(0).Item("Value"), 0)
      iSvTypeID = MyCommon.NZ(dt.Rows(0).Item("SVTypeID"), 0)
      Select Case iSvTypeID
        Case 1
          bPointsProgram = True
          ' is it Expire X months after end of current month
          If MyCommon.NZ(dt.Rows(0).Item("SVExpireType"), 0) = 5 Then
            ' is period in months
            If MyCommon.NZ(dt.Rows(0).Item("SVExpirePeriodType"), 0) = 3 Then
              bAccumulateProgram = True
            End If
          End If
        Case 3
          bFuelPartnerProgram = MyCommon.NZ(dt.Rows(0).Item("FuelPartner"), False)
      End Select
    Else
      Throw New ApplicationException("Invalid Stored Value Program ID: '" & sProgramId & "'")
    End If

  End Sub

  Private Sub ProcessCustGroupTransaction(ByRef xr As XmlReader, ByVal sStoreId As String)
    Dim sOp As String
    Dim sId As String
    Dim sIdPrevious As String = ""
    Dim sGroup As String
    Dim sDateTime As String
    Dim sStoreNum As String
    Dim sTerminalNum As String
    Dim sTransNum As String
    Dim sLogixTransNum As String
    Dim iTrxCardTypeId As Integer

    sCurrentMethod = "ProcessCustGroupTransaction"
    xr.ReadToFollowing("CustGroupTransaction")
    If xr.EOF Then
      Throw New ApplicationException("No 'CustGroupTransaction' element")
    End If

    iTrxCardTypeId = GetTrxCardTypeId()

    Do
      If xr.AttributeCount = 0 Then
        Throw New ApplicationException("No attributes in 'CustGroupTransaction' element.")
      End If

      sDateTime = GetLocalDateTime(xr.GetAttribute("DateTime"))

      If sStoreId.Length > 0 Then
        ' use store number from sending store (20 Alpha Numeric)
        sStoreNum = sStoreId
      Else
        ' use store number embedded in xml (unsigned short)
        sStoreNum = xr.GetAttribute("StoreNum")
      End If

      If sStoreNum Is Nothing OrElse sStoreNum.Length = 0 Then
        sStoreNum = "0"
      Else
        CheckForDuplicateMd5(sStoreNum)
        sCurrentMethod = "ProcessCustGroupTransaction"
      End If

      sTerminalNum = xr.GetAttribute("TerminalNum")
      If sTerminalNum Is Nothing OrElse sTerminalNum.Length = 0 Then
        sTerminalNum = "0"
      End If

      sTransNum = xr.GetAttribute("TransNum")
      If sTransNum Is Nothing OrElse sTerminalNum.Length = 0 Then
        sTransNum = "0"
      End If
        If sTransNum.Length >128 then
            Throw New ApplicationException("Transaction Number length of'" & sTransNum & "' should not be greater than 128 characters")
        End If
      sLogixTransNum = xr.GetAttribute("LogixTransNum")
      If sLogixTransNum Is Nothing OrElse sLogixTransNum.Length = 0 Then
        sLogixTransNum = "1"
      End If

      xr.ReadToFollowing("CustGroup")
      If xr.EOF Then
        Throw New ApplicationException("(CustGroupTransaction) No 'CustGroup' element")
      End If

      Do
        If xr.AttributeCount = 0 Then
          Throw New ApplicationException("(CustGroupTransaction) No attributes in 'CustGroup' element.")
        End If

        sOp = xr.GetAttribute("Type").ToUpper
        If sOp Is Nothing OrElse sOp.Length = 0 Then
          Throw New ApplicationException("(CustGroupTransaction) No valid 'Type' attribute in 'CustGroup' element.")
        End If

        sId = xr.GetAttribute("ID")

        If sId Is Nothing OrElse sId.Length = 0 Then
          Throw New ApplicationException("(CustGroupTransaction) No valid 'ID' attribute in 'CustGroup' element.")
        End If

        sId = MyCommon.Pad_ExtCardID(sId, iTrxCardTypeId)

        sGroup = xr.GetAttribute("Group")
        If sGroup Is Nothing OrElse sGroup.Length = 0 Then
          Throw New ApplicationException("(CustGroupTransaction) No 'Group' attribute in 'CustGroup' element.")
        End If
        If Not MyCommon.AllowToProcessCustomerCard(sId, iTrxCardTypeId, cardValidationResp) Then
          Throw New ApplicationException(MyCommon.CardValidationResponseMessage(sId, iTrxCardTypeId, cardValidationResp))
          Exit Sub
        End If
        UpdateCustGroup(sOp, sId, iTrxCardTypeId, sGroup, sDateTime, sStoreNum, sTerminalNum, sTransNum, sLogixTransNum)
        If sId <> sIdPrevious Then
          ' only update customer once per "batch"
          ' maybe the Customer Id should be at the "CustGroupTransaction" element level?
          UpdateCustomerCount(sId, iTrxCardTypeId)
          sIdPrevious = sId
        End If
      Loop While xr.ReadToNextSibling("CustGroup")
    Loop While xr.ReadToNextSibling("CustGroupTransaction")
  End Sub

  Private Sub ProcessPromoMovtTransaction(ByRef xr As XmlReader, ByVal sStoreId As String)
    Dim sId As String
    Dim sCount As String
    Dim sAmount As String
    Dim sDateTime As String
    Dim sStoreNum As String
    Dim sTerminalNum As String
    Dim sTransNum As String

    sCurrentMethod = "ProcessPromoMovtTransaction"
    WriteDebug(sCurrentMethod, DebugState.BeginTime)
    Try
      xr.ReadToFollowing("PromoMovtTransaction")
      If xr.EOF Then
        Throw New ApplicationException("No 'PromoMovtTransaction' element")
      End If
      Do
        If xr.AttributeCount = 0 Then
          Throw New ApplicationException("No attributes in 'PromoMovtTransaction' element.")
        End If

        sDateTime = GetLocalDateTime(xr.GetAttribute("DateTime"))

        If sStoreId.Length > 0 Then
          ' use store number from sending store (20 Alpha Numeric)
          sStoreNum = sStoreId
        Else
          ' use store number embedded in xml (unsigned short)
          sStoreNum = xr.GetAttribute("StoreNum")
        End If
        If sStoreNum Is Nothing OrElse sStoreNum.Length = 0 Then
          sStoreNum = "0"
        Else
          CheckForDuplicateMd5(sStoreNum)
          sCurrentMethod = "ProcessPromoMovtTransaction"
        End If
        sTerminalNum = xr.GetAttribute("TerminalNum")
        If sTerminalNum Is Nothing OrElse sTerminalNum.Length = 0 Then
          sTerminalNum = "0"
        End If
        sTransNum = xr.GetAttribute("TransNum")
        If sTransNum Is Nothing OrElse sTerminalNum.Length = 0 Then
          sTransNum = "0"
        End If
        If sTransNum.Length >128 then
            Throw New ApplicationException("Transaction Number length of'" & sTransNum & "' should not be greater than 128 characters")
        End If
        xr.ReadToFollowing("PromoMovt")
        If xr.EOF Then
          Throw New ApplicationException("(PromoMovtTransaction) No 'PromoMovt' element")
        End If

        Do
          If xr.AttributeCount = 0 Then
            Throw New ApplicationException("No attributes in 'PromoMovt' element.")
          End If

          sId = xr.GetAttribute("ID")
          If sId Is Nothing OrElse sId.Length = 0 Then
            Throw New ApplicationException("(PromoMovtTransaction) No valid 'ID' attribute in 'PromoMovt' element.")
          End If

          sCount = xr.GetAttribute("Count").ToUpper
          If sCount Is Nothing OrElse sCount.Length = 0 Then
            Throw New ApplicationException("(PromoMovtTransaction) No valid 'Count' attribute in 'PromoMovt' element.")
          End If

          sAmount = xr.GetAttribute("Amount")
          If sAmount Is Nothing OrElse sAmount.Length = 0 Then
            sAmount = "0"
          End If

          UpdatePromoMovt(sId, sCount, sAmount, sDateTime, sStoreNum, sTerminalNum, sTransNum)
        Loop While xr.ReadToNextSibling("PromoMovt")
      Loop While xr.ReadToNextSibling("PromoMovtTransaction")
    Catch
      Throw
    Finally
      WriteDebug("ProcessPromoMovtTransaction", DebugState.EndTime)
    End Try
  End Sub

  Private Sub PromoSummaryTransaction(ByVal xr As XmlReader, ByVal sStoreId As String)
    Dim sId As String
    Dim sCount As String
    Dim sAmount As String
    Dim sDateTime As String
    Dim sStoreNum As String
    Dim dtTemp As Date
    Const sDateOnlyFormat As String = "yyyy-MM-ddT00:00:00"

    sCurrentMethod = "PromoSummaryTransaction"
    WriteDebug(sCurrentMethod, DebugState.BeginTime)
    Try
      xr.ReadToFollowing("PromoMovtTransaction")
      If xr.EOF Then
        Throw New ApplicationException("No 'PromoMovtTransaction' element")
      End If
      Do
        If xr.AttributeCount = 0 Then
          Throw New ApplicationException("No attributes in 'PromoMovtTransaction' element.")
        End If

        sDateTime = GetLocalDateTime(xr.GetAttribute("DateTime"))
        dtTemp = Date.Parse(sDateTime)
        sDateTime = Format(dtTemp, sDateOnlyFormat)

        If sStoreId.Length > 0 Then
          ' use store number from sending store (20 Alpha Numeric)
          sStoreNum = sStoreId
        Else
          ' use store number embedded in xml (unsigned short)
          sStoreNum = xr.GetAttribute("StoreNum")
        End If
        If sStoreNum Is Nothing OrElse sStoreNum.Length = 0 Then
          sStoreNum = "0"
        Else
          CheckForDuplicateMd5(sStoreNum)
          sCurrentMethod = "PromoSummaryTransaction"
        End If

        xr.ReadToFollowing("PromoMovt")
        If xr.EOF Then
          Throw New ApplicationException("(PromoMovtTransaction) No 'PromoMovt' element")
        End If

        Do
          If xr.AttributeCount = 0 Then
            Throw New ApplicationException("No attributes in 'PromoMovt' element.")
          End If

          sId = xr.GetAttribute("ID")
          If sId Is Nothing OrElse sId.Length = 0 Then
            Throw New ApplicationException("(PromoMovtTransaction) No valid 'ID' attribute in 'PromoMovt' element.")
          End If

          sCount = xr.GetAttribute("Count").ToUpper
          If sCount Is Nothing OrElse sCount.Length = 0 Then
            Throw New ApplicationException("(PromoMovtTransaction) No valid 'Count' attribute in 'PromoMovt' element.")
          End If

          sAmount = xr.GetAttribute("Amount")
          If sAmount Is Nothing OrElse sAmount.Length = 0 Then
            sAmount = "0"
          End If

          UpdatePromoSummary(sId, sCount, sAmount, sDateTime, sStoreNum)
        Loop While xr.ReadToNextSibling("PromoMovt")
      Loop While xr.ReadToNextSibling("PromoMovtTransaction")
    Catch
      Throw
    Finally
      WriteDebug("PromoSummaryTransaction", DebugState.EndTime)
    End Try
  End Sub

  Private Sub ProcessPromoBulkTransaction(ByVal sTransactionXml As String, ByVal sStoreId As String)
    Dim data As String = ""
    Dim sDateTime As String
    Dim sFileName As String
    Dim sFilePath As String
    Dim iFileType As Integer = 1
    Dim iVersionPosition As Integer
    Dim sVersionNumber As String = "Version=""2.0"""

    sCurrentMethod = "ProcessPromoBulkTransaction"

    ' check for data...if none, then we're done.
    If (sTransactionXml.Length = 0) Then
      Exit Sub
    End If

    WriteDebug(sCurrentMethod, DebugState.BeginTime)
    Try
      'format the date
      sDateTime = Format(Date.Now, "yyyy-MM-ddTHH:mm:ss")
      sDateTime = sDateTime.Replace("T", "")
      sDateTime = sDateTime.Replace("-", "")
      sDateTime = sDateTime.Replace(":", "")

      If (sDateTime.Length > 14) Then
        sDateTime = Left(sDateTime, 14)
      End If
      sDateTime = sDateTime & Date.Now.Millisecond.ToString

      If sStoreId Is Nothing OrElse sStoreId.Length = 0 Then
        sStoreId = "0"
      Else
        CheckForDuplicateMd5(sStoreId) ' comment this line out during testing
        sCurrentMethod = "ProcessPromoBulkTransaction"
      End If

      iVersionPosition = sTransactionXml.IndexOf(sVersionNumber)
      If iVersionPosition > 0 Then
        iFileType = 4
      Else
        iFileType = 1
      End If

      Dim startStr As String = "<![CDATA["
      Dim startPos As Integer = sTransactionXml.IndexOf(startStr) + startStr.Length
      Dim endStr As String = "]]>"
      Dim endPos As Integer = sTransactionXml.LastIndexOf(endStr)
      data = TrimComplete(sTransactionXml.Substring(startPos, (endPos - startPos)))
      ' check for data...if none, then we're done.
      If (data.Length = 0) Then
        Exit Sub
      End If

      If (data.IndexOf(vbCr) < 0) Then
        data = data.Replace(vbLf, vbCrLf) 'SOAP strips out the CR leaving only a LF
      End If

      'Write file
      sFilePath = MyCommon.Fetch_SystemOption(29)
      If Not (sFilePath.Substring(sFilePath.Length - 1, 1) = "\") Then
        sFilePath = sFilePath & "\"
      End If

      sFileName = "TH" & sStoreId & sDateTime & ".txt"
      WriteBulkInsertFile(data, sFilePath & sFileName)

      'Insert entry in Queue table
      MyCommon.QueryStr = "dbo.pa_PromoMoveQueue_Insert"
      MyCommon.Open_LWHsp()
      MyCommon.LWHsp.Parameters.Add("@FileName", SqlDbType.NVarChar, 255).Value = sFileName
      MyCommon.LWHsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = 0
      MyCommon.LWHsp.Parameters.Add("@FileType", SqlDbType.Int).Value = iFileType
      MyCommon.LWHsp.ExecuteNonQuery()
      MyCommon.Close_LWHsp()
    Catch
      Throw
    Finally
      WriteDebug("ProcessPromoBulkTransaction", DebugState.EndTime)
    End Try

  End Sub

  Public Function TrimComplete(ByVal sValue As String) As String
    Dim sAns As String
    Dim sChar As String
    Dim lLen As Long
    Dim lCtr As Long

    sAns = sValue
    lLen = Len(sValue)

    If lLen > 0 Then
      'Ltrim
      For lCtr = 1 To lLen
        sChar = Mid(sAns, lCtr, 1)
        If Asc(sChar) > 32 Then Exit For
      Next

      sAns = Mid(sAns, lCtr)
      lLen = Len(sAns)

      'Rtrim
      If lLen > 0 Then
        For lCtr = lLen To 1 Step -1
          sChar = Mid(sAns, lCtr, 1)
          If Asc(sChar) > 32 Then Exit For
        Next
      End If
      sAns = Left$(sAns, lCtr)
    End If

    TrimComplete = sAns

  End Function

  Private Sub WriteBulkInsertFile(ByVal data As String, ByVal fileName As String)
    Dim sw As IO.StreamWriter = Nothing

    Try
      sw = New StreamWriter(fileName)
      sw.Write(data)
    Catch ex As Exception
      Throw New ApplicationException("Unable to write Bulk Insert File " & fileName)
    Finally
      If (Not sw Is Nothing) Then
        sw.Close()
        sw = Nothing
      End If
    End Try
  End Sub

  Private Sub UpdatePromoVars(ByVal lUseCustomerPk As Long, ByVal sExtCustomerID As String, ByVal sPromoVarId As String, ByVal sValue As String, ByVal sEarned As String, ByVal sDateTime As String, ByVal sStoreNum As String, ByVal sTerminalNum As String, ByVal sTransNum As String, ByVal sLogixTransNum As String, ByVal sLocationId As String, ByVal sLocalServerId As String, Optional ByVal sAwardName As String = "")
    Dim iStatus As Integer = 0
    Dim dt As DataTable
    Dim sExternalId As String = ""
    Dim sProgramName As String = ""
    Dim lEarned As Long
    Dim decValue As Decimal
    Dim iVarTypeID As Integer = 0
    Dim sIncentiveID As String = ""
    Dim sRewardOptionID As String = ""
    Dim bUpdateExternalPoints As Boolean = False
    Dim bPointsPromoVar As Boolean = False
    Dim bOptionalParams As Boolean = (MyCommon.Fetch_CM_SystemOption(109) = "1")
    Dim sLocationCode As String = ""
    Dim bDistributionPromoVar As Boolean = False
    
    MyCryptLib = New Copient.CryptLib

    sCurrentMethod = "UpdatePromoVars"
    Try

      ' check if this is a points program
      MyCommon.QueryStr = "select PromoVarID, VarTypeID, ExternalID, LinkID from PromoVariables with (NoLock) where Deleted=0 and PromoVarID=" & IIf(sPromoVarId = "", 0, sPromoVarId)
      dt = MyCommon.LXS_Select
      If (dt.Rows.Count > 0) Then
        iVarTypeID = MyCommon.NZ(dt.Rows(0).Item("VarTypeID"), 0)
        If (iVarTypeID = 3) Then
          bPointsPromoVar = True
          ' check if EME is setup
          If MyCommon.Fetch_SystemOption(80) = "1" Then
            sExternalId = MyCommon.NZ(dt.Rows(0).Item("ExternalID"), "")
            ' check if this is an external points program
            MyCommon.QueryStr = "select ProgramID, ProgramName from PointsPrograms with (NoLock) where ExternalProgram=1 and Deleted=0 and PromoVarID=" & IIf(sPromoVarId = "", 0, sPromoVarId)
            dt = MyCommon.LRT_Select
            If (dt.Rows.Count > 0) Then
              bUpdateExternalPoints = True
              sProgramName = MyCommon.NZ(dt.Rows(0).Item("ProgramName"), "")
            End If
          End If
        End If
      
        ' if we're sharing distribution data from CM to UE
        If MyCommon.Fetch_SystemOption(219) = "1" Then
          ' check to see if this promovar is a distribution type variable
          If (iVarTypeID = 1) Then
            ' get the Reward Option ID that tracks distribution for the UE version of the offer
            MyCommon.QueryStr = "select I.IncentiveID, RO.RewardOptionID from CPE_RewardOptions as RO join CPE_Incentives as I on RO.IncentiveID=I.IncentiveID where I.Deleted=0 and RO.Deleted=0 and I.ClientOfferID=" & MyCommon.NZ(dt.Rows(0).Item("LinkID"), "0")
            dt = MyCommon.LRT_Select
            If (dt.Rows.Count > 0) Then
              sIncentiveID = MyCommon.NZ(dt.Rows(0).Item("IncentiveID"), "")
              sRewardOptionID = MyCommon.NZ(dt.Rows(0).Item("RewardOptionID"), "")
              bDistributionPromoVar = True
            Else
              WriteDebug("Distribution PromoVar found without UE equivalent " & sPromoVarID, DebugState.CurrentTime)
            End If
          End If
        End If
      End If

      If (bOptionalParams) Then
        MyCommon.QueryStr = "select ExtLocationCode from Locations where LocationID=" & sLocationID
        dt = MyCommon.LRT_Select
        If (dt.Rows.Count > 0) Then
            sLocationCode = MyCommon.NZ(dt.Rows(0).Item("ExtLocationCode"), "")
        End If
        
      End If

      If bUpdateExternalPoints Then
        ' only send updates for external points programs
        WriteDebug("ExternalPointsUpdate", DebugState.BeginTime)
        Try
          Dim ExternalPP As Copient.ExternalRewards
          sDb2Connection = MyCommon.Fetch_CM_SystemOption(iDb2Connection)
          tempstr = sDb2Connection
          upos = InStr(tempstr, "UID=", CompareMethod.Text)
          ppos = InStr(tempstr, ";PWD=", CompareMethod.Text)
          pend = InStr(tempstr, ";host", CompareMethod.Text)
          euser = tempstr.Substring(upos+3, ppos-upos-4)
          epwd = tempstr.Substring(ppos+4, pend-ppos-5)
          user =  MyCryptLib.SQL_StringDecrypt(euser)
          pwd =  MyCryptLib.SQL_StringDecrypt(epwd)
          tempstr = tempstr.Replace(euser, user)
          sDb2Connection = tempstr.Replace(epwd, pwd)
          ExternalPP = New Copient.ExternalRewards("", "", "", sDb2Connection)
          If Not Decimal.TryParse(sValue, decValue) Then decValue = 0.0
          If (bOptionalParams) Then
            ExternalPP.updateExternalBalance(sExternalId, sProgramName, decValue, sExtCustomerID, lUseCustomerPk, sPromoVarId, sLocalServerId, sLocationId, sLogixTransNum, MyCommon, sTerminalNum, sDateTime, sTransNum, sLocationCode, sAwardName, 0)
          Else
            ExternalPP.updateExternalBalance(sExternalId, sProgramName, decValue, sExtCustomerID, lUseCustomerPk, sPromoVarId, sLocalServerId, sLocationId, sLogixTransNum, MyCommon)
          End If
        Catch
          Throw
        Finally
          WriteDebug("ExternalPointsUpdate", DebugState.EndTime)
        End Try
      Else
        MyCommon.QueryStr = "dbo.pa_ServicePromoVarAmt_Update"
        If bDistributionPromoVar Then
          MyCommon.QueryStr = "dbo.pa_ServicePromoVarAndRewardDistribution"
        End If
        MyCommon.Open_LXSsp()
        MyCommon.LXSsp.Parameters.Add("@UseCustomerPK", SqlDbType.BigInt).Value = lUseCustomerPk
        MyCommon.LXSsp.Parameters.Add("@PromoVarID", SqlDbType.BigInt).Value = sPromoVarId
        MyCommon.LXSsp.Parameters.Add("@Amount", SqlDbType.Decimal, 9).Value = sValue
        MyCommon.LXSsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = sLocalServerId
        MyCommon.LXSsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = sLocationId
        MyCommon.LXSsp.Parameters.Add("@LogixTransNum", SqlDbType.Char, 36).Value = sLogixTransNum
        If bDistributionPromoVar Then
          MyCommon.LXSsp.Parameters.Add("@TransDate", SqlDbType.DateTime).Value = sDateTime
          MyCommon.LXSsp.Parameters.Add("@ExtCustomerID", SqlDbType.NVarChar, 26).Value = sExtCustomerID
          MyCommon.LXSsp.Parameters.Add("@IncentiveID", SqlDbType.BigInt).Value = sIncentiveID
          MyCommon.LXSsp.Parameters.Add("@RewardOptionID", SqlDbType.BigInt).Value = sRewardOptionID
        End If
        MyCommon.LXSsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
        MyCommon.LXSsp.ExecuteNonQuery()
        iStatus = MyCommon.LXSsp.Parameters("@Status").Value
        MyCommon.Close_LXSsp()

        If iStatus > 0 Then
          Throw New ApplicationException("Invalid Promotion Variable Type: " & iStatus & " for PromoVarId: " & sPromoVarId)
        ElseIf iStatus = -1 Then
          Throw New ApplicationException("Promotion Variable Id: " & sPromoVarId & " NOT found.")
        ElseIf iStatus < -1 Then
          Throw New ApplicationException("Bad Status returned from stored procedure 'pa_ServicePromoVarAmt_Update'")
        End If

        If bPointsPromoVar And sEarned <> "" Then
          If Not Long.TryParse(sEarned, lEarned) Then lEarned = 0
          If lEarned > 0 Then
            MyCommon.QueryStr = "update CM_Points_Lifetime with (RowLock) set Amount = Amount + " & lEarned & _
                                " where CustomerPK=" & lUseCustomerPk & " and PromoVarID=" & sPromoVarId & ";"
            MyCommon.LXS_Execute()
            If MyCommon.RowsAffected = 0 Then
              MyCommon.QueryStr = "insert into CM_Points_Lifetime with (RowLock) (CustomerPK,PromoVarID,Amount) values (" & lUseCustomerPk & _
                                   "," & sPromoVarId & "," & lEarned & ");"
              MyCommon.LXS_Execute()
            End If
          End If
        End If
      End If

    Catch exApp As ApplicationException
      Throw
    Catch ex As Exception
      Throw
    End Try
  End Sub


  Private Sub UpdateStoreValue(ByVal sStatus As String, ByVal sCardId As String, ByVal iTrxCardTypeId As Integer, ByVal sExtLocationCode As String, _
                               ByVal sLocalID As String, ByVal sServerSerial As String, ByVal sExtId As String, _
                               ByVal sUnitValue As String, ByVal sExpiration As String, ByVal sQtyEarned As String, _
                               ByVal sQtyUsed As String, ByVal sTotalAmount As String, ByVal sBreakageAmount As String, _
                               ByVal sRedeemedAmount As String, ByVal sDateTime As String, ByVal sProgramId As String, _
                               ByVal sOfferId As String, ByVal sLocationId As String, ByVal sLocalServerId As String, _
                               ByVal sLogixTransNum As String, ByVal bAccumulate As Boolean, ByVal bPointsProgram As Boolean)
    Dim iStatus As Integer = 0
    Dim iUpdateCount As Integer = 0
    Dim decTemp As Decimal
    Dim iTemp As Integer

    sCurrentMethod = "UpdateStoreValue"
    Try

      Select Case sStatus
        Case "Issued"
          iStatus = 1
          sServerSerial = sLocalServerId
        Case "Reissued"
          iStatus = 10
        Case "Redeemed"
          iStatus = 4
          If sServerSerial = "0" Then
            ' This code compensates for a problem with POS translating a ServerSerial value of "-9" to "0"
            ' The code simply puts the "-9" back in place, ASSUMING that a redemtion should never
            ' have SeverSerial = "0" and "0" only occurs now as a result of the bug at POS.
            ' The "-9" occurs when a SV is issued via central server.
            sServerSerial = "-9"
          ElseIf sServerSerial = "-1" Then
            ' This code compensates for a problem with redeeminmg Stored Value points that were issued in same POS transaction
            ' If reddeemed in same transaction as issued, POS sets ServerSerial for redemption to "-1"
            ' This code uses the LocalServerID for this store to set ServerSerial 
            sServerSerial = sLocalServerId
          End If
          ' This is a temporary fix for Restricted use coupons which are scanned but have already been redeemed
          ' They are currenty being sent as part of the PutTransaction xml as "Redeemed"
          ' At some point, the POS will be modified not to send them or send them as "Released"
          If sCardId = "0" And sQtyUsed = "0" Then
            WriteLog("Skipping already redeemed coupon (" & sLocalID & ", ServerSerial: " & sServerSerial & ")", MessageType.Info)
            Exit Sub
          End If
          If sCardId = "0" Then ' RUC (Once & Done)
            sQtyUsed = "1"
          Else
            If bPointsProgram Then
              If Not bAccumulate Then
                iStatus = -4
              End If
              If Not Decimal.TryParse(sRedeemedAmount, decTemp) Then decTemp = 0.0
              iTemp = decTemp
              sQtyUsed = iTemp.ToString
            End If
          End If
        Case "Invalid"
          iStatus = 2
        Case Else
          Throw New ApplicationException("Invalid status - '" & sStatus & "'!")
      End Select

      sDefaultFirstName = MyCommon.Fetch_CM_SystemOption(20)
      sDefaultLastName = MyCommon.Fetch_CM_SystemOption(21)

      If bAccumulate Then
        MyCommon.QueryStr = "dbo.pa_LogixServ_AccumulateStoredValues"
        MyCommon.Open_LXSsp()
        MyCommon.LXSsp.Parameters.Add("@StatusFlag", SqlDbType.Int).Value = iStatus
        MyCommon.LXSsp.Parameters.Add("@ExtCustomerID", SqlDbType.NVarChar, 26).Value = sCardId
        MyCommon.LXSsp.Parameters.Add("@CardTypeID", SqlDbType.Int).Value = iTrxCardTypeId
        MyCommon.LXSsp.Parameters.Add("@LocationId", SqlDbType.BigInt).Value = sLocationId
        MyCommon.LXSsp.Parameters.Add("@LocalID", SqlDbType.BigInt).Value = sLocalID
        MyCommon.LXSsp.Parameters.Add("@ServerSerial", SqlDbType.Int).Value = sServerSerial
        MyCommon.LXSsp.Parameters.Add("@ExternalID", SqlDbType.NVarChar, 30).Value = sExtId
        MyCommon.LXSsp.Parameters.Add("@SVProgramID", SqlDbType.BigInt).Value = sProgramId
        MyCommon.LXSsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = sOfferId
        MyCommon.LXSsp.Parameters.Add("@QtyEarned", SqlDbType.Int).Value = sQtyEarned
        MyCommon.LXSsp.Parameters.Add("@QtyUsed", SqlDbType.Int).Value = sQtyUsed
        MyCommon.LXSsp.Parameters.Add("@Value", SqlDbType.Decimal, 12).Value = sUnitValue
        MyCommon.LXSsp.Parameters.Add("@EarnedDate", SqlDbType.DateTime).Value = sDateTime
        MyCommon.LXSsp.Parameters.Add("@ExpireDate", SqlDbType.DateTime).Value = sExpiration
        MyCommon.LXSsp.Parameters.Add("@TotalValueEarned", SqlDbType.Decimal, 12).Value = sTotalAmount
        MyCommon.LXSsp.Parameters.Add("@RedeemedValue", SqlDbType.Decimal, 12).Value = sRedeemedAmount
        MyCommon.LXSsp.Parameters.Add("@BreakageValue", SqlDbType.Decimal, 12).Value = sBreakageAmount
        MyCommon.LXSsp.Parameters.Add("@LogixTransNum", SqlDbType.Char, 36).Value = sLogixTransNum
        MyCommon.LXSsp.Parameters.Add("@FirstName", SqlDbType.NVarChar, 50).Value = sDefaultFirstName
        MyCommon.LXSsp.Parameters.Add("@LastName", SqlDbType.NVarChar, 50).Value = sDefaultLastName
        MyCommon.LXSsp.Parameters.Add("@UpdateCount", SqlDbType.Int).Direction = ParameterDirection.Output
        MyCommon.LXSsp.ExecuteNonQuery()
        iUpdateCount = MyCommon.LXSsp.Parameters("@UpdateCount").Value
        MyCommon.Close_LXSsp()
        If iUpdateCount = 0 Then
          Throw New ApplicationException("Record for Stored Value (ExtCustomerID: " & sCardId & _
                                         ", SVProgramID: " & sProgramId & ", ExpireDate: " & sExpiration & _
                                         ") not found, (" & sStatus & ") update failed!")
        End If
      Else
        MyCommon.QueryStr = "dbo.pa_LogixServ_UpdateStoredValues"
        MyCommon.Open_LXSsp()
        MyCommon.LXSsp.Parameters.Add("@StatusFlag", SqlDbType.Int).Value = iStatus
        MyCommon.LXSsp.Parameters.Add("@ExtCustomerID", SqlDbType.NVarChar, 26).Value = sCardId
        MyCommon.LXSsp.Parameters.Add("@CardTypeID", SqlDbType.Int).Value = iTrxCardTypeId
        MyCommon.LXSsp.Parameters.Add("@LocationId", SqlDbType.BigInt).Value = sLocationId
        MyCommon.LXSsp.Parameters.Add("@LocalID", SqlDbType.BigInt).Value = sLocalID
        MyCommon.LXSsp.Parameters.Add("@ServerSerial", SqlDbType.Int).Value = sServerSerial
        MyCommon.LXSsp.Parameters.Add("@ExternalID", SqlDbType.NVarChar, 30).Value = sExtId
        MyCommon.LXSsp.Parameters.Add("@SVProgramID", SqlDbType.BigInt).Value = sProgramId
        MyCommon.LXSsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = sOfferId
        MyCommon.LXSsp.Parameters.Add("@QtyEarned", SqlDbType.Int).Value = sQtyEarned
        MyCommon.LXSsp.Parameters.Add("@QtyUsed", SqlDbType.Int).Value = sQtyUsed
        MyCommon.LXSsp.Parameters.Add("@Value", SqlDbType.Decimal, 12).Value = sUnitValue
        MyCommon.LXSsp.Parameters.Add("@EarnedDate", SqlDbType.DateTime).Value = sDateTime
        MyCommon.LXSsp.Parameters.Add("@ExpireDate", SqlDbType.DateTime).Value = sExpiration
        MyCommon.LXSsp.Parameters.Add("@TotalValueEarned", SqlDbType.Decimal, 12).Value = sTotalAmount
        MyCommon.LXSsp.Parameters.Add("@RedeemedValue", SqlDbType.Decimal, 12).Value = sRedeemedAmount
        MyCommon.LXSsp.Parameters.Add("@BreakageValue", SqlDbType.Decimal, 12).Value = sBreakageAmount
        MyCommon.LXSsp.Parameters.Add("@LogixTransNum", SqlDbType.Char, 36).Value = sLogixTransNum
        MyCommon.LXSsp.Parameters.Add("@FirstName", SqlDbType.NVarChar, 50).Value = sDefaultFirstName
        MyCommon.LXSsp.Parameters.Add("@LastName", SqlDbType.NVarChar, 50).Value = sDefaultLastName
        MyCommon.LXSsp.Parameters.Add("@UpdateCount", SqlDbType.Int).Direction = ParameterDirection.Output
        MyCommon.LXSsp.ExecuteNonQuery()
        iUpdateCount = MyCommon.LXSsp.Parameters("@UpdateCount").Value
        MyCommon.Close_LXSsp()

        If iUpdateCount = 0 Then
          If iStatus = 1 Then
            Throw New ApplicationException("Duplicate issuance for Stored Value (LocalID: " & sLocalID & _
                                           ", ServerSerial: " & sServerSerial & ")!")
          Else
            Throw New ApplicationException("Record for Stored Value (LocalID: " & sLocalID & _
                                           ", ServerSerial: " & sServerSerial & ") not found, (" & _
                                           sStatus & ") update failed!")
          End If
        End If
      End If
    Catch exApp As ApplicationException
      Throw
    Catch ex As Exception
      Throw
    End Try
  End Sub

  Private Sub UpdateFuelPartnerSV(ByVal sStatus As String, ByVal sExtCustomerID As String, ByVal sExtLocationCode As String, _
                               ByVal sLocalID As String, ByVal sServerSerial As String, ByVal sExtId As String, _
                               ByVal sUnitValue As String, ByVal sExpiration As String, ByVal sQtyEarned As String, _
                               ByVal sQtyUsed As String, ByVal sTotalAmount As String, ByVal sBreakageAmount As String, _
                               ByVal sRedeemedAmount As String, ByVal sDateTime As String, ByVal sProgramId As String, _
                               ByVal sOfferId As String, ByVal sLocationId As String, ByVal sLocalServerId As String, _
                               ByVal sLogixTransNum As String)
    Dim iStatus As Integer = 0
    Dim iUpdateCount As Integer = 0

    sCurrentMethod = "UpdateFuelPartnerSV"
    Try

      If sStatus = "Issued" Or sStatus = "Redeemed" Then
        ' is this the first time in this ACS transaction for Fuel Partner SV
        If MyCommon.LEXadoConn.State <> ConnectionState.Open Then
          ' open and begin a SQL transaction
          ' connection and SQL transaction are closed in method "ProcessTransactionXml"
          MyCommon.Open_LogixEX()
          MyCommon.QueryStr = "begin transaction"
          MyCommon.LEX_Execute()
          bBeginTransactionEX = True
        End If

        If sStatus = "Issued" Then
          iStatus = 1
          sServerSerial = sLocalServerId
        Else
          ' Redeemed
          iStatus = 4
          ' This code compensates for a problem with POS translating a ServerSerial value of "-9" to "0"
          ' The code simply puts the "-9" back in place, ASSUMING that a redemtion should never
          ' have SeverSerial = "0" and "0" only occurs now as a result of the bug at POS.
          ' The "-9" occurs when a SV is issued via central server.
          If sServerSerial = "0" Then
            sServerSerial = "-9"
          End If
        End If

        MyCommon.QueryStr = "dbo.pa_LogixServ_InsertCmFuelPartnerSV"
        MyCommon.Open_LEXsp()
        MyCommon.LEXsp.Parameters.Add("@StatusFlag", SqlDbType.Int).Value = iStatus
        MyCommon.LEXsp.Parameters.Add("@ExtCustomerID", SqlDbType.NVarChar, 26).Value = sExtCustomerID
        MyCommon.LEXsp.Parameters.Add("@LocationId", SqlDbType.BigInt).Value = sLocationId
        MyCommon.LEXsp.Parameters.Add("@LocalID", SqlDbType.BigInt).Value = sLocalID
        MyCommon.LEXsp.Parameters.Add("@ServerSerial", SqlDbType.Int).Value = sServerSerial
        MyCommon.LEXsp.Parameters.Add("@ExternalID", SqlDbType.NVarChar, 30).Value = sExtId
        MyCommon.LEXsp.Parameters.Add("@SVProgramID", SqlDbType.BigInt).Value = sProgramId
        MyCommon.LEXsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = sOfferId
        MyCommon.LEXsp.Parameters.Add("@QtyEarned", SqlDbType.Int).Value = sQtyEarned
        MyCommon.LEXsp.Parameters.Add("@QtyUsed", SqlDbType.Int).Value = sQtyUsed
        MyCommon.LEXsp.Parameters.Add("@Value", SqlDbType.Decimal, 12).Value = sUnitValue
        MyCommon.LEXsp.Parameters.Add("@EarnedDate", SqlDbType.DateTime).Value = sDateTime
        MyCommon.LEXsp.Parameters.Add("@ExpireDate", SqlDbType.DateTime).Value = sExpiration
        MyCommon.LEXsp.Parameters.Add("@TotalValueEarned", SqlDbType.Decimal, 12).Value = sTotalAmount
        MyCommon.LEXsp.Parameters.Add("@RedeemedValue", SqlDbType.Decimal, 12).Value = sRedeemedAmount
        MyCommon.LEXsp.Parameters.Add("@BreakageValue", SqlDbType.Decimal, 12).Value = sBreakageAmount
        MyCommon.LEXsp.Parameters.Add("@LogixTransNum", SqlDbType.Char, 36).Value = sLogixTransNum
        MyCommon.LEXsp.Parameters.Add("@UpdateCount", SqlDbType.Int).Direction = ParameterDirection.Output
        MyCommon.LEXsp.ExecuteNonQuery()
        iUpdateCount = MyCommon.LEXsp.Parameters("@UpdateCount").Value
        MyCommon.Close_LEXsp()

        If iUpdateCount < 1 Then
          Throw New ApplicationException("Insert for Fuel Partner Program failed (LocalID: " & sLocalID & _
                                         ", ServerSerial: " & sServerSerial & ")!")
        End If
      End If
    Catch exApp As ApplicationException
      Throw
    Catch ex As Exception
      Throw
    End Try
  End Sub

  Private Sub UpdateCustGroup(ByVal sOp As String, ByVal sId As String, ByVal iTrxCardTypeId As Integer, ByVal sGroup As String, ByVal sDateTime As String, ByVal sStoreNum As String, ByVal sTerminalNum As String, ByVal sTransNum As String, ByVal sLogixTransNum As String)
    Dim iStatus As Integer = 0

    sCurrentMethod = "UpdateCustGroup"
    Try
      If sOp = "ADD" Then
        sDefaultFirstName = MyCommon.Fetch_CM_SystemOption(20)
        sDefaultLastName = MyCommon.Fetch_CM_SystemOption(21)

        MyCommon.QueryStr = "dbo.pa_LogixServ_GroupMembership_Insert"
        MyCommon.Open_LXSsp()
        MyCommon.LXSsp.Parameters.Add("@ExtCustomerID", SqlDbType.NVarChar, 26).Value = sId
        MyCommon.LXSsp.Parameters.Add("@CardTypeID", SqlDbType.Int).Value = iTrxCardTypeId
        MyCommon.LXSsp.Parameters.Add("@CustomerGroupID", SqlDbType.BigInt, 8).Value = sGroup
        MyCommon.LXSsp.Parameters.Add("@LogixTransNum", SqlDbType.Char, 36).Value = sLogixTransNum
        MyCommon.LXSsp.Parameters.Add("@FirstName", SqlDbType.NVarChar, 50).Value = sDefaultFirstName
        MyCommon.LXSsp.Parameters.Add("@LastName", SqlDbType.NVarChar, 50).Value = sDefaultLastName
        MyCommon.LXSsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
        MyCommon.LXSsp.ExecuteNonQuery()
        iStatus = MyCommon.LXSsp.Parameters("@Status").Value
        MyCommon.Close_LXSsp()
        If iStatus <> 0 Then
          WriteLog("Customer is already a member of Group: '" & sGroup & "'", MessageType.Info)
        End If
      ElseIf sOp = "REMOVE" Then
        MyCommon.QueryStr = "dbo.pa_LogixServ_GroupMembership_Delete"
        MyCommon.Open_LXSsp()
        MyCommon.LXSsp.Parameters.Add("@ExtCustomerID", SqlDbType.NVarChar, 26).Value = sId
        MyCommon.LXSsp.Parameters.Add("@CardTypeID", SqlDbType.Int).Value = iTrxCardTypeId
        MyCommon.LXSsp.Parameters.Add("@CustomerGroupID", SqlDbType.BigInt, 8).Value = sGroup
        MyCommon.LXSsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
        MyCommon.LXSsp.ExecuteNonQuery()
        iStatus = MyCommon.LXSsp.Parameters("@Status").Value
        MyCommon.Close_LXSsp()
        If iStatus <> 0 Then
          WriteLog("Customer was not a member of Group: '" & sGroup & "'", MessageType.Info)
        End If
      Else
        Throw New ApplicationException("Invalid Operation Type: '" & sOp & "'")
      End If
    Catch exApp As ApplicationException
      Throw
    Catch ex As Exception
      Throw
    End Try
  End Sub

  Private Sub UpdatePromoMovt(ByVal sOfferId As String, ByVal sCount As String, ByVal sAmount As String, ByVal sDateTime As String, ByVal sStoreNum As String, ByVal sTerminalNum As String, ByVal sTransNum As String)
    Dim iStatus As Integer = 0
    Dim dt As DataTable
    Dim sLocationId As String


    sCurrentMethod = "UpdatePromoMovt"
    Try
      ' first determine if both location and offer exists in LogixRT db
      MyCommon.QueryStr = "dbo.pa_LogixServ_GetLocation_CheckOffer"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@StoreNum", SqlDbType.NVarChar, 20).Value = sStoreNum
      MyCommon.LRTsp.Parameters.Add("@OfferId", SqlDbType.BigInt, 8).Value = sOfferId
      dt = MyCommon.LRTsp_select()
      MyCommon.Close_LRTsp()

      If dt.Rows.Count = 2 Then
        ' yes, we have both; the first row, first column is LocationId
        sLocationId = dt.Rows(0).Item("LocationId")

        '***************************
        'AL-9237: Commenting these lines since stored procedure couldn't be found in base or client scripts.
        '****************************

        'MyCommon.QueryStr = "dbo.pt_OfferRedemption_Update"
        'MyCommon.Open_LWHsp()
        'MyCommon.LWHsp.Parameters.Add("@OfferId", SqlDbType.BigInt, 8).Value = sOfferId
        'MyCommon.LWHsp.Parameters.Add("@LocationId", SqlDbType.BigInt, 8).Value = sLocationId
        'MyCommon.LWHsp.Parameters.Add("@Count", SqlDbType.Int, 4).Value = sCount
        'MyCommon.LWHsp.Parameters.Add("@Amount", SqlDbType.Decimal, 9).Value = sAmount
        'MyCommon.LWHsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
        'MyCommon.LWHsp.ExecuteNonQuery()
        'iStatus = MyCommon.LWHsp.Parameters("@Status").Value
        'MyCommon.Close_LWHsp()
        'If iStatus <> 0 Then
        '  Throw New ApplicationException("Bad Status returned from stored procedure 'pt_OfferRedemption_Update'")
        'End If
      ElseIf dt.Rows.Count = 1 Then
        sLocationId = dt.Rows(0).Item("LocationId")
        If sLocationId = "0" Then
          Throw New ApplicationException("Location '" & sStoreNum & "' NOT found!")
        Else
          Throw New ApplicationException("Offer with Id '" & sOfferId & "' NOT found!")
        End If
      Else
        Throw New ApplicationException("Offer with Id '" & sOfferId & "' NOT found; Location '" & sStoreNum & "' NOT found!")
      End If
    Catch exApp As ApplicationException
      Throw
    Catch ex As Exception
      Throw
    End Try
  End Sub

  Private Sub UpdatePromoSummary(ByVal sOfferId As String, ByVal sCount As String, ByVal sAmount As String, ByVal sDateTime As String, ByVal sStoreNum As String)
    sCurrentMethod = "UpdatePromoSummary"
    Try
      MyCommon.QueryStr = "dbo.pa_RedemptionBuffer_Insert"
      MyCommon.Open_LWHsp()
      MyCommon.LWHsp.Parameters.Add("@OfferId", SqlDbType.BigInt, 8).Value = sOfferId
      MyCommon.LWHsp.Parameters.Add("@ReportingDate", SqlDbType.DateTime).Value = sDateTime
      MyCommon.LWHsp.Parameters.Add("@Redemptions", SqlDbType.Int, 4).Value = sCount
      MyCommon.LWHsp.Parameters.Add("@AmtRedeemed", SqlDbType.Decimal, 9).Value = sAmount
      MyCommon.LWHsp.ExecuteNonQuery()
      MyCommon.Close_LWHsp()
    Catch exApp As ApplicationException
      Throw
    Catch ex As Exception
      Throw
    End Try
  End Sub

  Private Function CustomerLockSet(ByVal lCustomerPk As Long, ByVal sStoreNum As String, ByVal sTerminal As String, ByVal sTrxNum As String, ByRef lUpdateCount As Long, ByVal OrgCustomerPK As Long) As Integer
    Dim iLockedStatus As Integer = 0
    Dim sLockedDate As String
    Dim sPrepay As String
    Dim dt As DataTable
    Dim dtLocked As Date
    Dim lMinutes As Long
    Dim sLocationId As String = "0"
    Dim CustLockDelMin As Integer
    Dim elapsed_time As TimeSpan
    Dim bCheckForLock As Boolean = True
    Dim iCheckForLockCount As Int16 = 0
    Dim bUpdateExistingLock As Boolean = False
    Dim sLocationIdLocked As String
    Dim sTerminalNumberLocked As String
    sCurrentMethod = "CustomerLockSet"
    Dim LockExpireMinutes As Long
    Dim NewLockExpireDate As DateTime
    Long.TryParse(MyCommon.Fetch_SystemOption(68), LockExpireMinutes)
    If LockExpireMinutes < 0 Then LockExpireMinutes = 0
    NewLockExpireDate = DateAdd(DateInterval.Minute, LockExpireMinutes, DateTime.Now)

    MyCommon.QueryStr = "select LocationID from Locations with (NoLock) where EngineID=0 and Deleted=0 and ExtLocationCode='" & sStoreNum & "';"
    dt = MyCommon.LRT_Select
    If dt.Rows.Count > 0 Then
      sLocationId = dt.Rows(0).Item(0)
    Else
      Throw New ApplicationException("Location with ExtLocationCode='" & sStoreNum & "' does not exist!")
    End If
    Do
      Try
        MyCommon.QueryStr = "select Prepay, LockedDate, LocationId, TerminalNumber from CustomerLock with (NoLock) where CustomerPK=" & lCustomerPk & ";"
        dt = MyCommon.LXS_Select()
        If dt.Rows.Count > 0 Then
          sTerminalNumberLocked = dt.Rows(0).Item("TerminalNumber")
          sLocationIdLocked = dt.Rows(0).Item("LocationId")
          sPrepay = dt.Rows(0).Item("Prepay")
          If sTerminalNumberLocked.Trim = sTerminal.Trim And sLocationIdLocked.Trim = sLocationId.Trim And sPrepay.ToLower = "false" Then
            ' same store, terminal, and prepay is false 
            ' so must be double request from suspend/resume at POS
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
                                sLocationId & ",TerminalNumber=" & sTerminal & ",TransactionNumber='" & sTrxNum & "', LockedBy=" & OrgCustomerPK & ",UE_LockExpireDate='" & NewLockExpireDate & _
                                "' where CustomerPK=" & lCustomerPk & ";"
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
          MyCommon.QueryStr = "insert into CustomerLock with (RowLock) (CustomerPK,LocationID,TerminalNumber,TransactionNumber,Prepay,LockedDate, LockedBy, UE_LockExpireDate)" & _
                    " values (" & lCustomerPk & "," & sLocationId.Trim & "," & sTerminal.Trim & ",'" & sTrxNum.Trim & "',0,getdate(), " & OrgCustomerPK & ", '" & NewLockExpireDate & "');"
          MyCommon.LXS_Execute()
          bCheckForLock = False
        End If
      Catch exApp As ApplicationException
        Throw
      Catch exSql As SqlException
        If exSql.Number = 2627 Then
          ' duplicate - a row was inserted since the above select was attempted, so try again!
          bCheckForLock = True
          iCheckForLockCount += 1
          If iCheckForLockCount > 1 Then
            If iCheckForLockCount > 3 Then
              Throw New ApplicationException("Exceeded Customer Lock contentention limit (" & iCheckForLockCount & ") (CustomerPK = " & lCustomerPk & ")!")
            Else
              WriteLog("Customer Lock contentention (" & iCheckForLockCount & ") (CustomerPK = " & lCustomerPk & ")!", MessageType.Warning)
            End If
          End If
        Else
          Throw
        End If
      Catch ex As Exception
        Throw
      End Try
    Loop While bCheckForLock
    Return iLockedStatus
  End Function

  Private Sub CustomerLockRelease(ByVal sPrepay As String, ByVal sCustomerId As String, ByVal iTrxCardTypeId As Integer, ByVal sStoreNum As String, ByVal sTerminal As String)
    Dim dt As DataTable
    Dim lCustomerPk As Long = 0
    Dim lHHPk As Long = 0
    Dim sLocationId As String = "0"
    Dim bTestCustomer As Boolean = False
    Dim bCardFound As Boolean

    Try
      MyCommon.QueryStr = "select LocationID from Locations with (NoLock) where EngineID=0 and Deleted=0 and ExtLocationCode='" & sStoreNum & "';"
      dt = MyCommon.LRT_Select
      If dt.Rows.Count > 0 Then
        sLocationId = dt.Rows(0).Item(0)
      Else
        Throw New ApplicationException("Location with ExtLocationCode='" & sStoreNum & "' does not exist!")
      End If

      sCustomerId = MyCommon.Pad_ExtCardID(sCustomerId, iTrxCardTypeId)

      ' get the CustomerPK from the Card number
      bCardFound = GetCustomerPK(sCustomerId, iTrxCardTypeId, lCustomerPk, lHHPk, bTestCustomer)
      If bCardFound Then
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

  Private Function StoredValueLockSet(ByVal lStoredValueId As Long, ByVal sStoreNum As String, ByVal sTerminal As String, ByVal sTrxNum As String) As Integer
    Dim iLockedStatus As Integer = 0
    Dim sLockedDate As String
    Dim dt As DataTable
    Dim dtLocked As Date
    Dim lMinutes As Long
    Dim sLocationId As String
    Dim CustLockDelMin As Integer
    Dim elapsed_time As TimeSpan

    Try
      If lStoredValueId > 0 Then
        MyCommon.QueryStr = "select LockedDate from StoredValueLocks with (NoLock) where StoredValueId=" & lStoredValueId & ";"
        dt = MyCommon.LXS_Select()
        If dt.Rows.Count > 0 Then
          sLockedDate = dt.Rows(0).Item("LockedDate")
          dtLocked = Date.Parse(sLockedDate)
          elapsed_time = DateTime.Now.Subtract(dtLocked)
          lMinutes = elapsed_time.TotalMinutes
          CustLockDelMin = MyCommon.Extract_Val(MyCommon.Fetch_SystemOption(68))
          If (CustLockDelMin = 0) OrElse (lMinutes < CustLockDelMin) Then
            iLockedStatus = 1
          Else
            MyCommon.QueryStr = "select LocationID from Locations with (NoLock) where EngineID=0 and Deleted=0 and ExtLocationCode='" & sStoreNum & "';"
            dt = MyCommon.LRT_Select
            If dt.Rows.Count > 0 Then
              sLocationId = dt.Rows(0).Item(0)
            Else
              Throw New ApplicationException("Location with ExtLocationCode='" & sStoreNum & "' does not exist!")
            End If

            MyCommon.QueryStr = "update StoredValueLocks with (RowLock) set LockedDate=getdate(), LocationId=" & _
                                sLocationId & ",TerminalNumber=" & sTerminal & ",TransactionNumber=" & sTrxNum & _
                                " where StoredValueId=" & lStoredValueId & ";"
            MyCommon.LXS_Execute()
          End If
        Else
          MyCommon.QueryStr = "select LocationID from Locations with (NoLock) where EngineID=0 and Deleted=0 and ExtLocationCode='" & sStoreNum & "';"
          dt = MyCommon.LRT_Select
          If dt.Rows.Count > 0 Then
            sLocationId = dt.Rows(0).Item(0)
          Else
            Throw New ApplicationException("Location with ExtLocationCode='" & sStoreNum & "' does not exist!")
          End If

          MyCommon.QueryStr = "insert into StoredValueLocks with (RowLock) (StoredValueId,LocationID,TerminalNumber,TransactionNumber,LockedDate)" & _
                    " values (" & lStoredValueId & "," & sLocationId & "," & sTerminal & "," & sTrxNum & ",getdate());"
          MyCommon.LXS_Execute()
        End If
      End If
      Return iLockedStatus
    Catch exApp As ApplicationException
      Throw
    Catch ex As Exception
      Throw
    End Try

  End Function

  Private Sub StoredValueLockRelease(ByVal sExtId As String)
    Try
      MyCommon.QueryStr = "delete from StoredValueLocks with (RowLock) where StoredValueId in" & _
                          " (select StoredValueId from StoredValue with (NoLock) where ExternalId='" & sExtId & _
                          "' and CustomerPK=0);"
      MyCommon.LXS_Execute()
    Catch exApp As ApplicationException
      Throw
    Catch ex As Exception
      Throw
    End Try

  End Sub

  Private Function UpdateLocationHealth(ByVal sStoreId As String, ByVal iClientCode As Integer, ByRef bTestLocation As Boolean) As String
    Dim i64LocationId As Int64 = 0
    Dim i64LocalServerId As Int64 = 0
    Dim bDistributeFiles As Boolean
    Dim sResponse As String = sOkStatus
    Dim objTemp As Object
    Dim iMaxStores As Integer
    Dim iMinFiles As Integer
    Dim iStoresDownLoadCount As Integer
    Dim iMinutesBeforeExpireDownloadCount As Integer

    sCurrentMethod = "UpdateLocationHealth"
    objTemp = MyCommon.Fetch_CM_SystemOption(29)
    If Not (Integer.TryParse(objTemp.ToString, iMaxStores)) Then iMaxStores = 0
    If iMaxStores < 0 Then
      iMaxStores = 0
    End If
    objTemp = MyCommon.Fetch_CM_SystemOption(30)
    If Not (Integer.TryParse(objTemp.ToString, iMinFiles)) Then iMinFiles = 1
    If iMinFiles < 1 Then
      iMinFiles = 1
    End If
    objTemp = MyCommon.Fetch_CM_SystemOption(34)
    If Not (Integer.TryParse(objTemp.ToString, iMinutesBeforeExpireDownloadCount)) Then iMinutesBeforeExpireDownloadCount = 0
    If iMinutesBeforeExpireDownloadCount <> 0 Then
      If iMinutesBeforeExpireDownloadCount < 5 Then
        iMinutesBeforeExpireDownloadCount = 5
      End If
    End If

    Try
      MyCommon.QueryStr = "dbo.pa_CmConnector_HealthUpdate"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@ExtLocationCode", SqlDbType.NVarChar, 20).Value = sStoreId
      MyCommon.LRTsp.Parameters.Add("@MaxStoresDownLoadCount", SqlDbType.Int).Value = iMaxStores
      MyCommon.LRTsp.Parameters.Add("@MinFilesDownLoadCount", SqlDbType.Int).Value = iMinFiles
      MyCommon.LRTsp.Parameters.Add("@MinutesBeforeExpireDownLoadCount", SqlDbType.Int).Value = iMinutesBeforeExpireDownloadCount
      MyCommon.LRTsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Direction = ParameterDirection.Output
      MyCommon.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.BigInt).Direction = ParameterDirection.Output
      MyCommon.LRTsp.Parameters.Add("@DistributeFiles", SqlDbType.Bit).Direction = ParameterDirection.Output
      MyCommon.LRTsp.Parameters.Add("@TestLocation", SqlDbType.Bit).Direction = ParameterDirection.Output
      MyCommon.LRTsp.Parameters.Add("@StoresDownLoadCount", SqlDbType.Int).Direction = ParameterDirection.Output
      MyCommon.LRTsp.ExecuteNonQuery()
      i64LocationId = MyCommon.LRTsp.Parameters("@LocationID").Value
      i64LocalServerId = MyCommon.LRTsp.Parameters("@LocalServerID").Value
      bDistributeFiles = MyCommon.LRTsp.Parameters("@DistributeFiles").Value
      bTestLocation = MyCommon.LRTsp.Parameters("@TestLocation").Value
      iStoresDownLoadCount = MyCommon.LRTsp.Parameters("@StoresDownLoadCount").Value
      MyCommon.Close_LRTsp()
      If i64LocationId > 0 Then
        If i64LocalServerId > 0 Then
          If MyCommon.Fetch_CM_SystemOption(iOptionOfferDistributionEnabled) = "1" Then
            If (bDistributeFiles) Then
              sResponse = sOkStatusDistRequestTrue
              WriteDebug("Files are ready for distribution.", DebugState.CurrentTime)
            Else
              If iStoresDownLoadCount >= iMaxStores Then
                If iStoresDownLoadCount = 1 Then
                  If iMinFiles = 1 Then
                    WriteDebug("Files are ready, but " & iStoresDownLoadCount & " other store (max " & iMaxStores & ") is waiting for a download of at least " & iMinFiles & " file.", DebugState.CurrentTime)
                  Else
                    WriteDebug("Files are ready, but " & iStoresDownLoadCount & " other store (max " & iMaxStores & ") is waiting for a download of at least " & iMinFiles & " files.", DebugState.CurrentTime)
                  End If
                Else
                  If iMinFiles = 1 Then
                    WriteDebug("Files are ready, but " & iStoresDownLoadCount & " other stores (max " & iMaxStores & ") are waiting for a download of at least " & iMinFiles & " file each.", DebugState.CurrentTime)
                  Else
                    WriteDebug("Files are ready, but " & iStoresDownLoadCount & " other stores (max " & iMaxStores & ") are waiting for a download of at least " & iMinFiles & " files each.", DebugState.CurrentTime)
                  End If
                End If
              End If
            End If
          End If
        Else
          Throw New ApplicationException("Error generating new LocalServer for ExtLocationCode:" & sStoreId)
        End If
      Else
        Throw New ApplicationException("Location with ExtLocationCode='" & sStoreId & "' does not exist!")
      End If
    Catch exApp As ApplicationException
      sResponse = BuildErrorXml(exApp.Message, eDefaultErrorType, False)
    Catch ex As Exception
      sResponse = BuildErrorXml(ex.Message, eDefaultErrorType, True)
    End Try
    Return sResponse
  End Function

  Private Sub UpdateCustomerCount(ByVal sId As String, ByVal iCardTypeId As Integer)
    Dim iStatus As Integer = 0
    sCurrentMethod = "UpdateCustomerCount"
    Try
      MyCommon.QueryStr = "dbo.pa_LogixServ_CustCountUpdate"
      MyCommon.Open_LXSsp()
      MyCommon.LXSsp.Parameters.Add("@CardID", SqlDbType.NVarChar, 256).Value = sId
      MyCommon.LXSsp.Parameters.Add("@CardTypeID", SqlDbType.Int).Value = iCardTypeId
      MyCommon.LXSsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
      MyCommon.LXSsp.ExecuteNonQuery()
      iStatus = MyCommon.LXSsp.Parameters("@Status").Value
      MyCommon.Close_LXSsp()
      If iStatus <> 0 Then
        Throw New ApplicationException("Card NOT found for Card ID '" & sId & "' of type '" & iCardTypeId & "' in procedure 'pa_LogixServ_CustCountUpdate'")
      End If
    Catch exApp As ApplicationException
      WriteLog(exApp.Message, MessageType.Warning)
    Catch ex As Exception
      Throw
    End Try
  End Sub

  Private Sub CheckForDuplicateMd5(ByVal sStoreId As String)

    If bDuplicateTransactionXml = DuplicateStatus.HasNotBeenChecked Then
      Dim dt As DataTable

      sCurrentMethod = sCurrentMethod & " - CheckForDuplicateMd5"
      bDuplicateTransactionXml = DuplicateStatus.IsNotDuplicate

      MyCommon.QueryStr = "select LocationId, CMConnMD5 from Locations with (NoLock) where EngineID=0 and Deleted=0 and ExtLocationCode='" & sStoreId & "';"
      dt = MyCommon.LRT_Select
      If (dt.Rows.Count > 0 AndAlso Not (dt.Rows(0).Item("LocationId").Equals(System.DBNull.Value))) Then
        i64CurrentLocationId = dt.Rows(0).Item("LocationId")
        If Not (dt.Rows(0).Item("CMConnMD5").Equals(System.DBNull.Value)) Then
          If sXmlMd5Hash = dt.Rows(0).Item("CMConnMD5") Then
            bDuplicateTransactionXml = DuplicateStatus.IsDuplicate
            Throw New ApplicationException("Duplicate Transation (ExtLocationCode=" & sStoreId & ") (hash=" & sXmlMd5Hash & ")")
          End If
        End If
      Else
        Throw New ApplicationException("Location with ExtLocationCode='" & sStoreId & "' does not exist!")
      End If
    End If

  End Sub

  Private Sub UpdateLocationMd5()

    sCurrentMethod = "UpdateLocationMd5"
    If i64CurrentLocationId > 0 Then
      MyCommon.QueryStr = "Update Locations with (RowLock) set CMConnMD5='" & sXmlMd5Hash & "' where LocationID=" & i64CurrentLocationId
      MyCommon.LRT_Execute()
    End If

  End Sub

  Private Function SetIpl(ByVal sStoreId As String) As String
    Dim sResponse As String = sOkStatus

    Try
      sCurrentStoreId = sStoreId
      sCurrentMethod = "SetIpl"
      sInputForLog = "*(Type=Input) (Method=SetIpl)] - (StoreId='" & sStoreId & "')"
      MyCommon = New Copient.CommonInc
      MyCommon.AppName = sAppName
      sStoreId = sStoreId.Trim(" ")

      eDefaultErrorType = ErrorType.SqlServer
      If MyCommon.LRTadoConn.State <> ConnectionState.Open Then
        MyCommon.Open_LogixRT()
        MyCommon.Load_System_Info()
        sInstallationName = MyCommon.InstallationName
      End If
      eDefaultErrorType = ErrorType.General

      MyCommon.QueryStr = "update Locations with (RowLock) set GenIpl=1 where Deleted=0 and ExtLocationCode='" & sStoreId & "'"
      MyCommon.LRT_Execute()
      If MyCommon.RowsAffected > 0 Then
        WriteLog("Location '" & sStoreId & "' has requested an IPL!", MessageType.Info)
      Else
        sResponse = BuildErrorXml("Location with ExtLocationCode='" & sStoreId & "' does not exist!", eDefaultErrorType, False)
      End If
      sInputForLog = ""

    Catch ex As Exception
      sResponse = BuildErrorXml(ex.Message, eDefaultErrorType)
    Finally
      If MyCommon.LWHadoConn.State <> ConnectionState.Closed Then
        MyCommon.Close_LogixWH()
      End If
      If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then
        MyCommon.Close_LogixXS()
      End If
      If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then
        MyCommon.Close_LogixRT()
      End If
    End Try
    Return sResponse
  End Function

  Private Function GetFileList(ByVal sStoreId As String) As String
    Dim sResponse As String = sOkStatus
    Dim dst As DataTable
    Dim row As DataRow
    Dim drs() As DataRow
    Dim sFileName As String
    Dim Settings As XmlWriterSettings
    Dim Writer As XmlWriter = Nothing
    Dim sw As StringWriter = Nothing
    Dim bTestLocation As Boolean = False
    Dim i64LocationId As Int64 = 0
    Dim objTemp As Object
    Dim iMaxRows As Integer = 0
    Dim iRunAgain As Integer = 0
    Dim iNumRowsUpdated As Integer = 0
    Dim iTotalRowsUpdated As Integer = 0

    Try
      sCurrentStoreId = sStoreId
      sCurrentMethod = "GetFileList"
      sInputForLog = "*(Type=Input) (Method=GetFileList)] - (StoreId='" & sStoreId & "')"
      MyCommon = New Copient.CommonInc
      MyCommon.AppName = sAppName
      sStoreId = sStoreId.Trim(" ")

      eDefaultErrorType = ErrorType.SqlServer
      If MyCommon.LRTadoConn.State <> ConnectionState.Open Then
        MyCommon.Open_LogixRT()
        MyCommon.Load_System_Info()
        sInstallationName = MyCommon.InstallationName
      End If
      eDefaultErrorType = ErrorType.General

      objTemp = MyCommon.Fetch_CM_SystemOption(27)
      If Not (Integer.TryParse(objTemp.ToString, iMaxRows)) Then iMaxRows = 0
      If iMaxRows < 0 Then
        iMaxRows = 0
      End If

      If iMaxRows > 0 Then
        WriteDebug(sCurrentMethod & " (Maximun number of files: " & iMaxRows & ")", DebugState.BeginTime)
      Else
        WriteDebug(sCurrentMethod, DebugState.BeginTime)
      End If

      sw = New StringWriter()
      Settings = New XmlWriterSettings()
      Settings.Indent = True
      Writer = XmlWriter.Create(sw, Settings)
      Writer.WriteStartDocument()
      Writer.WriteStartElement("OfferFiles")

      MyCommon.QueryStr = "select LocationId, TestingLocation from Locations with (NoLock) where Deleted=0 and ExtLocationCode='" & sStoreId & "';"
      dst = MyCommon.LRT_Select
      If dst.Rows.Count > 0 Then
        i64LocationId = MyCommon.NZ(dst.Rows(0).Item("LocationId"), 0)
        ' Test Cards enabled?
        If MyCommon.Fetch_SystemOption(88) = "1" Then
          bTestLocation = MyCommon.NZ(dst.Rows(0).Item("TestingLocation"), False)
          If bTestLocation Then
            Writer.WriteAttributeString("TestLocation", "yes")
          End If
        End If
        dst = Nothing

        ' get list of files with StatuFlag = 0 - Ready
        MyCommon.QueryStr = "dbo.pa_CmConnector_FileListGet"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = i64LocationId
        MyCommon.LRTsp.Parameters.Add("@MaxRowCount", SqlDbType.Int).Value = iMaxRows
        dst = MyCommon.LRTsp_select
        If dst.Rows.Count = 1 Then
          WriteDebug(dst.Rows.Count & " file is ready to be downloaded.", DebugState.CurrentTime)
        Else
          WriteDebug(dst.Rows.Count & " files are ready to be downloaded.", DebugState.CurrentTime)
        End If
        MyCommon.Close_LRTsp()
        If dst.Rows.Count > 0 Then
          drs = dst.Select("FileTypeId in (40,41)")
          If drs.Length > 0 Then
            Writer.WriteAttributeString("IPL", "yes")
          End If
          For Each row In dst.Rows
            sFileName = MyCommon.NZ(row.Item("Filename"), "")

            Writer.WriteStartElement("OfferFile")
            Writer.WriteAttributeString("Filename", sFileName)
            Writer.WriteAttributeString("StatusCode", "0")
            Writer.WriteEndElement() ' File
          Next

          ' set StatusFlag to 1 - Notified
          Do
            MyCommon.QueryStr = "dbo.pa_CmConnector_FileListGet_Update"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = i64LocationId
            MyCommon.LRTsp.Parameters.Add("@NumRowsUpdated", SqlDbType.Int).Direction = ParameterDirection.Output
            MyCommon.LRTsp.Parameters.Add("@RunAgain", SqlDbType.Int).Direction = ParameterDirection.Output
            MyCommon.LRTsp.ExecuteNonQuery()
            iNumRowsUpdated = MyCommon.LRTsp.Parameters("@NumRowsUpdated").Value
            iRunAgain = MyCommon.LRTsp.Parameters("@RunAgain").Value
            MyCommon.Close_LRTsp()
            If iNumRowsUpdated > 0 Then
              If iNumRowsUpdated = 1 Then
                WriteDebug("Store has been notified of " & iNumRowsUpdated & " file to download.", DebugState.CurrentTime)
              Else
                WriteDebug("Store has been notified of " & iNumRowsUpdated & " files to download.", DebugState.CurrentTime)
              End If
              iTotalRowsUpdated += iNumRowsUpdated
            End If
          Loop While iRunAgain > 0
          ' set number of files in Local Server for this location
          MyCommon.QueryStr = "dbo.pa_CmConnector_DownloadCountSet"
          MyCommon.Open_LRTsp()
          MyCommon.LRTsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = i64LocationId
          MyCommon.LRTsp.Parameters.Add("@Count", SqlDbType.BigInt).Value = iTotalRowsUpdated
          MyCommon.LRTsp.ExecuteNonQuery()
          MyCommon.Close_LRTsp()
        End If
      End If

      dst = Nothing
      row = Nothing

      Writer.WriteEndElement() ' OfferFiles
      Writer.WriteEndDocument()
      Writer.Flush()
      Writer.Close()

      sResponse = sw.ToString()
    Catch ex As Exception
      sResponse = BuildErrorXml(ex.Message, eDefaultErrorType)
    Finally
      WriteDebug(sCurrentMethod, DebugState.EndTime)
      If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then
        MyCommon.Close_LogixRT()
      End If
    End Try
    Return sResponse
  End Function

  Private Function UpdateFileStatus(ByVal sStoreId As String, ByVal sStatusXml As String) As String
    Dim sResponse As String = sOkStatus
    Dim sLocationId As String
    Dim sFilename As String
    Dim sStatus As String
    Dim sError As String
    Dim sMsg As String
    Dim iStatus As Integer

    Dim sr As StringReader = Nothing
    Dim Settings As XmlReaderSettings
    Dim xr As XmlReader = Nothing
    Dim xrSubTree As XmlReader = Nothing
    Dim dt As DataTable = Nothing
    Dim i As Integer
    Dim iMaxRows As Integer = 50
    Dim sStatusFileList(5) As String
    Dim iStatusRowCount(5) As Integer
    Dim iStatusTotalRowCount(5) As Integer
    Dim iTotalFileCount As Integer
    Dim objTemp As Object

    Try
      sCurrentStoreId = sStoreId
      sCurrentMethod = "UpdateFileStatus"
      sInputForLog = "*(Type=Input) (Method=UpdateFileStatus)] - (StoreId='" & sStoreId & "' - " & sStatusXml & ")"
      MyCommon = New Copient.CommonInc
      MyCommon.AppName = sAppName
      sStoreId = sStoreId.Trim(" ")

      WriteDebug(sCurrentMethod, DebugState.BeginTime)
      eDefaultErrorType = ErrorType.SqlServer
      If MyCommon.LRTadoConn.State <> ConnectionState.Open Then
        MyCommon.Open_LogixRT()
        MyCommon.Load_System_Info()
        sInstallationName = MyCommon.InstallationName
      End If
      eDefaultErrorType = ErrorType.General

      MyCommon.QueryStr = "select LocationId from Locations with (NoLock) where EngineID=0 and Deleted=0 and ExtLocationCode='" & sStoreId & "';"
      dt = MyCommon.LRT_Select
      If dt.Rows.Count > 0 Then
        sLocationId = MyCommon.NZ(dt.Rows(0).Item("LocationId"), "")

        objTemp = MyCommon.Fetch_CM_SystemOption(28)
        If Not (Integer.TryParse(objTemp.ToString, iMaxRows)) Then iMaxRows = 1
        If iMaxRows < 1 Then
          iMaxRows = 1
        ElseIf iMaxRows > 1800 Then
          iMaxRows = 1800
        End If

        sr = New StringReader(sStatusXml)
        Settings = New XmlReaderSettings()
        xr = XmlReader.Create(sr, Settings)
        xr.ReadToFollowing("OfferFiles")
        If xr.EOF Then
          Throw New ApplicationException("No 'OfferFiles' root element")
        End If

        ' will parse filelist into groups by status
        For i = 0 To 4
          sStatusFileList(i) = ""
          iStatusRowCount(i) = 0
          iStatusTotalRowCount(i) = 0
        Next

        Do
          xr.ReadToFollowing("OfferFile")
          If xr.EOF Then
            Exit Do
          End If
          Do
            If xr.AttributeCount = 0 Then
              Throw New ApplicationException("No attributes in 'OfferFile' element.")
            End If

            sFilename = xr.GetAttribute("Filename")
            If sFilename Is Nothing Then
              Throw New ApplicationException("No 'Filename' attribute in 'OfferFile' element.")
            End If

            sStatus = xr.GetAttribute("StatusCode")
            If sStatus Is Nothing Then
              Throw New ApplicationException("No 'StatusCode' attribute in 'OfferFile' element.")
            End If
            iStatus = Integer.Parse(sStatus)

            If (iMaxRows = 1) Or (iStatus = 3) Then
              ' always use block = 1 for error status, since each message can be unique
              sMsg = aStatusMsgs(iStatus)
              If iStatus = 3 Then
                ' If error then check for error description
                xrSubTree = xr.ReadSubtree()
                xrSubTree.ReadToFollowing("ErrorDescription")
                If Not xrSubTree.EOF Then
                  Do
                    If xrSubTree.AttributeCount = 0 Then
                      Throw New ApplicationException("No attributes in 'ErrorDescription' element.")
                    End If
                    sError = xrSubTree.GetAttribute("Text")
                    If Not sError Is Nothing Then
                      If sError.Length > 0 Then
                        sMsg += " - (" & sError & ")"
                      End If
                    End If
                  Loop While xrSubTree.ReadToNextSibling("ErrorDescription")
                End If
                xrSubTree.Close()
                xrSubTree = Nothing
                If sMsg.Length > 0 Then
                  sMsg = sMsg.Replace("'", "^")
                End If
              End If

              MyCommon.QueryStr = "dbo.pa_LogixServ_FileListUpdate"
              MyCommon.Open_LRTsp()
              MyCommon.LRTsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = sLocationId
              MyCommon.LRTsp.Parameters.Add("@FileName", SqlDbType.NVarChar, 50).Value = sFilename
              MyCommon.LRTsp.Parameters.Add("@StatusFlag", SqlDbType.Int).Value = sStatus
              MyCommon.LRTsp.Parameters.Add("@StatusMsg", SqlDbType.NVarChar, 1000).Value = sMsg
              MyCommon.LRTsp.ExecuteNonQuery()
              MyCommon.Close_LRTsp()

              iStatusTotalRowCount(iStatus) += 1
            Else
              ' accumulate filelist until block is full
              If iStatusRowCount(iStatus) > 0 Then
                sStatusFileList(iStatus) += ",'" & sFilename & "'"
              Else
                sStatusFileList(iStatus) = "'" & sFilename & "'"
              End If
              iStatusRowCount(iStatus) += 1
              If iStatusRowCount(iStatus) >= iMaxRows Then
                UpdateFileBatch(sLocationId, iStatus, aStatusMsgs(iStatus), sStatusFileList(iStatus))
                iStatusTotalRowCount(iStatus) += iStatusRowCount(iStatus)
                iStatusRowCount(iStatus) = 0
                sStatusFileList(iStatus) = ""
              End If
            End If
          Loop While xr.ReadToNextSibling("OfferFile")
        Loop While xr.ReadToNextSibling("OfferFiles")

        ' Update in partially full batches
        For i = 0 To 4
          If iStatusRowCount(i) > 0 Then
            UpdateFileBatch(sLocationId, i, aStatusMsgs(i), sStatusFileList(i))
            iStatusTotalRowCount(i) += iStatusRowCount(i)
          End If
        Next

        ' Log Updates by Status
        iTotalFileCount = 0
        For i = 0 To 4
          If iStatusTotalRowCount(i) > 0 Then
            ' Error updates already done, so just log the number
            If i = 3 Then
              WriteDebug("Updated File Status to '" & aStatusMsgs(i) & "' for " & iStatusTotalRowCount(i) & " file(s).", DebugState.CurrentTime)
            Else
              WriteDebug("Updated File Status to '" & aStatusMsgs(i) & "' for " & iStatusTotalRowCount(i) & " file(s). (block=" & iMaxRows & ")", DebugState.CurrentTime)
            End If
            iTotalFileCount += iStatusTotalRowCount(i)
          End If
        Next

        ' reset number of files in Local Server for this location
        MyCommon.QueryStr = "dbo.pa_CmConnector_DownloadCountSet"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = sLocationId
        MyCommon.LRTsp.Parameters.Add("@Count", SqlDbType.BigInt).Value = 0
        MyCommon.LRTsp.ExecuteNonQuery()
        MyCommon.Close_LRTsp()

      Else
        Throw New ApplicationException("Location with ExtLocationCode='" & sStoreId & "' does not exist!")
      End If
    Catch exApp As ApplicationException
      sResponse = BuildErrorXml(exApp.Message, eDefaultErrorType, False)
    Catch exXml As XmlException
      sResponse = BuildErrorXml(exXml.Message, eDefaultErrorType, False)
    Catch ex As Exception
      sResponse = BuildErrorXml(ex.Message, eDefaultErrorType)
    Finally
      If Not dt Is Nothing Then
        dt.Dispose()
        dt = Nothing
      End If
      If Not xrSubTree Is Nothing Then
        xrSubTree.Close()
      End If
      If Not xr Is Nothing Then
        xr.Close()
      End If
      If Not sr Is Nothing Then
        sr.Close()
        sr.Dispose()
      End If
      WriteDebug(sCurrentMethod, DebugState.EndTime)
      If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then
        MyCommon.Close_LogixRT()
      End If
    End Try
    Return sResponse

  End Function

  Private Function UpdateFileBatch(ByVal sLocationId As Long, ByVal istatus As Integer, ByVal sMsg As String, ByRef sFileList As String) As Boolean
    Dim bOk As Boolean = True

    MyCommon.QueryStr = "update OfferDLBuffer with (RowLock) set StatusFlag=" & istatus & _
                        " ,StatusDate=getdate(), StatusMessage='" & sMsg & "'" & _
                        " where LocationId=" & sLocationId & " and FileName in (" & sFileList & ");"
    MyCommon.LRT_Execute()

    Return bOk
  End Function

  Private Function GetStringMd5(ByVal sXml As String) As String
    Dim bites() As Byte
    Dim md5 As New MD5CryptoServiceProvider()
    Dim sHex As String = ""

    bites = md5.ComputeHash(ConvertStringToByteArray(sXml))
    sHex = ConvertByteArrayToHexString(bites)

    Return sHex

  End Function

  Private Function GetFileMd5(ByVal sFullPath As String) As String
    Dim bites() As Byte
    Dim md5 As New MD5CryptoServiceProvider()
    Dim sHex As String = ""
    Dim sr As System.IO.Stream

    sr = New System.IO.FileStream(sFullPath, IO.FileMode.Open, IO.FileAccess.Read)
    bites = md5.ComputeHash(sr)
    sr.Close()

    sHex = ConvertByteArrayToHexString(bites)

    Return sHex

  End Function

  Private Function ConvertStringToByteArray(ByVal sStringToConvert As String) As Byte()
    Dim encode As New UTF8Encoding
    Dim bites() As Byte

    bites = encode.GetBytes(sStringToConvert)

    Return bites

  End Function

  Private Function ConvertByteArrayToHexString(ByVal bites As Byte()) As String
    Dim b As Byte
    Dim sbHex As New System.Text.StringBuilder(bites.Length)

    For Each b In bites
      sbHex.Append(b.ToString("x2"))
    Next

    Return sbHex.ToString()
  End Function

  Private Function BuildErrorXml(ByVal sText As String, ByVal eType As ErrorType) As String
    Dim sXml As String

    sXml = BuildErrorXml(sText, eType, True)
    Return sXml
  End Function

  Private Function BuildErrorXml(ByVal sText As String, ByVal eErrType As ErrorType, ByVal bSystemError As Boolean) As String
    Dim sXml As String
    Dim iErrType As Integer
    Dim sLogText As String
    Dim eMsgType As MessageType

    iErrType = eErrType
    If bSystemError Then
      eMsgType = MessageType.SysError
      Try
        MyCommon.Error_Processor(, sCurrentMethod, sAppName, sInstallationName)
      Catch
      End Try
    Else
      eMsgType = MessageType.AppError
    End If
    sLogText = WriteLog(sText, eMsgType)

    sXml = "<Response Status =""Error"">" & vbCrLf & _
    "  <Error>" & vbCrLf & _
    "    <Version>" & sVersion & "</Version>" & vbCrLf & _
    "    <Type>" & iErrType.ToString & "</Type>" & vbCrLf & _
    "    <Method>" & sCurrentMethod & "</Method>" & vbCrLf & _
    "    <Text>" & sLogText & "</Text>" & vbCrLf & _
    "  </Error>" & vbCrLf & _
    "</Response>" & vbCrLf

    Return sXml
  End Function

  Private Sub WriteDebug(ByVal sText As String, ByVal mode As DebugState)
    If bDebugLogOn Then
      Dim TotalSeconds As Double
      Dim sIndent As String
      Select Case mode
        Case DebugState.BeginTime
          ' first call
          DebugStartTimes.Add(Now)
          If DebugStartTimes.Count = 1 Then
            Dim sIPAddress As String
            Dim sLocInfo As String = ""

            If i32CurrentLocalServerId = -1 Then
              If sCurrentStoreId.Length > 0 Then
                If MyCommon.LRTadoConn.State <> ConnectionState.Open Then
                  MyCommon.Open_LogixRT()
                  MyCommon.Load_System_Info()
                  sInstallationName = MyCommon.InstallationName
                End If

                Try
                  MyCommon.QueryStr = "dbo.pa_LogixServ_GetLocalServerId"
                  MyCommon.Open_LRTsp()
                  MyCommon.LRTsp.Parameters.Add("@ExtLocationCode", SqlDbType.NVarChar, 20).Value = sCurrentStoreId
                  MyCommon.LRTsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Direction = ParameterDirection.Output
                  MyCommon.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Direction = ParameterDirection.Output
                  MyCommon.LRTsp.ExecuteNonQuery()
                  i64CurrentLocationId = MyCommon.LRTsp.Parameters("@LocationID").Value
                  i32CurrentLocalServerId = MyCommon.LRTsp.Parameters("@LocalServerID").Value
                  MyCommon.Close_LRTsp()
                Catch
                  i32CurrentLocalServerId = 0
                  i64CurrentLocationId = 0
                End Try
              Else
                i32CurrentLocalServerId = 0
                i64CurrentLocationId = 0
              End If
            End If

            If sCurrentStoreId.Length > 0 Then
              sLocInfo = " Location: '" & sCurrentStoreId & "', ID: " & i64CurrentLocationId & ", LSID: " & i32CurrentLocalServerId & ", IP: "
            Else
              sLocInfo = " IP: "
            End If
            'sIPAddress = HttpContext.Current.Request.ServerVariables("HTTP_VIA")
            'If sIPAddress = "" Then
            '  sIPAddress = HttpContext.Current.Request.ServerVariables("REMOTE_ADDR")
            'Else
            '  sIPAddress = HttpContext.Current.Request.ServerVariables("HTTP_X_FORWARDED_FOR")
            'End If
            sIPAddress = GetComputerIP()
            sLocInfo += sIPAddress
            WriteLog(scDashes & sLocInfo, MessageType.Debug)
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

  Private Function GetComputerIP() As String
    Dim sHostName As String
    Dim sServerIP As String

    sHostName = System.Net.Dns.GetHostName()
    sServerIP = System.Net.Dns.GetHostEntry(sHostName).AddressList.GetValue(0).ToString

    Return sServerIP
  End Function

  Private Function WriteLog(ByVal sText As String, ByVal eType As MessageType) As String
    Dim sFileName As String
    Dim sLogText As String = ""
    Try
      If MyCommon.LRTadoConn.State <> ConnectionState.Open Then
        MyCommon.Open_LogixRT()
        MyCommon.Load_System_Info()
        sInstallationName = MyCommon.InstallationName
      End If
      If bUseStoreIdForLog Then
        If sCurrentStoreId.Length > 0 Then
          sFileName = sLogFileName & "-" & sCurrentStoreId.Trim & "." & MyCommon.Leading_Zero_Fill(Year(Today), 4) & MyCommon.Leading_Zero_Fill(Month(Today), 2) & MyCommon.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
        Else
          sFileName = sLogFileName & "." & MyCommon.Leading_Zero_Fill(Year(Today), 4) & MyCommon.Leading_Zero_Fill(Month(Today), 2) & MyCommon.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
        End If
      Else
        If i32CurrentLocalServerId = -1 Then
          If sCurrentStoreId.Length > 0 Then
            Try
              MyCommon.QueryStr = "dbo.pa_LogixServ_GetLocalServerId"
              MyCommon.Open_LRTsp()
              MyCommon.LRTsp.Parameters.Add("@ExtLocationCode", SqlDbType.NVarChar, 20).Value = sCurrentStoreId
              MyCommon.LRTsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Direction = ParameterDirection.Output
              MyCommon.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Direction = ParameterDirection.Output
              MyCommon.LRTsp.ExecuteNonQuery()
              i64CurrentLocationId = MyCommon.LRTsp.Parameters("@LocationID").Value
              i32CurrentLocalServerId = MyCommon.LRTsp.Parameters("@LocalServerID").Value
              MyCommon.Close_LRTsp()
            Catch
              i32CurrentLocalServerId = 0
            End Try
          Else
            i32CurrentLocalServerId = 0
          End If
        End If
        sFileName = sLogFileName & "-" & Format(i32CurrentLocalServerId, "00000") & "." & MyCommon.Leading_Zero_Fill(Year(Today), 4) & MyCommon.Leading_Zero_Fill(Month(Today), 2) & MyCommon.Leading_Zero_Fill(Microsoft.VisualBasic.DateAndTime.Day(Today), 2) & ".txt"
      End If

      If eType = MessageType.Debug Then
        sLogText = "[" & Format(Date.Now, "MM/dd/yyyy HH:mm:ss.fffzzz") & " (Type=" & eType.ToString & ")] " & sText
      Else
        sText = sText.Replace(ControlChars.CrLf, " ")
        If sInputForLog.Length > 0 Then
          sLogText = "[" & Format(Date.Now, "MM/dd/yyyy HH:mm:ss.fffzzz") & sInputForLog & ControlChars.CrLf
          sInputForLog = ""
        End If
        sLogText = sLogText & "[" & Format(Date.Now, "MM/dd/yyyy HH:mm:ss.fffzzz") & " (Type=" & eType.ToString & ") (Method=" & sCurrentMethod & ")] " & sText
      End If
      MyCommon.Write_Log(sFileName, sLogText)
    Catch ex As Exception
      Try
        MyCommon.Error_Processor(, "WriteLog", sAppName, sInstallationName)
      Catch
      End Try
      sText += " (WriteLog Error: " & ex.Message & ")"
    End Try
    Return sText
  End Function

  Private Sub WriteBalancesToLog(ByVal dtBalances As DataTable)
    Dim row As DataRow
    Dim rowText As String

    For Each row In dtBalances.Rows
      rowText = "ID: " & MyCommon.NZ(row.Item("ID"), "") & "  Desc: " & MyCommon.NZ(row.Item("Desc"), "") & _
                "  Val: " & MyCommon.NZ(row.Item("Val"), "")
      WriteLog(rowText, MessageType.Info)
    Next

  End Sub

  Private Function GetLocalDateTime(ByVal sDateTime As String) As String
    Dim sLocalDateTime As String
    Dim dtTemp As Date

    If sDateTime Is Nothing OrElse sDateTime.Length = 0 Then
      dtTemp = Date.Now
      sLocalDateTime = Format(dtTemp, sDateFormat)
    Else
      sLocalDateTime = sDateTime.Substring(0, 19)
    End If

    Return sLocalDateTime
  End Function

  Private Function GetMembershipLevel(ByVal lCustomerPK As Long) As String
    Dim sLevel As String = ","
    Dim dt1, dt2 As DataTable
    Dim sMembershipGroupList As String
    Dim sMemberList() As String
    Dim rows() As DataRow
    Dim i As Integer
    Dim iLevel As Integer

    sMembershipGroupList = MyCommon.Fetch_CM_SystemOption(38)
    If sMembershipGroupList <> "" Then
      sMemberList = sMembershipGroupList.Split(",")
      If sMemberList.Length > 0 Then
        MyCommon.QueryStr = "select distinct CustomerGroupID from GroupMembership with (NoLock)" & _
                            " where Deleted=0 and CustomerPK=" & lCustomerPK & _
                            " and CustomerGroupID in (" & sMembershipGroupList & ");"
        dt1 = MyCommon.LXS_Select
        If (dt1.Rows.Count > 0) Then
          For i = sMemberList.Length - 1 To 0 Step -1
            rows = dt1.Select("CustomerGroupID=" & sMemberList(i))
            If rows.Length > 0 Then
              MyCommon.QueryStr = "select Name from CustomerGroups with (NoLock) where CustomerGroupID=" & sMemberList(i) & ";"
              dt2 = MyCommon.LRT_Select
              If (dt2.Rows.Count > 0) Then
                iLevel = i + 1
                sLevel = iLevel.ToString & "," & MyCommon.NZ(dt2.Rows(0).Item("Name"), "")
                Exit For
              End If
            End If
          Next
        End If
      End If
    End If

    Return sLevel
  End Function

  Protected Function makeCSVCustomerGroupList(ByRef xmlCustomerGroupListForHousehold As XmlDocument) As String
    Dim csvCustomerGroupList As New StringBuilder ' = String.Empty
    Dim nodes As XmlNodeList = xmlCustomerGroupListForHousehold.SelectNodes("//GroupMembership/CustomerGroupID")

    For Each aNode As XmlNode In nodes
      csvCustomerGroupList.AppendFormat(",{0}", aNode.InnerText)
    Next

    Return csvCustomerGroupList.ToString
  End Function

  Private Function GetMembershipLevelForAutoHouseHold(ByRef xmlCustomerGroupListForHousehold As XmlDocument) As String
    ' find the intersection of the sets customer-group-list-for-household and membership-level-customer-group-ids 
    Dim sLevel As String = ","
    Dim sMembershipGroupList As String = MyCommon.Fetch_CM_SystemOption(38) ' 38, N'Membership Level customer group IDs'
    If sMembershipGroupList <> "" Then

      Dim sCustomerGroupListForHousehold As String = makeCSVCustomerGroupList(xmlCustomerGroupListForHousehold)
      Dim sMemberList() As String = sMembershipGroupList.Split(",")
      If sMemberList.Length > 0 AndAlso sCustomerGroupListForHousehold.Length > 0 Then

        ' add bookends for string search
        sCustomerGroupListForHousehold = "," & sCustomerGroupListForHousehold & ","
        For i As Integer = sMemberList.Length - 1 To 0 Step -1

          Dim delimitedMember As String = "," & sMemberList(i) & ","
          If sCustomerGroupListForHousehold.IndexOf(delimitedMember) > -1 Then ' sMemberList(i) is included in sCustomerGroupListForHousehold

            MyCommon.QueryStr = "select Name from CustomerGroups with (NoLock) where CustomerGroupID = " & sMemberList(i) & ";"
            Dim dt As DataTable = MyCommon.LRT_Select
            If (dt.Rows.Count > 0) Then
              ' ? This code does NOT append to sLevel, it resets it. Is this the desired behaviour?
              Dim iLevel As Integer = i + 1
              sLevel = iLevel.ToString & "," & MyCommon.NZ(dt.Rows(0).Item("Name"), "")
              Exit For
            End If

          End If
        Next

      End If

    End If

    Return sLevel
  End Function

  Public Function getCustomerGroups(ByVal customerpk As Long, ByVal hhpk As Long, ByVal AutoHouseholdCustGrpEnabled As Boolean) As System.Xml.XmlDocument
    Dim query As String = IIf(AutoHouseholdCustGrpEnabled, "dbo.pa_LogixServ_FetchCustGroups_MemberOrHousehold_asXML", "dbo.pa_LogixServ_FetchCustGroups_asXML")

    MyCommon.Open_LogixXS()
    Dim sqlcmd As New Data.SqlClient.SqlCommand(query, MyCommon.LXSadoConn)
    sqlcmd.CommandType = CommandType.StoredProcedure
    sqlcmd.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = customerpk
    sqlcmd.Parameters.Add("@HHPK", SqlDbType.BigInt).Value = hhpk
    Dim xr As XmlReader = sqlcmd.ExecuteXmlReader()

    Dim customer_groups As New XmlDocument
    customer_groups.Load(xr)

    If (customer_groups.DocumentElement Is Nothing) Then
      Const EMPTY_CUSTOMERGROUP_LIST_XML As String = "<LXSsp_select />"
      WriteLog(String.Format("CustomerPK {0} is a member of no groups.", customerpk), MessageType.Debug)
      customer_groups.LoadXml(EMPTY_CUSTOMERGROUP_LIST_XML)
    End If

    Return customer_groups

  End Function

  Public Function getNewCustomerPromotions_sqlcmd(ByVal startDate As String, ByVal endDate As String) As Data.SqlClient.SqlCommand

        MyCommon.QueryStr = "SELECT DISTINCT OCLV.OfferId, isnull(CPSub.PrefCondition, 0) as HasPrefCondition " & _
                            "FROM CM_ST_OfferCustLocView as OCLV WITH (NoLock) " & _
                            "Left Join (select distinct OC.OfferID, 1 as PrefCondition from CM_ST_ConditionPreferenceValues as CP with (NoLock) Inner Join CM_ST_OfferConditions as OC with(NoLock) on CP.ConditionID=OC.ConditionID) as CPSub " & _
                            "    on CPSub.OfferID=OCLV.OfferID " & _
                            "WHERE OCLV.ProdStartDate <= '" & startDate & "' AND OCLV.ProdEndDate >= '" & endDate & "' AND (OCLV.NewCardholders=1 OR (OCLV.AnyCardHolder=1 AND (OCLV.ExCustGroupId is Not null))) " & _
                            "ORDER BY OfferId"
    MyCommon.Open_LogixRT() ' open the connection to RT if it has not been

    Dim sqlcmd As New Data.SqlClient.SqlCommand(MyCommon.QueryStr, MyCommon.LRTadoConn)

    Return sqlcmd

  End Function

  Public Function getExistingCustomerPromotions_sqlcmd(ByRef customerGroups As System.Xml.XmlReader, ByVal startDate As String, ByVal endDate As String, ByVal useTest As Boolean) As Data.SqlClient.SqlCommand
    Dim query As String = IIf(useTest, "pa_cmConnector_TestMemberInfo", "pa_cmConnector_MemberInfo")

    MyCommon.Open_LogixRT() ' open the connection to RT if it has not been
    Dim sqlcmd As New Data.SqlClient.SqlCommand(query, MyCommon.LRTadoConn)
    MyCommon.QueryStr = query ' this is for the Error_Processor() routine, which adds the value of this to the error messages it prints (tight coupling)

    sqlcmd.CommandType = CommandType.StoredProcedure
    sqlcmd.Parameters.Add("@sBusDateStart", SqlDbType.DateTime).Value = startDate
    sqlcmd.Parameters.Add("@sBusDateEnd", SqlDbType.DateTime).Value = endDate
    sqlcmd.Parameters.Add("@LXSsp_select", SqlDbType.Xml).Value = customerGroups

    Return sqlcmd

  End Function

  Public Function getPromotions(ByRef customerGroups As System.Xml.XmlReader, ByVal newCustomer As Boolean, ByVal startDate As String, ByVal endDate As String, ByVal useTest As Boolean, ByVal CustomerPK As Long) As System.Data.DataTable
    Dim sqlcmd As Data.SqlClient.SqlCommand

    If (newCustomer) Then
      sqlcmd = getNewCustomerPromotions_sqlcmd(startDate, endDate)
    Else
      sqlcmd = getExistingCustomerPromotions_sqlcmd(customerGroups, startDate, endDate, useTest)
    End If

    Dim adapter As New SqlDataAdapter(sqlcmd)
    Dim TempDataSet As New DataSet
    adapter.Fill(TempDataSet)
    Dim dtPromos As DataTable = TempDataSet.Tables(0)
    dtPromos.TableName = "Promos"
    Evaluate_Preference_Offers(dtPromos, CustomerPK)
    Return dtPromos

  End Function

  Private Function getPromotionsFobOnly(ByVal startDate As String, ByVal endDate As String, ByVal CustomerPK As Long) As System.Data.DataTable
    Dim dtPromos As DataTable

    MyCommon.QueryStr = "pa_cmConnector_getFobOnly"
    MyCommon.Open_LRTsp()
    MyCommon.LRTsp.Parameters.Add("@sBusDateStart", SqlDbType.DateTime).Value = startDate
    MyCommon.LRTsp.Parameters.Add("@sBusDateEnd", SqlDbType.DateTime).Value = endDate
    dtPromos = MyCommon.LRTsp_select()
    MyCommon.Close_LRTsp()
    dtPromos.TableName = "Promos"

    Return dtPromos

  End Function

  Private Function getPromotionsFobTargeted(ByRef customerGroups As System.Xml.XmlReader, ByVal startDate As String, ByVal endDate As String, ByVal CustomerPK As Long) As System.Data.DataTable
    Dim dtPromos As DataTable

    MyCommon.QueryStr = "pa_cmConnector_getFobTargeted"
    MyCommon.Open_LRTsp()
    MyCommon.LRTsp.Parameters.Add("@sBusDateStart", SqlDbType.DateTime).Value = startDate
    MyCommon.LRTsp.Parameters.Add("@sBusDateEnd", SqlDbType.DateTime).Value = endDate
    MyCommon.LRTsp.Parameters.Add("@CustomerGroupsXml", SqlDbType.Xml).Value = customerGroups
    dtPromos = MyCommon.LRTsp_select()
    MyCommon.Close_LRTsp()
    dtPromos.TableName = "Promos"

    Return dtPromos

  End Function

  Private Sub Evaluate_Preference_Offers(ByRef dtPromos As DataTable, ByVal CustomerPK As Long)
    Dim row As DataRow
    Dim OfferID As Long = 0

    'if the customer is eligible for any targeted offers and EPM is installed
    If Not (dtPromos Is Nothing) AndAlso MyCommon.IsIntegrationInstalled(Integrations.PREFERENCE_MANAGER) Then
      WriteDebug("Begin Evaluate Preferences", DebugState.CurrentTime)
      'loop through the customers offers that have a preference condition 
      For Each row In dtPromos.Select("HasPrefCondition=1")
        OfferID = row.Item("OfferID")
        If Not (PrefLib.Evaluate_CMOffer_PrefConditions(OfferID, CustomerPK)) Then
          'the customer does not meet the preference condition for this offer
          'we need to remove the offer from the data table
          dtPromos.Rows.Remove(row)
        End If
      Next
      dtPromos.AcceptChanges()
      WriteDebug("End Evaluate Preferences", DebugState.CurrentTime)
    End If

  End Sub

  Private Function GetPendingInfo(ByVal CustomerId As String) As String
    Dim sXml As String = ""
    Dim dtPendingPoints As DataTable = Nothing
    Dim dtPendingDistribution As DataTable = Nothing
    Dim dtPendingCPERewardDistribution As DataTable = Nothing
    Dim dtPendingRewardLimits As DataTable = Nothing
    Dim lCustomerPk As Long = 0
    Dim lHhPk As Long = 0
    Dim lUseCustomerPk As Long
    Dim sMaskedId As String
    Dim sResponse As String = sOkStatus
    Dim sDebugMsg As String
    Dim bTestCustomer As Boolean = False

    Dim bCardFound As Boolean

    Dim sMemberId As String = ""

    Try
      MyCommon = New Copient.CommonInc
      MyCommon.AppName = sAppName

      sCurrentMethod = "GetPendingInfo"

      sMaskedId = Mask(CustomerId)

      sInputForLog = "*(Type=Input) (Method=GetPendingInfo)] - (CustomerId='" & sMaskedId & "')"

      WriteDebug("GetAccountPending", DebugState.BeginTime)

      MyCommon.Open_LogixXS()
      eDefaultErrorType = ErrorType.General

      ' obtain CustomerPK for this Card
      bCardFound = GetCustomerPK(CustomerId, 0, lCustomerPk, lHhPk, bTestCustomer)
      If Not bCardFound Then
        sXml = ConvertMemberNotFoundToXml(CustomerId, 0)
        Exit Try
      End If

      sMemberId = CustomerId

      If lHhPk = 0 Then
        lUseCustomerPk = lCustomerPk
        sDebugMsg = "got CustomerPK (" & lCustomerPk & ") via ID (" & sMaskedId & ")"
      Else
        lUseCustomerPk = lHhPk
        sDebugMsg = "got CustomerPK (" & lCustomerPk & ") via ID (" & sMaskedId & "), using HouseholdPK (" & lHhPk & ")"
      End If
      WriteDebug(sDebugMsg, DebugState.CurrentTime)

      ' If Pending Points is enabled, return pending point values
      If MyCommon.Fetch_SystemOption(251) = "1" Then
        MyCommon.QueryStr = "select PromoVarID, EarnedAmount, RedeemedAmount, ProgramID, CartID, ExtLocationCode, POSTimeStamp " & _
                            "from PointsPending with (NoLock) where CustomerPK=" & lUseCustomerPk & " and Deleted=0"
        dtPendingPoints = MyCommon.LXS_Select

        MyCommon.QueryStr = "select PromoVarID, Amount, CartID, ExtLocationCode, POSTimeStamp " & _
                            "from DistributionVariablesPending with (NoLock) where CustomerPK=" & lUseCustomerPk & " and Deleted=0"
        dtPendingDistribution = MyCommon.LXS_Select

        MyCommon.QueryStr = "select IncentiveID, RewardOptionID, CartID, ExtLocationCode, POSTimeStamp " & _
                            "from CPE_RewardDistributionPending with (NoLock) where CustomerPK=" & lUseCustomerPk & " and Deleted=0"
        dtPendingCPERewardDistribution = MyCommon.LXS_Select
        
        MyCommon.QueryStr = "select PromoVarID, Amount, CartID, ExtLocationCode, POSTimeStamp " & _
                            "from RewardLimitVariablesPending with (NoLock) where CustomerPK=" & lUseCustomerPk & " and Deleted=0"
        dtPendingRewardLimits = MyCommon.LXS_Select
      End If  

      WriteDebug("got pending data", DebugState.CurrentTime)

      sXml = ConvertPendingInfoToXml(sMemberId, dtPendingPoints, dtPendingDistribution, dtPendingCPERewardDistribution, dtPendingRewardLimits)

    Catch exApp As ApplicationException
      sXml = BuildErrorXml(exApp.Message, ErrorType.Db2General, False)
    Catch ex As Exception
      sXml = BuildErrorXml(ex.Message, eDefaultErrorType)
    Finally
      If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then
        MyCommon.Close_LogixXS()
      End If
      WriteDebug("GetAccountPending", DebugState.EndTime)
      If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then
        MyCommon.Close_LogixRT()
      End If
    End Try
    Return sXml
  End Function

  Private Function DeletePendingInfo(ByVal sCartID As String) As String
    Dim sResponse As String = sOkStatus
    Dim dt As DataTable
    Dim lTargetetLocationId As Long = 0

    Try
      MyCommon = New Copient.CommonInc
      MyCommon.AppName = sAppName

      MyCommon.Open_LogixXS()
      eDefaultErrorType = ErrorType.General

      If MyCommon.Fetch_SystemOption(219) = "1" Then
        ' share data between CM & UE
        ' get Broker LocationID
        MyCommon.QueryStr = "select LocationID from Locations with (NoLock) where Deleted=0 and EngineID=9 and LocationTypeID=2;"
        dt = MyCommon.LRT_Select
        If (dt.Rows.Count > 0) Then
          lTargetetLocationId = dt.Rows(0).Item(0)
        End If
        If lTargetetLocationId = 0 Then
          WriteLog("Invalid LocationID was found for UE Broker, so the deletion of pending data for CartID '" & sCartID & "' was not sent to the broker!", MessageType.AppError)
        End If
      End If

      If lTargetetLocationId > 0 Then
        MyCommon.QueryStr = "dbo.pt_PendingDeleteByCartID_UpdateBroker"
      Else
        MyCommon.QueryStr = "dbo.pt_PendingDeleteByCartID"
      End If
      MyCommon.Open_LXSsp()
      MyCommon.LXSsp.Parameters.Add("@TableNum", SqlDbType.VarChar, 4).Value = ""
      MyCommon.LXSsp.Parameters.Add("@Operation", SqlDbType.VarChar, 2).Value = ""
      MyCommon.LXSsp.Parameters.Add("@Col1", SqlDbType.VarChar, 36).Value = sCartID
      MyCommon.LXSsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = 0
      MyCommon.LXSsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = lTargetetLocationId
      MyCommon.LXSsp.Parameters.Add("@WaitingACK", SqlDbType.Int).Value = 0
      MyCommon.LXSsp.ExecuteNonQuery()
      MyCommon.Close_LXSsp()
      WriteLog("Deleted CartID " & sCartID, MessageType.Info)
    Catch ex As Exception
      sResponse = BuildErrorXml(ex.Message, eDefaultErrorType)
      WriteLog("Delete CartID error " & sCartID, MessageType.AppError)
    Finally
      If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then
        MyCommon.Close_LogixXS()
      End If
    End Try
    Return sResponse
  End Function

End Class