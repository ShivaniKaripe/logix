Imports System
Imports System.Web.Services
Imports System.Data
Imports System.IO
Imports Copient.CommonInc
Imports Copient
Imports CMS.AMS
Imports CMS.AMS.Contract
Imports System.Diagnostics

<WebService(Namespace:="http://www.copienttech.com/UniversalOfferConnector/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class UniversalOfferConnector
    Inherits System.Web.Services.WebService

    ' Return error codes
    Private Enum ERROR_CODES As Integer
        ERROR_NONE = 0
        ERROR_APPLICATION = 1
        ERROR_OFFER_DOES_NOT_EXIST = 2
        ERROR_OFFER_ENGINE_ID_INVALID = 3
        ERROR_INVALID_CRMENGINE = 4
        ERROR_INCORRECT_DATE_FORMAT = 5
        ERROR_INVALID_GUID_OR_EXTINTERFACEID = 6
        ERROR_NO_OFFER_FOUND_FOR_SEARCH_CRITERIA = 7
        ERROR_DEFAULT_OFFER_LIMIT_NOT_SET_CORRECTLY = 8
        ERROR_AUXILARY_NOT_APPLIED_WITH_CM_ENGINE = 9
        'added
        ERROR_INVALID_CARDID = 10
        ERROR_INVALID_CARDTYPEID = 11
        ERROR_NOT_FOUND_CUSTOMER = 12
        ERROR_INVALID_PRODUCTID = 13
        ERROR_INVALID_PRODUCTTYPEID = 14
        ERROR_INVALID_EXTLOCATIONCODE = 16
        ERROR_OFFER_AWAITING_RECOMMENDATIONS = 17
    End Enum

    Private Enum RESPONSE_TYPES As Integer
        GET_OFFERDATA = 1
        GET_OFFERS = 2
        'added
        GET_OFFERDATA_AUXILARY = 3
        GET_OFFERLIST_WITH_FILTERS = 4
        GET_OFFERLIST_WITH_FILTER_AUXILARY = 5
        GET_OFFERS_AUXILARY = 6
    End Enum

    Private Enum ENGINE_ID As Integer
        CM = 0
        CPE = 2
        UE = 9
    End Enum

    Private Enum TableTypeEnum
        MAINTAINABLE = 0
        DEPLOYED = 1
    End Enum

    Private TableType As TableTypeEnum = TableTypeEnum.MAINTAINABLE
    Private MyCommon As New Copient.CommonInc
    Private MyCryptLib As New Copient.CryptLib

    Private ExtInterfaceID As Long
    Private UOCLogFile As String = "UOCLog." & Date.Now.ToString("yyyyMMdd") & ".txt"
    Private UOCErrorLogFile As String = "UOCErrorLog." & Date.Now.ToString("yyyyMMdd") & ".txt"
    Private m_Offers As IGiftCardRewardService
    Private m_OfferProximity As IProximityMessageRewardService
    Private m_AnalyticsCGService As IAnalyticsCustomerGroups
    Private m_ExportXMLUE As ExportXmlUE

    <WebMethod()> _
    Public Function GetOfferData(ByVal Guid As String, ByVal ExtInterfaceID As Integer, ByVal EngineID As Integer, ByVal OfferID As Long) As String
        Dim ResponseXml As String = ""
        Dim OfferXml As String = ""
        Dim ErrorXml As New StringBuilder()
        Dim ErrorCode As ERROR_CODES = ERROR_CODES.ERROR_NONE
        Dim ErrorMsg As String = ""
        Dim OfferEngineId As Long = -1
        Dim MyExportXml As New Copient.ExportXml
        Dim MyExportXmlCpe As New Copient.ExportXmlCPE
        Dim dt As DataTable = Nothing
        Dim TempFileName As String
        TempFileName = System.AppDomain.CurrentDomain.BaseDirectory() & "Connectors\TempOfferFile" & Date.Now.ToString("yyyyMMdd") & OfferID.ToString & ".txt"

        Try
            Startup()
            If ValidateExtInterfaceID(Guid, ExtInterfaceID) Then
                If MyCommon.LRTadoConn.State = Data.ConnectionState.Closed Then MyCommon.Open_LogixRT()

                Select Case EngineID
                    Case ENGINE_ID.CM
                        MyCommon.QueryStr = "select OfferID, EngineID from Offers with (NoLock) where OfferID = @OfferID " & _
                                              "and InboundCRMEngineID = @InboundCRMEngineID and EngineID = @EngineID and Deleted=0"
                        MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
                        MyCommon.DBParameters.Add("@InboundCRMEngineID", SqlDbType.Int).Value = ExtInterfaceID
                        MyCommon.DBParameters.Add("@EngineID", SqlDbType.Int).Value = EngineID
                        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                        If (dt.Rows.Count > 0) Then

                            If MyExportXml.GenOfferXml(OfferID.ToString, TempFileName, True, 0) Then
                                OfferXml = File.ReadAllText(TempFileName)
                                File.Delete(TempFileName)
                            End If
                        Else
                            ErrorCode = ERROR_CODES.ERROR_OFFER_DOES_NOT_EXIST
                            ErrorMsg = "Offer " & OfferID & " does not exist for the ExtInterfaceID."
                        End If

                    Case ENGINE_ID.CPE
                        MyCommon.QueryStr = "select IncentiveID as OfferID, EngineId from CPE_Incentives with (NoLock) where IncentiveID = @IncentiveID " & _
                                            "and InboundCRMEngineID = @InboundCRMEngineID and EngineID = @EngineID and Deleted=0"
                        MyCommon.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = OfferID
                        MyCommon.DBParameters.Add("@InboundCRMEngineID", SqlDbType.Int).Value = ExtInterfaceID
                        MyCommon.DBParameters.Add("@EngineID", SqlDbType.Int).Value = EngineID
                        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                        If (dt.Rows.Count > 0) Then
                            OfferXml = MyExportXmlCpe.GetOfferXML(OfferID)
                        Else
                            ErrorCode = ERROR_CODES.ERROR_OFFER_DOES_NOT_EXIST
                            ErrorMsg = "Offer " & OfferID & " does not exist for the ExtInterfaceID."
                        End If

                    Case ENGINE_ID.UE
                        If m_AnalyticsCGService.CanExport(OfferID) Then
                            MyCommon.QueryStr = "select IncentiveID as OfferID, EngineId from CPE_Incentives with (NoLock) where IncentiveID = @IncentiveID " &
                                            "and InboundCRMEngineID = @InboundCRMEngineID and EngineID = @EngineID and Deleted=0"
                            MyCommon.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = OfferID
                            MyCommon.DBParameters.Add("@InboundCRMEngineID", SqlDbType.Int).Value = ExtInterfaceID
                            MyCommon.DBParameters.Add("@EngineID", SqlDbType.Int).Value = EngineID
                            dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                            If (dt.Rows.Count > 0) Then

                                OfferXml = m_ExportXMLUE.GetOfferXML(OfferID)
                            Else
                                ErrorCode = ERROR_CODES.ERROR_OFFER_DOES_NOT_EXIST
                                ErrorMsg = "Offer " & OfferID & " does not exist for the ExtInterfaceID."
                            End If
                        Else
                            ErrorCode = ERROR_CODES.ERROR_OFFER_AWAITING_RECOMMENDATIONS
                            ErrorMsg = "Offer " & OfferID & " is awaiting recommendations and cannot be exported."
                        End If

                    Case Else
                        ErrorCode = ERROR_CODES.ERROR_OFFER_ENGINE_ID_INVALID
                        ErrorMsg = "Invalid Offer EngineID: " & OfferEngineId
                End Select
            Else
                ErrorCode = ERROR_CODES.ERROR_INVALID_GUID_OR_EXTINTERFACEID
                ErrorMsg = "Invalid GUID or ExternalInterfaceID"
            End If

        Catch ex As Exception
            ErrorCode = ERROR_CODES.ERROR_APPLICATION
            ErrorMsg = "Application error encountering: " & ex.ToString
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try

        If (ErrorCode = ERROR_CODES.ERROR_NONE) Then
            ResponseXml = OfferXml.Replace("encoding=""utf-16""", "encoding=""utf-8""")
            OfferXml = Nothing
        Else
            ' create error xml
            ErrorXml.Append("<?xml version=""1.0"" encoding=""UTF-8""?>" & ControlChars.CrLf)
            ErrorXml.Append("<UniversalOfferConnector>" & ControlChars.CrLf)
            ErrorXml.Append("  <Offer id=""" & OfferID & """ operation=""" & RESPONSE_TYPES.GET_OFFERDATA.ToString & """ success=""false"" />" & ControlChars.CrLf)
            ErrorXml.Append("  <Error code=""" & ErrorCode.ToString & """ message=""" & ErrorMsg & """ />" & ControlChars.CrLf)
            ErrorXml.Append("</UniversalOfferConnector>")

            ResponseXml = ErrorXml.ToString
        End If

        AppendToLog(ExtInterfaceID.ToString, "GetOfferData", EngineID.ToString, "OfferID=" & OfferID, ErrorCode, ErrorMsg)

        Return ResponseXml
    End Function

    Private Sub Startup()
        'Resolving the dependency on the giftcardreward in ExportXMLUE
        CurrentRequest.Resolver.AppName = "UniversalOfferConnector"
        m_Offers = CurrentRequest.Resolver.Resolve(Of IGiftCardRewardService)()
        m_OfferProximity = CurrentRequest.Resolver.Resolve(Of IProximityMessageRewardService)()
        m_ExportXMLUE = CurrentRequest.Resolver.Resolve(Of ExportXmlUE)()
        m_AnalyticsCGService = CurrentRequest.Resolver.Resolve(Of IAnalyticsCustomerGroups)()
    End Sub

    <WebMethod()> _
    Public Function GetOffers(ByVal Guid As String, ByVal ExtInterfaceID As Integer, ByVal EngineID As Integer, ByVal StartDate As String, ByVal EndDate As String) As String
        Dim ResponseXml As String = ""
        Dim OfferXml As String = ""
        Dim ErrorXml As New StringBuilder()
        Dim ErrorCode As ERROR_CODES = ERROR_CODES.ERROR_NONE
        Dim ErrorMsg As String = ""
        Dim OfferEngineId As Long = -1
        Dim MaxOffers As Integer = 1000
        Static RequestTime As DateTime = Now
        Dim ElapsedTime As TimeSpan

        Try
            Startup()
            If (StartDate = "" And EndDate = "") Or (IsDate(StartDate) And StartDate.Length >= 10 And IsDate(EndDate) And EndDate.Length >= 10) Then
                If ValidateExtInterfaceID(Guid, ExtInterfaceID) Then
                    Select Case EngineID

                        Case ENGINE_ID.CM, ENGINE_ID.CPE, ENGINE_ID.UE

                            MaxOffers = GetDefaultOfferLimit()
                            If MaxOffers >= 0 Then
                                OfferXml = OffersXML(StartDate, EndDate, EngineID, ExtInterfaceID, MaxOffers)
                                If OfferXml.Length = 0 Then
                                    ErrorCode = ERROR_CODES.ERROR_NO_OFFER_FOUND_FOR_SEARCH_CRITERIA
                                    ErrorMsg = "No Offers Found for The Search Criteria Entered."
                                End If

                            Else
                                ErrorCode = ERROR_CODES.ERROR_DEFAULT_OFFER_LIMIT_NOT_SET_CORRECTLY
                                ErrorMsg = "System Option Default Offer Limit Not Set Correctly."
                            End If

                        Case Else
                            ErrorCode = ERROR_CODES.ERROR_OFFER_ENGINE_ID_INVALID
                            ErrorMsg = "Invalid Offer EngineID: " & OfferEngineId
                    End Select
                Else
                    ErrorCode = ERROR_CODES.ERROR_INVALID_GUID_OR_EXTINTERFACEID
                    ErrorMsg = "Invalid GUID or ExternalInterfaceID"
                End If
            Else
                ErrorCode = ERROR_CODES.ERROR_INCORRECT_DATE_FORMAT
                ErrorMsg = "Either the format of StartDate or EndDate is incorrect."
            End If
        Catch ex As Exception
            ErrorCode = ERROR_CODES.ERROR_APPLICATION
            ErrorMsg = "Application error encountering: " & ex.ToString
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try

        If (ErrorCode = ERROR_CODES.ERROR_NONE) Then
            ResponseXml = OfferXml.Replace("encoding=""utf-16""", "encoding=""utf-8""")
            OfferXml = Nothing
        Else
            ' create error xml
            ErrorXml.Append("<?xml version=""1.0"" encoding=""UTF-8""?>" & ControlChars.CrLf)
            ErrorXml.Append("<UniversalOfferConnector>" & ControlChars.CrLf)
            ErrorXml.Append("  <Offer id=""" & """ operation=""" & RESPONSE_TYPES.GET_OFFERS.ToString & """ success=""false"" />" & ControlChars.CrLf)
            ErrorXml.Append("  <Error code=""" & ErrorCode.ToString & """ message=""" & ErrorMsg & """ />" & ControlChars.CrLf)
            ErrorXml.Append("</UniversalOfferConnector>")

            ResponseXml = ErrorXml.ToString
        End If

        ElapsedTime = Now.Subtract(RequestTime)
        AppendToLog(ExtInterfaceID, "GetOffers", EngineID.ToString, "GetOffers Count=" & MaxOffers & "; Elapsed Time= " & ElapsedTime.ToString, ErrorCode, ErrorMsg)

        Return ResponseXml
    End Function
<WebMethod()> _
    Public Function GetAnyCustomerOffers(ByVal Guid As String, ByVal ExtInterfaceID As Integer, ByVal EngineID As Integer, ByVal StartDate As String, ByVal EndDate As String) As String
        Dim ResponseXml As String = ""
        Dim OfferXml As String = ""
        Dim ErrorXml As New StringBuilder()
        Dim ErrorCode As ERROR_CODES = ERROR_CODES.ERROR_NONE
        Dim ErrorMsg As String = ""
        Dim OfferEngineId As Long = -1
        Dim MaxOffers As Integer = 1000
        Static RequestTime As DateTime = Now
        Dim ElapsedTime As TimeSpan
        Dim CustGrpID As Integer = 1

        Try
            Startup()
            If (StartDate = "" And EndDate = "") Or (IsDate(StartDate) And StartDate.Length >= 10 And IsDate(EndDate) And EndDate.Length >= 10) Then
                If ValidateExtInterfaceID(Guid, ExtInterfaceID) Then
                    Select Case EngineID

                        Case ENGINE_ID.CM, ENGINE_ID.CPE, ENGINE_ID.UE

                            MaxOffers = GetDefaultOfferLimit()
                            If MaxOffers >= 0 Then
                                ' OfferXml = OffersXML(StartDate, EndDate, EngineID, ExtInterfaceID, MaxOffers, IncludeChannelData, IncludeUserDefinedFields, IncludeMultiLanguageReceipt, 1)

                                OfferXml = OffersXML(StartDate, EndDate, EngineID, ExtInterfaceID, MaxOffers, CustGrpID)
                                ' AnyCustomerOffersXML(ByVal StartDate As String, ByVal EndDate As String, ByVal OfferEngineID As Integer, ByVal ExtInterfaceID As Integer, ByRef MaxOffers As Integer, Optional ByVal IncludeChannelData As Boolean = False, Optional ByVal IncludeUserDefinedFields As Boolean = False, Optional ByVal IncludeMultiLanguageReceipt As Boolean = False, Optional ByVal CustGrpID As Integer = 0)
                                If OfferXml.Length = 0 Then
                                    ErrorCode = ERROR_CODES.ERROR_NO_OFFER_FOUND_FOR_SEARCH_CRITERIA
                                    ErrorMsg = "No Offers Found for The Search Criteria Entered."
                                End If

                            Else
                                ErrorCode = ERROR_CODES.ERROR_DEFAULT_OFFER_LIMIT_NOT_SET_CORRECTLY
                                ErrorMsg = "System Option Default Offer Limit Not Set Correctly."
                            End If

                        Case Else
                            ErrorCode = ERROR_CODES.ERROR_OFFER_ENGINE_ID_INVALID
                            ErrorMsg = "Invalid Offer EngineID: " & OfferEngineId
                    End Select
                Else
                    ErrorCode = ERROR_CODES.ERROR_INVALID_GUID_OR_EXTINTERFACEID
                    ErrorMsg = "Invalid GUID or ExternalInterfaceID"
                End If
            Else
                ErrorCode = ERROR_CODES.ERROR_INCORRECT_DATE_FORMAT
                ErrorMsg = "Either the format of StartDate or EndDate is incorrect."
            End If
        Catch ex As Exception
            ErrorCode = ERROR_CODES.ERROR_APPLICATION
            ErrorMsg = "Application error encountering: " & ex.ToString
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try

        If (ErrorCode = ERROR_CODES.ERROR_NONE) Then
            ResponseXml = OfferXml.Replace("encoding=""utf-16""", "encoding=""utf-8""")
            OfferXml = Nothing
        Else
            ' create error xml
            ErrorXml.Append("<?xml version=""1.0"" encoding=""UTF-8""?>" & ControlChars.CrLf)
            ErrorXml.Append("<UniversalOfferConnector>" & ControlChars.CrLf)
            ErrorXml.Append("  <Offer id=""" & """ operation=""" & RESPONSE_TYPES.GET_OFFERS.ToString & """ success=""false"" />" & ControlChars.CrLf)
            ErrorXml.Append("  <Error code=""" & ErrorCode.ToString & """ message=""" & ErrorMsg & """ />" & ControlChars.CrLf)
            ErrorXml.Append("</UniversalOfferConnector>")

            ResponseXml = ErrorXml.ToString
        End If

        ElapsedTime = Now.Subtract(RequestTime)
        AppendToLog(ExtInterfaceID, "GetOffers", EngineID.ToString, "GetOffers Count=" & MaxOffers & "; Elapsed Time= " & ElapsedTime.ToString, ErrorCode, ErrorMsg)

        Return ResponseXml
    End Function
    <WebMethod()> _
    Public Function GetAnyCardHolderOffers(ByVal Guid As String, ByVal ExtInterfaceID As Integer, ByVal EngineID As Integer, ByVal StartDate As String, ByVal EndDate As String) As String
        Dim ResponseXml As String = ""
        Dim OfferXml As String = ""
        Dim ErrorXml As New StringBuilder()
        Dim ErrorCode As ERROR_CODES = ERROR_CODES.ERROR_NONE
        Dim ErrorMsg As String = ""
        Dim OfferEngineId As Long = -1
        Dim MaxOffers As Integer = 1000
        Static RequestTime As DateTime = Now
        Dim ElapsedTime As TimeSpan
        Dim CustGrpID As Integer = 2

        Try
            Startup()
            If (StartDate = "" And EndDate = "") Or (IsDate(StartDate) And StartDate.Length >= 10 And IsDate(EndDate) And EndDate.Length >= 10) Then
                If ValidateExtInterfaceID(Guid, ExtInterfaceID) Then
                    Select Case EngineID

                        Case ENGINE_ID.CM, ENGINE_ID.CPE, ENGINE_ID.UE

                            MaxOffers = GetDefaultOfferLimit()
                            If MaxOffers >= 0 Then

                                OfferXml = OffersXML(StartDate, EndDate, EngineID, ExtInterfaceID, MaxOffers, CustGrpID)

                                If OfferXml.Length = 0 Then
                                    ErrorCode = ERROR_CODES.ERROR_NO_OFFER_FOUND_FOR_SEARCH_CRITERIA
                                    ErrorMsg = "No Offers Found for The Search Criteria Entered."
                                End If

                            Else
                                ErrorCode = ERROR_CODES.ERROR_DEFAULT_OFFER_LIMIT_NOT_SET_CORRECTLY
                                ErrorMsg = "System Option Default Offer Limit Not Set Correctly."
                            End If

                        Case Else
                            ErrorCode = ERROR_CODES.ERROR_OFFER_ENGINE_ID_INVALID
                            ErrorMsg = "Invalid Offer EngineID: " & OfferEngineId
                    End Select
                Else
                    ErrorCode = ERROR_CODES.ERROR_INVALID_GUID_OR_EXTINTERFACEID
                    ErrorMsg = "Invalid GUID or ExternalInterfaceID"
                End If
            Else
                ErrorCode = ERROR_CODES.ERROR_INCORRECT_DATE_FORMAT
                ErrorMsg = "Either the format of StartDate or EndDate is incorrect."
            End If
        Catch ex As Exception
            ErrorCode = ERROR_CODES.ERROR_APPLICATION
            ErrorMsg = "Application error encountering: " & ex.ToString
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try

        If (ErrorCode = ERROR_CODES.ERROR_NONE) Then
            ResponseXml = OfferXml.Replace("encoding=""utf-16""", "encoding=""utf-8""")
            OfferXml = Nothing
        Else
            ' create error xml
            ErrorXml.Append("<?xml version=""1.0"" encoding=""UTF-8""?>" & ControlChars.CrLf)
            ErrorXml.Append("<UniversalOfferConnector>" & ControlChars.CrLf)
            ErrorXml.Append("  <Offer id=""" & """ operation=""" & RESPONSE_TYPES.GET_OFFERS.ToString & """ success=""false"" />" & ControlChars.CrLf)
            ErrorXml.Append("  <Error code=""" & ErrorCode.ToString & """ message=""" & ErrorMsg & """ />" & ControlChars.CrLf)
            ErrorXml.Append("</UniversalOfferConnector>")

            ResponseXml = ErrorXml.ToString
        End If

        ElapsedTime = Now.Subtract(RequestTime)
        AppendToLog(ExtInterfaceID, "GetOffers", EngineID.ToString, "GetOffers Count=" & MaxOffers & "; Elapsed Time= " & ElapsedTime.ToString, ErrorCode, ErrorMsg)

        Return ResponseXml
    End Function
    <WebMethod()> _
    Public Function GetOfferDataAuxilary(ByVal Guid As String, ByVal ExtInterfaceID As String, ByVal EngineID As String, ByVal OfferID As String, ByVal IncludeAuxilaryDetailsCustProdLoc As String, ByVal IncludeAuxilaryOthers As String) As String
        Dim iExtInterfaceID As Integer = -1
        Dim iEngineID As Integer = -1
        Dim lOfferID As Long = -1
       
        Dim ErrorCode As ERROR_CODES = ERROR_CODES.ERROR_NONE
        Dim ErrorMsg As String = ""
        Dim ResponseXml As String = ""
        Dim ErrorXml As New StringBuilder()
        Try
            If Not String.IsNullOrEmpty(ExtInterfaceID) AndAlso Not ExtInterfaceID.Trim = String.Empty AndAlso IsNumeric(ExtInterfaceID) Then iExtInterfaceID = Convert.ToInt32(ExtInterfaceID)
            If Not String.IsNullOrEmpty(EngineID) AndAlso Not EngineID.Trim = String.Empty AndAlso IsNumeric(EngineID) Then iEngineID = Convert.ToInt32(EngineID) Else iEngineID = Convert.ToInt32(EngineID)
            If Not String.IsNullOrEmpty(OfferID) AndAlso Not OfferID.Trim = String.Empty AndAlso IsNumeric(OfferID) Then lOfferID = Convert.ToInt64(OfferID) Else lOfferID = Convert.ToInt64(OfferID)
        Catch e As FormatException
            Dim st As New StackTrace(True)
            Dim sf As StackFrame = st.GetFrame(0)
            If (sf.GetFileLineNumber() = 265) Then 'line no 265 for engineid formatexception
                ErrorCode = ERROR_CODES.ERROR_OFFER_ENGINE_ID_INVALID
                ErrorMsg = "Invalid Offer EngineID: " & EngineID
            Else
                ErrorCode = ERROR_CODES.ERROR_OFFER_DOES_NOT_EXIST
                ErrorMsg = "Offer " & OfferID & " does not exist for the ExtInterfaceID."
            End If
            ErrorXml.Append("<?xml version=""1.0"" encoding=""UTF-8""?>" & ControlChars.CrLf)
            ErrorXml.Append("<UniversalOfferConnector>" & ControlChars.CrLf)
            ErrorXml.Append("  <Offer id=""" & OfferID & """ operation=""" & RESPONSE_TYPES.GET_OFFERDATA_AUXILARY.ToString & """ success=""false"" />" & ControlChars.CrLf)
            ErrorXml.Append("  <Error code=""" & ErrorCode.ToString & """ message=""" & ErrorMsg & """ />" & ControlChars.CrLf)
            ErrorXml.Append("</UniversalOfferConnector>")
            ResponseXml = ErrorXml.ToString
            AppendToLog(ExtInterfaceID.ToString, "GetOfferData", EngineID.ToString, "OfferID=" & OfferID, ErrorCode, ErrorMsg)
        End Try
        If (ErrorCode = ERROR_CODES.ERROR_NONE) Then
            Return _GetOfferDataAuxilary(Guid, iExtInterfaceID, iEngineID, lOfferID, IncludeAuxilaryDetailsCustProdLoc, IncludeAuxilaryOthers)
        End If
        Return ResponseXml
    End Function

    Private Function _GetOfferDataAuxilary(ByVal Guid As String, ByVal ExtInterfaceID As Integer, ByVal EngineID As Integer, ByVal OfferID As Long, ByVal IncludeAuxilaryDetailsCustProdLoc As String, ByVal IncludeAuxilaryOthers As String) As String
        Dim ResponseXml As String = ""
        Dim OfferXml As String = ""
        Dim ErrorXml As New StringBuilder()
        Dim ErrorCode As ERROR_CODES = ERROR_CODES.ERROR_NONE
        Dim ErrorMsg As String = ""
        ' Dim OfferEngineId As Long = -1
        Dim MyExportXmlCpe As New Copient.ExportXmlCPE
        Dim dt As DataTable = Nothing
        Dim TempFileName As String
        TempFileName = System.AppDomain.CurrentDomain.BaseDirectory() & "Connectors\TempOfferFile" & Date.Now.ToString("yyyyMMdd") & OfferID.ToString & ".txt"
        Try
            Startup()
            If ValidateExtInterfaceID(Guid, ExtInterfaceID) Then
                If ValidateAuxilaryDetails(IncludeAuxilaryDetailsCustProdLoc, IncludeAuxilaryOthers) Then
                    If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
                    Select Case EngineID
                        Case ENGINE_ID.CM
                            ErrorCode = ERROR_CODES.ERROR_AUXILARY_NOT_APPLIED_WITH_CM_ENGINE
                            ErrorMsg = "Auxiliary is not applied with CM engine, call GetOfferData method instead."

                        Case ENGINE_ID.CPE
                            MyCommon.QueryStr = "select IncentiveID as OfferID, EngineId from CPE_Incentives with (NoLock) where IncentiveID = @IncentiveID " & _
                                                "and InboundCRMEngineID = @InboundCRMEngineID and EngineID = @EngineID and Deleted=0"
                            MyCommon.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = OfferID
                            MyCommon.DBParameters.Add("@InboundCRMEngineID", SqlDbType.Int).Value = ExtInterfaceID
                            MyCommon.DBParameters.Add("@EngineID", SqlDbType.Int).Value = EngineID
                            dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                            If (dt.Rows.Count > 0) Then
                                OfferXml = MyExportXmlCpe.GetOfferXMLAuxilary(OfferID, IncludeAuxilaryDetailsCustProdLoc, IncludeAuxilaryOthers)
                            Else
                                ErrorCode = ERROR_CODES.ERROR_OFFER_DOES_NOT_EXIST
                                ErrorMsg = "Offer " & OfferID & " does not exist for the ExtInterfaceID."
                            End If

                        Case ENGINE_ID.UE
                            If m_AnalyticsCGService.CanExport(OfferID) Then
                                MyCommon.QueryStr = "select IncentiveID as OfferID, EngineId from CPE_Incentives with (NoLock) where IncentiveID = @IncentiveID " &
                                                "and InboundCRMEngineID = @InboundCRMEngineID and EngineID = @EngineID and Deleted=0"
                                MyCommon.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = OfferID
                                MyCommon.DBParameters.Add("@InboundCRMEngineID", SqlDbType.Int).Value = ExtInterfaceID
                                MyCommon.DBParameters.Add("@EngineID", SqlDbType.Int).Value = EngineID
                                dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                                If (dt.Rows.Count > 0) Then
                                    OfferXml = m_ExportXMLUE.GetOfferXMLAuxilary(OfferID, IncludeAuxilaryDetailsCustProdLoc, IncludeAuxilaryOthers)
                                Else
                                    ErrorCode = ERROR_CODES.ERROR_OFFER_DOES_NOT_EXIST
                                    ErrorMsg = "Offer " & OfferID & " does not exist for the ExtInterfaceID."
                                End If
                            Else
                                ErrorCode = ERROR_CODES.ERROR_OFFER_AWAITING_RECOMMENDATIONS
                                ErrorMsg = "Offer " & OfferID & " is awaiting recommendations and cannot be exported."
                            End If


                        Case Else
                            ErrorCode = ERROR_CODES.ERROR_OFFER_ENGINE_ID_INVALID
                            ErrorMsg = "Invalid Offer EngineID: " & EngineID '& OfferEngineId
                    End Select

                Else
                    ErrorCode = ERROR_CODES.ERROR_AUXILARY_NOT_APPLIED_WITH_CM_ENGINE
                    ErrorMsg = "Fields IncludeAuxilaryDetailsCustProdLoc\IncludeAuxilaryOthers accept only boolean vlaue(True or False)."
                End If

            Else
                ErrorCode = ERROR_CODES.ERROR_INVALID_GUID_OR_EXTINTERFACEID
                ErrorMsg = "Invalid GUID or ExternalInterfaceID"
            End If

        Catch ex As Exception
            ErrorCode = ERROR_CODES.ERROR_APPLICATION
            ErrorMsg = "Application error encountering: " & ex.ToString
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try

        If (ErrorCode = ERROR_CODES.ERROR_NONE) Then
            ResponseXml = OfferXml.Replace("encoding=""utf-16""", "encoding=""utf-8""")
            OfferXml = Nothing
        Else
            ' create error xml
            ErrorXml.Append("<?xml version=""1.0"" encoding=""UTF-8""?>" & ControlChars.CrLf)
            ErrorXml.Append("<UniversalOfferConnector>" & ControlChars.CrLf)
            ErrorXml.Append("  <Offer id=""" & OfferID & """ operation=""" & RESPONSE_TYPES.GET_OFFERDATA_AUXILARY.ToString & """ success=""false"" />" & ControlChars.CrLf)
            ErrorXml.Append("  <Error code=""" & ErrorCode.ToString & """ message=""" & ErrorMsg & """ />" & ControlChars.CrLf)
            ErrorXml.Append("</UniversalOfferConnector>")
            ResponseXml = ErrorXml.ToString
        End If
        AppendToLog(ExtInterfaceID.ToString, "GetOfferData", EngineID.ToString, "OfferID=" & OfferID, ErrorCode, ErrorMsg)
        Return ResponseXml
    End Function

    <WebMethod()> _
    Public Function GetOffersAuxilary(ByVal Guid As String, ByVal ExtInterfaceID As String, ByVal EngineID As String, ByVal StartDate As String, ByVal EndDate As String, ByVal IncludeAuxilaryDetailsCustProdLoc As String, ByVal IncludeAuxilaryOthers As String) As String
        Dim iExtInterfaceID As Integer = -1
        'Dim iEngineID As Integer = -1
        StartDate = StartDate.Trim
        EndDate = EndDate.Trim
        Startup()
        If Not String.IsNullOrEmpty(ExtInterfaceID) AndAlso Not ExtInterfaceID.Trim = String.Empty AndAlso IsNumeric(ExtInterfaceID) Then iExtInterfaceID = Convert.ToInt32(ExtInterfaceID)
        'If Not String.IsNullOrEmpty(EngineID) AndAlso Not EngineID.Trim = String.Empty AndAlso IsNumeric(EngineID) Then iEngineID = Convert.ToInt32(EngineID)
        Return _GetOffersAuxilary(Guid, iExtInterfaceID, EngineID, StartDate, EndDate, IncludeAuxilaryDetailsCustProdLoc, IncludeAuxilaryOthers)
    End Function

    Private Function _GetOffersAuxilary(ByVal Guid As String, ByVal ExtInterfaceID As Integer, ByVal EngineID As String, ByVal StartDate As String, ByVal EndDate As String, ByVal IncludeAuxilaryDetailsCustProdLoc As String, ByVal IncludeAuxilaryOthers As String) As String
        Dim ResponseXml As String = ""
        Dim OfferXml As String = ""
        Dim ErrorXml As New StringBuilder()
        Dim ErrorCode As ERROR_CODES = ERROR_CODES.ERROR_NONE
        Dim ErrorMsg As String = ""
        Dim OfferEngineId As Long = -1
        Dim MaxOffers As Integer = 1000
        Static RequestTime As DateTime = Now
        Dim ElapsedTime As TimeSpan
        Try
            Startup()
            If (StartDate = "" And EndDate = "") Or (IsDate(StartDate) And StartDate.Length >= 10 And IsDate(EndDate) And EndDate.Length >= 10) Then
                If ValidateExtInterfaceID(Guid, ExtInterfaceID) Then
                    If ValidateAuxilaryDetails(IncludeAuxilaryDetailsCustProdLoc, IncludeAuxilaryOthers) Then
                        Select Case EngineID
                            Case ENGINE_ID.CM
                                ErrorCode = ERROR_CODES.ERROR_AUXILARY_NOT_APPLIED_WITH_CM_ENGINE
                                ErrorMsg = "Auxiliary is not applied with CM engine, call GetOffers method instead."

                            Case ENGINE_ID.CPE, ENGINE_ID.UE
                                MaxOffers = GetDefaultOfferLimit()
                                If MaxOffers >= 0 Then
                                    Dim iEngineID As Integer = -1
                                    If Not String.IsNullOrEmpty(EngineID) AndAlso Not EngineID.Trim = String.Empty AndAlso IsNumeric(EngineID) Then iEngineID = Convert.ToInt32(EngineID)
                                    OfferXml = OffersXMLAuxilary(StartDate, EndDate, iEngineID, ExtInterfaceID, MaxOffers, IncludeAuxilaryDetailsCustProdLoc, IncludeAuxilaryOthers)
                                    If OfferXml.Length = 0 Then
                                        ErrorCode = ERROR_CODES.ERROR_NO_OFFER_FOUND_FOR_SEARCH_CRITERIA
                                        ErrorMsg = "No Offers Found for The Search Criteria Entered."
                                    End If

                                Else
                                    ErrorCode = ERROR_CODES.ERROR_DEFAULT_OFFER_LIMIT_NOT_SET_CORRECTLY
                                    ErrorMsg = "System Option Default Offer Limit Not Set Correctly."
                                End If

                            Case Else
                                ErrorCode = ERROR_CODES.ERROR_OFFER_ENGINE_ID_INVALID
                                ErrorMsg = "Invalid Offer EngineID: " & EngineID '& OfferEngineId
                        End Select
                    Else
                        ErrorCode = ERROR_CODES.ERROR_AUXILARY_NOT_APPLIED_WITH_CM_ENGINE
                        ErrorMsg = "Fields IncludeAuxilaryDetailsCustProdLoc or IncludeAuxilaryOthers accept only boolean vlaue(True or False)."
                    End If

                Else
                    ErrorCode = ERROR_CODES.ERROR_INVALID_GUID_OR_EXTINTERFACEID
                    ErrorMsg = "Invalid GUID or ExternalInterfaceID"
                End If
            Else
                ErrorCode = ERROR_CODES.ERROR_INCORRECT_DATE_FORMAT
                ErrorMsg = "Either the format of StartDate or EndDate is incorrect."
            End If
        Catch ex As Exception
            ErrorCode = ERROR_CODES.ERROR_APPLICATION
            ErrorMsg = "Application error encountering: " & ex.ToString
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try

        If (ErrorCode = ERROR_CODES.ERROR_NONE) Then
            ResponseXml = OfferXml.Replace("encoding=""utf-16""", "encoding=""utf-8""")
            OfferXml = Nothing
        Else
            ' create error xml
            ErrorXml.Append("<?xml version=""1.0"" encoding=""UTF-8""?>" & ControlChars.CrLf)
            ErrorXml.Append("<UniversalOfferConnector>" & ControlChars.CrLf)
            ErrorXml.Append("  <Offer id=""" & """ operation=""" & RESPONSE_TYPES.GET_OFFERS_AUXILARY.ToString & """ success=""false"" />" & ControlChars.CrLf)
            ErrorXml.Append("  <Error code=""" & ErrorCode.ToString & """ message=""" & ErrorMsg & """ />" & ControlChars.CrLf)
            ErrorXml.Append("</UniversalOfferConnector>")

            ResponseXml = ErrorXml.ToString
        End If

        ElapsedTime = Now.Subtract(RequestTime)
        AppendToLog(ExtInterfaceID, "GetOffers", EngineID.ToString, "GetOffers Count=" & MaxOffers & "; Elapsed Time= " & ElapsedTime.ToString, ErrorCode, ErrorMsg)

        Return ResponseXml
    End Function

    <WebMethod()> _
    Public Function GetOfferListWithFiltersAuxilary(ByVal Guid As String, ByVal ExtInterfaceID As String, ByVal EngineID As String, ByVal TransactionDate As String, ByVal CardID As String, ByVal CardTypeID As String, ByVal ProductID As String, ByVal ProductTypeID As String, ByVal ExtLocationCode As String, ByVal IncludeAuxilaryDetailsCustProdLoc As String, ByVal IncludeAuxilaryOthers As String) As String
        Dim iExtInterfaceID As Integer = -1
        'Dim iEngineID As Integer = -1
        'Dim iCardTypeID As Integer = -1
        Dim iProductTypeID As Integer = -1
        Dim ExtCardID As String = ""
        Dim ExtProductID As String = ""

        TransactionDate = TransactionDate.Trim
        ExtCardID = CardID.Trim
        ExtLocationCode = ExtLocationCode.Trim
        ExtProductID = ProductID.Trim

        If Not String.IsNullOrEmpty(ExtInterfaceID) AndAlso Not ExtInterfaceID.Trim = String.Empty AndAlso IsNumeric(ExtInterfaceID) Then iExtInterfaceID = Convert.ToInt32(ExtInterfaceID)
        'If Not String.IsNullOrEmpty(EngineID) AndAlso Not EngineID.Trim = String.Empty AndAlso IsNumeric(EngineID) Then iEngineID = Convert.ToInt32(EngineID)
        'If Not String.IsNullOrEmpty(CardTypeID) AndAlso Not CardTypeID.Trim = String.Empty AndAlso IsNumeric(CardTypeID) Then iCardTypeID = Convert.ToInt32(CardTypeID)
        If Not String.IsNullOrEmpty(ProductTypeID) AndAlso Not ProductTypeID.Trim = String.Empty AndAlso IsNumeric(ProductTypeID) Then iProductTypeID = Convert.ToInt32(ProductTypeID)

        Return _GetOfferListWithFiltersAuxilary(Guid, iExtInterfaceID, EngineID, TransactionDate, ExtCardID, CardTypeID, ExtProductID, iProductTypeID, ExtLocationCode, IncludeAuxilaryDetailsCustProdLoc, IncludeAuxilaryOthers)
    End Function

    Private Function _GetOfferListWithFiltersAuxilary(ByVal Guid As String, ByVal ExtInterfaceID As Integer, ByVal EngineID As String, ByVal TransactionDate As String, ByVal ExtCardID As String, ByVal CardTypeID As String, ByVal ExtProductID As String, ByVal ProductTypeID As Integer, ByVal ExtLocationCode As String, ByVal IncludeAuxilaryDetailsCustProdLoc As String, ByVal IncludeAuxilaryOthers As String) As String
        Dim ResponseXml As String = ""
        Dim OfferXml As String = ""
        Dim ErrorXml As New StringBuilder()
        Dim ErrorCode As ERROR_CODES = ERROR_CODES.ERROR_NONE
        Dim ErrorMsg As String = ""
        'Dim OfferEngineId As Long = -1
        Dim MaxOffers As Integer = 1000
        Static RequestTime As DateTime = Now
        Dim ElapsedTime As TimeSpan

        Try
            Startup()
            'If (TransactionDate = "") Or (IsDate(TransactionDate) And TransactionDate.Length >= 10) Then
            If (Not (TransactionDate = "") Or (IsDate(TransactionDate) And TransactionDate.Length >= 10)) And IsDate(TransactionDate) Then
                If ValidateExtInterfaceID(Guid, ExtInterfaceID) Then
                    If IsValidCustomerCard(ExtCardID, CardTypeID, ErrorCode, ErrorMsg) Then
                        Dim iCardTypeID As Integer = -1
                        If Not String.IsNullOrEmpty(CardTypeID) AndAlso Not CardTypeID.Trim = String.Empty AndAlso IsNumeric(CardTypeID) Then iCardTypeID = Convert.ToInt32(CardTypeID)
                        If IsValidProductID(ExtProductID, ProductTypeID, ErrorCode, ErrorMsg) Then
                            If IsValidExtLocationCode(ExtLocationCode, ErrorCode, ErrorMsg) Then
                                If ValidateAuxilaryDetails(IncludeAuxilaryDetailsCustProdLoc, IncludeAuxilaryOthers) Then
                                    Select Case EngineID
                                        Case ENGINE_ID.CM
                                            ErrorCode = ERROR_CODES.ERROR_AUXILARY_NOT_APPLIED_WITH_CM_ENGINE
                                            ErrorMsg = "Auxiliary is not applied with CM engine, call GetOffers method instead."

                                        Case ENGINE_ID.CPE, ENGINE_ID.UE
                                            MaxOffers = GetDefaultOfferLimit()
                                            If MaxOffers >= 0 Then
                                                Dim iEngineID As Integer = -1
                                                If Not String.IsNullOrEmpty(EngineID) AndAlso Not EngineID.Trim = String.Empty AndAlso IsNumeric(EngineID) Then iEngineID = Convert.ToInt32(EngineID)
                                                OfferXml = OfferListWithFiltersXMLAuxilary(TransactionDate, ExtCardID, iCardTypeID, ExtProductID, ProductTypeID, ExtLocationCode, iEngineID, ExtInterfaceID, MaxOffers, IncludeAuxilaryDetailsCustProdLoc, IncludeAuxilaryOthers)
                                                If OfferXml.Length = 0 Then
                                                    ErrorCode = ERROR_CODES.ERROR_NO_OFFER_FOUND_FOR_SEARCH_CRITERIA
                                                    ErrorMsg = "No Offers Found for The Search Criteria Entered."
                                                End If

                                            Else
                                                ErrorCode = ERROR_CODES.ERROR_DEFAULT_OFFER_LIMIT_NOT_SET_CORRECTLY
                                                ErrorMsg = "System Option Default Offer Limit Not Set Correctly."
                                            End If

                                        Case Else
                                            ErrorCode = ERROR_CODES.ERROR_OFFER_ENGINE_ID_INVALID
                                            ErrorMsg = "Invalid Offer EngineID: " & EngineID '& OfferEngineId

                                    End Select
                                Else
                                    ErrorCode = ERROR_CODES.ERROR_AUXILARY_NOT_APPLIED_WITH_CM_ENGINE
                                    ErrorMsg = "Fields IncludeAuxilaryDetailsCustProdLoc\IncludeAuxilaryOthers accept only boolean vlaue(True or False)."
                                End If
                            Else
                                ErrorCode = ErrorCode
                                ErrorMsg = ErrorMsg
                            End If
                        Else
                            ErrorCode = ErrorCode
                            ErrorMsg = ErrorMsg
                        End If
                    Else
                        ErrorCode = ErrorCode
                        ErrorMsg = ErrorMsg
                    End If
                Else
                    ErrorCode = ERROR_CODES.ERROR_INVALID_GUID_OR_EXTINTERFACEID
                    ErrorMsg = "Invalid GUID or ExternalInterfaceID"
                End If
            Else

                ErrorCode = ERROR_CODES.ERROR_INCORRECT_DATE_FORMAT
                ErrorMsg = "Either the format of StartDate or EndDate is incorrect."
            End If
        Catch ex As Exception
            ErrorCode = ERROR_CODES.ERROR_APPLICATION
            ErrorMsg = "Application error encountering: " & ex.ToString
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try

        If (ErrorCode = ERROR_CODES.ERROR_NONE) Then
            ResponseXml = OfferXml.Replace("encoding=""utf-16""", "encoding=""utf-8""")
            OfferXml = Nothing
        Else
            ' create error xml
            ErrorXml.Append("<?xml version=""1.0"" encoding=""UTF-8""?>" & ControlChars.CrLf)
            ErrorXml.Append("<UniversalOfferConnector>" & ControlChars.CrLf)
            ErrorXml.Append("  <Offer id=""" & """ operation=""" & RESPONSE_TYPES.GET_OFFERLIST_WITH_FILTER_AUXILARY.ToString & """ success=""false"" />" & ControlChars.CrLf)
            ErrorXml.Append("  <Error code=""" & ErrorCode.ToString & """ message=""" & ErrorMsg & """ />" & ControlChars.CrLf)
            ErrorXml.Append("</UniversalOfferConnector>")

            ResponseXml = ErrorXml.ToString
        End If

        ElapsedTime = Now.Subtract(RequestTime)
        AppendToLog(ExtInterfaceID, "GetOffers", EngineID.ToString, "GetOffers Count=" & MaxOffers & "; Elapsed Time= " & ElapsedTime.ToString, ErrorCode, ErrorMsg)

        Return ResponseXml
    End Function

    <WebMethod()>
    Public Function GetOfferListWithFilters(<System.Xml.Serialization.XmlElementAttribute(IsNullable:=True)> ByVal Guid As String, <System.Xml.Serialization.XmlElementAttribute(IsNullable:=True)> ByVal ExtInterfaceID As String, <System.Xml.Serialization.XmlElementAttribute(IsNullable:=True)> ByVal EngineID As String, <System.Xml.Serialization.XmlElementAttribute(IsNullable:=True)> ByVal TransactionDate As String, <System.Xml.Serialization.XmlElementAttribute(IsNullable:=True)> ByVal CardID As String, <System.Xml.Serialization.XmlElementAttribute(IsNullable:=True)> ByVal CardTypeID As String, <System.Xml.Serialization.XmlElementAttribute(IsNullable:=True)> ByVal ProductID As String, <System.Xml.Serialization.XmlElementAttribute(IsNullable:=True)> ByVal ProductTypeID As String, <System.Xml.Serialization.XmlElementAttribute(IsNullable:=True)> ByVal ExtLocationCode As String) As String
        Dim iExtInterfaceID As Integer = -1
        ' Dim iEngineID As Integer = -1
        ' Dim iCardTypeID As Integer = -1
        Dim iProductTypeID As Integer = -1
        Dim ExtCardID As String = ""
        Dim ExtProductID As String = ""

        TransactionDate = TransactionDate.Trim
        ExtLocationCode = ExtLocationCode.Trim
        ExtCardID = CardID.Trim
        ExtProductID = ProductID.Trim

        If Not String.IsNullOrEmpty(ExtInterfaceID) AndAlso Not ExtInterfaceID.Trim = String.Empty AndAlso IsNumeric(ExtInterfaceID) Then iExtInterfaceID = Convert.ToInt32(ExtInterfaceID)
        ' If Not String.IsNullOrEmpty(EngineID) AndAlso Not EngineID.Trim = String.Empty AndAlso IsNumeric(EngineID) Then iEngineID = Convert.ToInt32(EngineID)
        'If Not String.IsNullOrEmpty(CardTypeID) AndAlso Not CardTypeID.Trim = String.Empty AndAlso IsNumeric(CardTypeID) Then iCardTypeID = Convert.ToInt32(CardTypeID)
        If Not String.IsNullOrEmpty(ProductTypeID) AndAlso Not ProductTypeID.Trim = String.Empty AndAlso IsNumeric(ProductTypeID) Then iProductTypeID = Convert.ToInt32(ProductTypeID)

        Return _GetOfferListWithFilters(Guid, iExtInterfaceID, EngineID, TransactionDate, ExtCardID, CardTypeID, ProductID, iProductTypeID, ExtLocationCode)
    End Function

    Private Function _GetOfferListWithFilters(ByVal Guid As String, ByVal ExtInterfaceID As Integer, ByVal EngineID As String, ByVal TransactionDate As String, ByVal ExtCardID As String, ByVal CardTypeID As String, ByVal ExtProductID As String, ByVal ProductTypeID As Integer, ByVal ExtLocationCode As String) As String
        Dim ResponseXml As String = ""
        Dim OfferXml As String = ""
        Dim ErrorXml As New StringBuilder()
        Dim ErrorCode As ERROR_CODES = ERROR_CODES.ERROR_NONE
        Dim ErrorMsg As String = ""
        Dim OfferEngineId As Long = -1
        Dim MaxOffers As Integer = 1000
        Static RequestTime As DateTime = Now
        Dim ElapsedTime As TimeSpan

        Try
            Startup()
            ' If (TransactionDate = "") Or (IsDate(TransactionDate) And TransactionDate.Length >= 10) Then
            If (Not (TransactionDate = "") Or (IsDate(TransactionDate) And TransactionDate.Length >= 10)) And IsDate(TransactionDate) Then
                If ValidateExtInterfaceID(Guid, ExtInterfaceID) Then
                    If IsValidCustomerCard(ExtCardID, CardTypeID, ErrorCode, ErrorMsg) Then
                        Dim iCardTypeID As Integer = -1
                        If Not String.IsNullOrEmpty(CardTypeID) AndAlso Not CardTypeID.Trim = String.Empty AndAlso IsNumeric(CardTypeID) Then iCardTypeID = Convert.ToInt32(CardTypeID)

                        If IsValidProductID(ExtProductID, ProductTypeID, ErrorCode, ErrorMsg) Then
                            If IsValidExtLocationCode(ExtLocationCode, ErrorCode, ErrorMsg) Then

                                Select Case EngineID
                                    Case ENGINE_ID.CM, ENGINE_ID.CPE, ENGINE_ID.UE
                                        MaxOffers = GetDefaultOfferLimit()
                                        If MaxOffers >= 0 Then
                                            Dim iEngineID As Integer = -1
                                            If Not String.IsNullOrEmpty(EngineID) AndAlso Not EngineID.Trim = String.Empty AndAlso IsNumeric(EngineID) Then iEngineID = Convert.ToInt32(EngineID)

                                            OfferXml = OfferListWithFiltersXML(TransactionDate, ExtCardID, iCardTypeID, ExtProductID, ProductTypeID, ExtLocationCode, iEngineID, ExtInterfaceID, MaxOffers)
                                            If OfferXml.Length = 0 Then
                                                ErrorCode = ERROR_CODES.ERROR_NO_OFFER_FOUND_FOR_SEARCH_CRITERIA
                                                ErrorMsg = "No Offers Found for The Search Criteria Entered."
                                            End If

                                        Else
                                            ErrorCode = ERROR_CODES.ERROR_DEFAULT_OFFER_LIMIT_NOT_SET_CORRECTLY
                                            ErrorMsg = "System Option Default Offer Limit Not Set Correctly."
                                        End If

                                    Case Else
                                        ErrorCode = ERROR_CODES.ERROR_OFFER_ENGINE_ID_INVALID
                                        ErrorMsg = "Invalid Offer EngineID: " & EngineID '& OfferEngineId
                                End Select
                            Else
                                ErrorCode = ErrorCode
                                ErrorMsg = ErrorMsg
                            End If
                        Else
                            ErrorCode = ErrorCode
                            ErrorMsg = ErrorMsg
                        End If
                    Else
                        ErrorCode = ErrorCode
                        ErrorMsg = ErrorMsg
                    End If
                Else
                    ErrorCode = ERROR_CODES.ERROR_INVALID_GUID_OR_EXTINTERFACEID
                    ErrorMsg = "Invalid GUID or ExternalInterfaceID"
                End If
            Else
                ErrorCode = ERROR_CODES.ERROR_INCORRECT_DATE_FORMAT
                ErrorMsg = "Either the format of StartDate or EndDate is incorrect."
            End If
        Catch ex As Exception
            ErrorCode = ERROR_CODES.ERROR_APPLICATION
            ErrorMsg = "Application error encountering: " & ex.ToString
        Finally
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        End Try

        If (ErrorCode = ERROR_CODES.ERROR_NONE) Then
            ResponseXml = OfferXml.Replace("encoding=""utf-16""", "encoding=""utf-8""")
            OfferXml = Nothing
        Else
            ' create error xml
            ErrorXml.Append("<?xml version=""1.0"" encoding=""UTF-8""?>" & ControlChars.CrLf)
            ErrorXml.Append("<UniversalOfferConnector>" & ControlChars.CrLf)
            ErrorXml.Append("  <Offer id=""" & """ operation=""" & RESPONSE_TYPES.GET_OFFERLIST_WITH_FILTERS.ToString & """ success=""false"" />" & ControlChars.CrLf)
            ErrorXml.Append("  <Error code=""" & ErrorCode.ToString & """ message=""" & ErrorMsg & """ />" & ControlChars.CrLf)
            ErrorXml.Append("</UniversalOfferConnector>")

            ResponseXml = ErrorXml.ToString
        End If

        ElapsedTime = Now.Subtract(RequestTime)
        AppendToLog(ExtInterfaceID, "GetOffers", EngineID.ToString, "GetOffers Count=" & MaxOffers & "; Elapsed Time= " & ElapsedTime.ToString, ErrorCode, ErrorMsg)

        Return ResponseXml
    End Function

    Private Function GetEngineID(ByRef sEngineType As String) As Integer
        Dim EngineID As Integer = -1
        Dim dt As DataTable

        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
        If sEngineType = "" Then
            MyCommon.QueryStr = "select EngineID, Description from PromoEngines with (NoLock) where DefaultEngine=1 and Installed=1"
        Else
            MyCommon.QueryStr = "select EngineID, Description from PromoEngines with (NoLock) where Description = @Description and Installed=1"
            MyCommon.DBParameters.Add("@Description", SqlDbType.NVarChar).Value = sEngineType.ConvertBlankIfNothing()
        End If
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        If (dt.Rows.Count > 0) Then
            EngineID = MyCommon.NZ(dt.Rows(0).Item("EngineID"), -1)
            If sEngineType = "" Then
                sEngineType = MyCommon.NZ(dt.Rows(0).Item("Description"), "No engines installed")
            End If
        End If

        Return EngineID
    End Function

    Private Function GetExtInterfaceID(ByVal ExternalSourceID As String, ByRef AutoDeploy As Boolean, Optional ByRef DefaultAsLogixID As Boolean = False) As Integer
        Dim dt As DataTable

        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
        MyCommon.QueryStr = "select ExtInterfaceID, AutoDeploy, DefaultAsLogixID from ExtCRMInterfaces with (NoLock) where ExtCode = @ExtCode and Deleted=0"
        MyCommon.DBParameters.Add("@ExtCode", SqlDbType.NVarChar).Value = ExternalSourceID.ConvertBlankIfNothing
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        If (dt.Rows.Count > 0) Then
            ExtInterfaceID = MyCommon.NZ(dt.Rows(0).Item("ExtInterfaceID"), -1)
            AutoDeploy = MyCommon.NZ(dt.Rows(0).Item("AutoDeploy"), False)
            DefaultAsLogixID = MyCommon.NZ(dt.Rows(0).Item("DefaultAsLogixID"), False)
        End If

        Return ExtInterfaceID
    End Function

    Private Function CleanXmlString(ByVal sUnformattedXml As String) As String
        Dim CleanStr As String = ""

        If (sUnformattedXml IsNot Nothing) Then
            CleanStr = sUnformattedXml
            CleanStr = CleanStr.Replace("&", "&amp;")
            CleanStr = CleanStr.Replace("""", "&quot;")
            CleanStr = CleanStr.Replace("<", "&lt;")
        End If

        Return CleanStr
    End Function

    Private Sub AppendToLog(ByVal ExternalInterfaceID As String, ByVal FunctionName As String, ByVal EngineID As String, _
                          ByVal LogText As String, ByVal ErrorCode As ERROR_CODES, ByVal ErrorMsg As String)
        Dim LogString As String = ""
        Dim AndLogText As String = ""

        If LogText <> "" Then AndLogText = LogText & ";"

        Select Case ErrorCode
            Case ERROR_CODES.ERROR_NONE
                LogString = "Success=" & FunctionName & "; " & AndLogText & " ExternalInterfaceID=" & ExternalInterfaceID & "; EngineID= " & EngineID & "; Server= " & Environment.MachineName

                Copient.Logger.Write_Log(UOCLogFile, LogString, True)
            Case Else
                LogString = "Error=" & FunctionName & "; " & AndLogText & " ExternalInterfaceID=" & ExternalInterfaceID & "; EngineID= " & EngineID & "; Error_encountered=" & ErrorCode.ToString & " " & ErrorMsg & "; Server=" & Environment.MachineName '& "; IP = " & IP

                Copient.Logger.Write_Log(UOCErrorLogFile, LogString, True)
        End Select

    End Sub

    Private Function IsStringNumeric(ByVal str As String) As Boolean
        Dim NumericOnly As Boolean = False
        Dim i As Integer
        Dim strUbound As Integer

        If str IsNot Nothing Then
            strUbound = (str.Length - 1)
            For i = 0 To strUbound
                NumericOnly = Char.IsDigit(str.Chars(i))
                If Not NumericOnly Then Exit For
            Next
        End If

        Return NumericOnly
    End Function

    Private Function GetDefaultOfferLimit() As Integer
        Dim limit As Integer = -1
        Dim strLimit As String = ""

        Try
            strLimit = MyCommon.Fetch_SystemOption(149)
            If IsStringNumeric(strLimit) Then
                limit = Int32.Parse(strLimit)
            End If
        Catch ex As Exception
            Throw ex
        End Try

        Return limit
    End Function

    Private Function ValidateExtInterfaceID(ByVal Guid As String, ByVal ExtInterfaceID As Integer) As Boolean
        Dim dt As DataTable
        Dim ValidationOk As Boolean = False

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

            MyCommon.QueryStr = "select GUID from ConnectorGUIDs with (NoLock) where ConnectorID=53 and GUID = @GUID and ExtInterfaceID = @ExtInterfaceID"
            MyCommon.DBParameters.Add("@GUID", SqlDbType.NVarChar).Value = Guid.ConvertBlankIfNothing
            MyCommon.DBParameters.Add("@ExtInterfaceID", SqlDbType.Int).Value = ExtInterfaceID
            dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            If (dt.Rows.Count > 0) Then
                ValidationOk = True
            End If

        Catch ex As Exception
            Throw ex
        End Try

        Return ValidationOk
    End Function

    'added validation for Auxilary Details
    Private Function ValidateAuxilaryDetails(ByVal IncludeAuxilaryDetailsCustProdLoc As String, ByVal IncludeAuxilaryOthers As String) As Boolean
        Dim IsValid As Boolean = False
        IncludeAuxilaryDetailsCustProdLoc = IncludeAuxilaryDetailsCustProdLoc.ToLower()
        IncludeAuxilaryOthers = IncludeAuxilaryOthers.ToLower()
        If System.Text.RegularExpressions.Regex.IsMatch(IncludeAuxilaryDetailsCustProdLoc, "^(true|false)$") AndAlso System.Text.RegularExpressions.Regex.IsMatch(IncludeAuxilaryOthers, "^(true|false)$") Then
            IsValid = True
        Else
            IsValid = False
        End If
        Return IsValid
    End Function
    'added validation for customer card
    Private Function IsValidCustomerCard(ByVal CardID As String, ByVal CardTypeID As String, ByRef ErrorCode As ERROR_CODES, ByRef ErrorMsg As String) As Boolean
        Dim IsValid As Boolean = False
        Dim validationRespCode As CardValidationResponse
        ErrorCode = ERROR_CODES.ERROR_NONE
        ErrorMsg = ""

        Try
            If (MyCommon.AllowToProcessCustomerCard(CardID, CardTypeID, validationRespCode) = False) Then
                'If CardID Is Nothing OrElse CardID.Trim = "" Then
                '    'Bad customer ID
                '    ErrorCode = ERROR_CODES.ERROR_INVALID_CARDID
                '    ErrorMsg = "Failure:CardID is not provided"
                'ElseIf Not String.IsNullOrEmpty(CardID) AndAlso Not CardTypeID = -1 Then
                '    ErrorCode = ERROR_CODES.ERROR_NOT_FOUND_CUSTOMER
                '    ErrorMsg = "CardID: " & CardID & " with CardTypeID: " & CardTypeID & " not found."
                'Else
                '    'Bad customer type ID
                '    ErrorCode = ERROR_CODES.ERROR_INVALID_CARDTYPEID
                '    ErrorMsg = "Failure: CardTypeID is not provided"
                'End If

                If validationRespCode <> CardValidationResponse.SUCCESS Then
                    If validationRespCode = CardValidationResponse.CARDIDNOTNUMERIC OrElse validationRespCode = CardValidationResponse.INVALIDCARDFORMAT Then
                        ErrorCode = ERROR_CODES.ERROR_INVALID_CARDID
                    ElseIf validationRespCode = CardValidationResponse.CARDTYPENOTFOUND OrElse validationRespCode = CardValidationResponse.INVALIDCARDTYPEFORMAT Then
                        ErrorCode = ERROR_CODES.ERROR_INVALID_CARDTYPEID
                    ElseIf validationRespCode = CardValidationResponse.ERROR_APPLICATION Then
                        ErrorCode = ERROR_CODES.ERROR_APPLICATION
                    End If
                    ErrorMsg = MyCommon.CardValidationResponseMessage(CardID, CardTypeID, validationRespCode)

                End If

            Else
                IsValid = True
            End If

        Catch ex As Exception
            IsValid = False
        End Try
        Return IsValid
    End Function
    'added validation for Product ID
    Private Function IsValidProductID(ByVal ProductID As String, ByVal ProductTypeID As Integer, ByRef ErrorCode As ERROR_CODES, ByRef ErrorMsg As String) As Boolean
        Dim IsValid As Boolean = False
        Dim dt As DataTable
        ErrorCode = ERROR_CODES.ERROR_NONE
        ErrorMsg = ""

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            MyCommon.QueryStr = "select * from ProductTypes where ProductTypeID=@ProductTypeID"
            MyCommon.DBParameters.Add("@ProductTypeID", SqlDbType.NVarChar).Value = ProductTypeID
            dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            If System.Text.RegularExpressions.Regex.IsMatch(ProductID, "^[a-zA-Z0-9_-]+$") AndAlso Not (System.Text.RegularExpressions.Regex.IsMatch(ProductID, "^[a-zA-Z_]+$")) AndAlso dt.Rows.Count > 0 Then 'prodtypeid is only 5 types available i.e 1 to 5
                IsValid = True
            Else
                If dt.Rows.Count = 0 Then
                    ErrorCode = ERROR_CODES.ERROR_INVALID_PRODUCTTYPEID
                    ErrorMsg = "Failure: ProductTypeID not found."
                Else
                    ErrorCode = ERROR_CODES.ERROR_INVALID_PRODUCTID
                    ErrorMsg = "Failure:Product IDs must be alphanumeric."
                End If
            End If

        Catch ex As Exception
            IsValid = False
        End Try
        Return IsValid
    End Function

    'added checking for extlocationcode availability
    Private Function IsValidExtLocationCode(ByVal ExtLocationCode As String, ByRef ErrorCode As ERROR_CODES, ByRef ErrorMsg As String) As Boolean
        Dim dt As DataTable
        Dim IsValid As Boolean = False

        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            MyCommon.QueryStr = "select ExtLocationCode from Locations where ExtLocationCode=@ExtLocationCode"
            MyCommon.DBParameters.Add("@ExtLocationCode", SqlDbType.NVarChar).Value = ExtLocationCode.ConvertBlankIfNothing
            dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            If (dt.Rows.Count > 0) Then
                IsValid = True
            Else
                ErrorCode = ERROR_CODES.ERROR_INVALID_EXTLOCATIONCODE
                ErrorMsg = "Failure:Extlocationcode not found."
            End If
        Catch ex As Exception
            IsValid = False
        End Try
        Return IsValid
    End Function

    Private Function OffersXML(ByVal StartDate As String, ByVal EndDate As String, ByVal OfferEngineID As Integer, ByVal ExtInterfaceID As Integer, ByRef MaxOffers As Integer,Optional ByVal CustGrpID As Integer = 0) As String
        Dim OfferXml As String = ""
        Dim DateRange As String = ""
        Dim ExtIntIdNotZero As String = ""
        Dim TopLimit As String = ""
        Dim SbOffersXml As New StringBuilder()
        Dim dt As DataTable = Nothing
        Dim row As DataRow
        Dim MyExport As New Copient.ExportXml
        Dim MyExportCpe As New Copient.ExportXmlCPE
        Dim n, offersCount As Integer
        Dim TempFileName As String
		Dim StartDateParsed As Boolean
        Dim EndDateParsed As Boolean
        Dim dtStartDate, dtEndDate As Date
        TempFileName = System.AppDomain.CurrentDomain.BaseDirectory() & "Connectors\TempOfferFile" & Date.Now.ToString("yyyyMMdd") & OfferEngineID.ToString & ".txt"

        'Make sure StartDate starts at beginning of the date if no time entered
        If StartDate.Length = 10 Then
            StartDate = StartDate & " 00:00:00"
        End If

        'Make sure EndDate ends at the last second of the date if no time entered
        If EndDate.Length = 10 Then
            EndDate = EndDate & " 23:59:59"
        End If

         If StartDate <> "" And EndDate <> "" Then
            StartDateParsed = Date.TryParse(StartDate, dtStartDate)
            EndDateParsed = Date.TryParse(EndDate, dtEndDate)
        End If


        If MaxOffers > 0 Then
            TopLimit = " TOP (" & MaxOffers & ")"
        End If

        Try
            If (MyCommon.LRTadoConn.State = ConnectionState.Closed) Then MyCommon.Open_LogixRT()
            MyCommon.QueryStr = "dbo.pa_GetOffersByCondition"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@MaxOffers", SqlDbType.Int).Value = MaxOffers
            MyCommon.LRTsp.Parameters.Add("@StartDate", SqlDbType.DateTime).Value = IIf(Not StartDateParsed, DBNull.Value, dtStartDate)
            MyCommon.LRTsp.Parameters.Add("@EndDate", SqlDbType.DateTime).Value = IIf(Not EndDateParsed, DBNull.Value, dtEndDate)
            MyCommon.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = OfferEngineID
            MyCommon.LRTsp.Parameters.Add("@CRMEngineID", SqlDbType.Int).Value = ExtInterfaceID
            MyCommon.LRTsp.Parameters.Add("@ConditionTypeID", SqlDbType.Int).Value = 1
            MyCommon.LRTsp.Parameters.Add("@CustGrpID ", SqlDbType.Int).Value = custgrpID
            dt = MyCommon.LRTsp_select()
            MyCommon.Close_LRTsp()
            If (dt.Rows.Count > 0) Then
                Select Case OfferEngineID
                    Case ENGINE_ID.CM
                        For Each row In dt.Rows
                            If MyExport.GenOfferXml(MyCommon.NZ(row.Item("OfferID"), 0).ToString, TempFileName, True, 0) Then
                                OfferXml = File.ReadAllText(TempFileName)
                                If SbOffersXml.Length = 0 Then
                                    SbOffersXml.Append(OfferXml)
                                Else
                                    n = OfferXml.IndexOf("<PromoMaint")
                                    SbOffersXml.Append(OfferXml.Substring(n, (OfferXml.Length - n)))
                                End If
                                File.Delete(TempFileName)
                            End If
                            If OfferXml.Length > 0 Then offersCount = offersCount + 1
                        Next

                    Case ENGINE_ID.CPE
                        For Each row In dt.Rows
                            OfferXml = MyExportCpe.GetOfferXML(MyCommon.NZ(row.Item("OfferID"), 0))
                            If SbOffersXml.Length = 0 Then
                                SbOffersXml.Append(OfferXml)
                            Else
                                n = OfferXml.IndexOf("<Offer>")
                                SbOffersXml.Append(OfferXml.Substring(n, (OfferXml.Length - n)))
                            End If
                            If OfferXml.Length > 0 Then offersCount = offersCount + 1
                        Next

                    Case ENGINE_ID.UE
                        For Each row In dt.Rows
                            OfferXml = m_ExportXMLUE.GetOfferXML(MyCommon.NZ(row.Item("OfferID"), 0))
                            If SbOffersXml.Length = 0 Then
                                SbOffersXml.Append(OfferXml)
                            Else
                                n = OfferXml.IndexOf("<Offer>")
                                SbOffersXml.Append(OfferXml.Substring(n, (OfferXml.Length - n)))
                            End If
                            If OfferXml.Length > 0 Then offersCount = offersCount + 1
                        Next
                End Select
                If (MyCommon.LRTadoConn.State = Data.ConnectionState.Closed) Then MyCommon.Close_LogixRT()
            End If

        Catch ex As Exception
            Throw ex
        End Try

        MaxOffers = offersCount

        'for test number of offers can the system handle
        'Dim str = "dt.Rows.Count = " & dt.Rows.Count & "   offersCount = " & offersCount.ToString & " From: " & startTime.ToString & " To: " & Date.Now.ToString
        'SbOffersXml.Append(str)
        Return SbOffersXml.ToString
    End Function

    Private Function OffersXMLAuxilary(ByVal StartDate As String, ByVal EndDate As String, ByVal OfferEngineID As Integer, ByVal ExtInterfaceID As Integer, ByRef MaxOffers As Integer, ByVal IncludeAuxilaryDetailsCustProdLoc As Boolean, ByVal IncludeAuxilaryOthers As Boolean) As String
        Dim OfferXml As String = ""
        Dim DateRange As String = ""
        Dim ExtIntIdNotZero As String = ""
        Dim TopLimit As String = ""
        Dim SbOffersXml As New StringBuilder()
        Dim dt As DataTable = Nothing
        Dim row As DataRow
        Dim MyExportCpe As New Copient.ExportXmlCPE
        Dim n, offersCount As Integer
        Dim TempFileName As String
        TempFileName = System.AppDomain.CurrentDomain.BaseDirectory() & "Connectors\TempOfferFile" & Date.Now.ToString("yyyyMMdd") & OfferEngineID.ToString & ".txt"

        'Make sure StartDate starts at beginning of the date if no time entered
        If StartDate.Length = 10 Then
            StartDate = StartDate & " 00:00:00"
        End If

        'Make sure EndDate ends at the last second of the date if no time entered
        If EndDate.Length = 10 Then
            EndDate = EndDate & " 23:59:59"
        End If

        If StartDate <> "" And EndDate <> "" Then
            DateRange = "ci.LastUpdate between @StartDate and @EndDate and "
            MyCommon.DBParameters.Add("@StartDate", SqlDbType.DateTime).Value = StartDate
            MyCommon.DBParameters.Add("@EndDate", SqlDbType.DateTime).Value = EndDate
        End If

        If ExtInterfaceID > 0 Then
            ExtIntIdNotZero = " and ci.InboundCRMEngineID = @InboundCRMEngineID "
            MyCommon.DBParameters.Add("@InboundCRMEngineID", SqlDbType.Int).Value = ExtInterfaceID
        End If

        If MaxOffers > 0 Then
            TopLimit = " TOP (" & MaxOffers & ")"
        End If

        Try
            If (MyCommon.LRTadoConn.State = Data.ConnectionState.Closed) Then MyCommon.Open_LogixRT()
            Select Case OfferEngineID

                Case ENGINE_ID.CPE, ENGINE_ID.UE
                    MyCommon.QueryStr = "select" & TopLimit & " ci.IncentiveID as OfferID from CPE_Incentives ci with (NoLock) " &
                                        "where " &
                                        DateRange & "EngineID= @EngineID" & ExtIntIdNotZero & " And Deleted=0 And IsTemplate=0 and ci.StatusFlag <> 11 and ci.StatusFlag <> 12"
                    MyCommon.DBParameters.Add("@EngineID", SqlDbType.Int).Value = OfferEngineID

            End Select

            dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            If (dt.Rows.Count > 0) Then
                Select Case OfferEngineID

                    Case ENGINE_ID.CPE
                        For Each row In dt.Rows
                            OfferXml = MyExportCpe.GetOfferXMLAuxilary(MyCommon.NZ(row.Item("OfferID"), 0), IncludeAuxilaryDetailsCustProdLoc, IncludeAuxilaryOthers)
                            If SbOffersXml.Length = 0 Then
                                SbOffersXml.Append(OfferXml)
                            Else
                                n = OfferXml.IndexOf("<Offer>")
                                SbOffersXml.Append(OfferXml.Substring(n, (OfferXml.Length - n)))
                            End If
                            If OfferXml.Length > 0 Then offersCount = offersCount + 1
                        Next

                    Case ENGINE_ID.UE
                        For Each row In dt.Rows
                            OfferXml = m_ExportXMLUE.GetOfferXMLAuxilary(MyCommon.NZ(row.Item("OfferID"), 0), IncludeAuxilaryDetailsCustProdLoc, IncludeAuxilaryOthers)
                            If SbOffersXml.Length = 0 Then
                                SbOffersXml.Append(OfferXml)
                            Else
                                n = OfferXml.IndexOf("<Offer>")
                                SbOffersXml.Append(OfferXml.Substring(n, (OfferXml.Length - n)))
                            End If
                            If OfferXml.Length > 0 Then offersCount = offersCount + 1
                        Next
                    Case Else

                End Select
                If (MyCommon.LRTadoConn.State = Data.ConnectionState.Closed) Then MyCommon.Close_LogixRT()
            End If

        Catch ex As Exception
            Throw ex
        End Try

        MaxOffers = offersCount

        'for test number of offers can the system handle
        'Dim str = "dt.Rows.Count = " & dt.Rows.Count & "   offersCount = " & offersCount.ToString & " From: " & startTime.ToString & " To: " & Date.Now.ToString
        'SbOffersXml.Append(str)
        Return SbOffersXml.ToString
    End Function

    Private Function OfferListWithFiltersXMLAuxilary(ByVal TransactionDate As String, ByVal ExtCardID As String, ByVal CardTypeID As Integer, ByVal ExtProductID As String, ByVal ProductTypeID As Integer, ByVal ExtLocationCode As String, ByVal OfferEngineID As Integer, ByVal ExtInterfaceID As Integer, ByRef MaxOffers As Integer, ByVal IncludeAuxilaryDetailsCustProdLoc As Boolean, ByVal IncludeAuxilaryOthers As Boolean) As String
        Dim OfferXml As String = ""
        Dim DateRange As String = ""
        Dim ExtIntIdNotZero As String = ""
        Dim TopLimit As String = ""
        Dim SbOffersXml As New StringBuilder()
        Dim dt As DataTable = Nothing
        Dim row As DataRow
        Dim MyExportCpe As New Copient.ExportXmlCPE
        Dim n, offersCount As Integer
        Dim JoinOfferLocations As String = ""
        Dim LocationGroupIDs As String = ""
        Dim JoinCustomerGroup1 As String = ""
        Dim CustomerGroupIDs1 As String = ""
        Dim JoinCustomerGroup2 As String = ""
        Dim CustomerGroupIDs2 As String = ""
        Dim JoinProductGroup1 As String = ""
        Dim ProductGroupIDs1 As String = ""
        Dim JoinProductGroup2 As String = ""
        Dim ProductGroupIDs2 As String = ""
        Dim UnionCustProd As String = ""

        Dim DBParam_CustomerGroupIDs As String = ""
        Dim DBParam_ProductGroupIDs As String = ""
        Dim DBParam_LocationGroupIDs As String = ""
        Dim DBParam_InboundCRMEngineID As Integer = -1
        Dim DBParam_TransactionDate As String = ""

        Dim TempFileName As String
        TempFileName = System.AppDomain.CurrentDomain.BaseDirectory() & "Connectors\TempOfferFile" & Date.Now.ToString("yyyyMMdd") & OfferEngineID.ToString & ".txt"
        Startup()
        If ExtInterfaceID > 0 Then
            If (OfferEngineID = ENGINE_ID.CPE Or OfferEngineID = ENGINE_ID.UE) Then
                ExtIntIdNotZero = " And INC.InboundCRMEngineID = @InboundCRMEngineID "
                DBParam_InboundCRMEngineID = ExtInterfaceID
            End If
        End If

        If MaxOffers > 0 Then
            TopLimit = " TOP (" & MaxOffers & ")"
        End If

        If TransactionDate <> "" Then
            If (OfferEngineID = ENGINE_ID.CPE Or OfferEngineID = ENGINE_ID.UE) Then
                DateRange = " @TransactionDate between INC.StartDate And INC.EndDate And "
                DBParam_TransactionDate = TransactionDate
            End If
        End If

        If ExtCardID <> "" And CardTypeID <> -1 Then
            If (OfferEngineID = ENGINE_ID.CPE Or OfferEngineID = ENGINE_ID.UE) Then
                JoinCustomerGroup1 = " LEFT JOIN " & If(TableType = TableTypeEnum.MAINTAINABLE, "CPE_IncentiveCustomerGroups", "CPE_ST_IncentiveCustomerGroups") &
                                     " as ICG with (NoLock) ON RO.RewardOptionID = ICG.RewardOptionID "
                CustomerGroupIDs1 = " And (ICG.Deleted = 0 And ICG.CustomerGroupID In (select items from dbo.Split(@CustomerGroupIDs, ',')))"

                JoinCustomerGroup2 = " INNER JOIN " & If(TableType = TableTypeEnum.MAINTAINABLE, "CPE_Deliverables", "CPE_ST_Deliverables") & _
                                     " as DELV with (NoLock) ON RO.RewardOptionID = DELV.RewardOptionID "
                CustomerGroupIDs2 = " and (DELV.Deleted = 0 and DELV.DeliverableTypeId IN (5,6) " & _
                                    " and DELV.OutputID In (select items from dbo.Split(@CustomerGroupIDs, ',')))"
                DBParam_CustomerGroupIDs = GetCustomerGroupIDs(ExtCardID, CardTypeID)
            End If
        End If

        If ExtProductID <> "" And ProductTypeID <> -1 Then
            If (OfferEngineID = ENGINE_ID.CPE Or OfferEngineID = ENGINE_ID.UE) Then
                JoinProductGroup1 = " RIGHT OUTER JOIN " & If(TableType = TableTypeEnum.MAINTAINABLE, "CPE_IncentiveProductGroups", "CPE_ST_IncentiveProductGroups") & _
                                    " as IPG with (NoLock) ON RO.RewardOptionID = IPG.RewardOptionID "
                ProductGroupIDs1 = " and IPG.Deleted = 0 and (IPG.ProductGroupID In (select items from dbo.Split(@ProductGroupIDs, ',')))"

                If JoinCustomerGroup2 = "" Then
                    JoinProductGroup2 = " INNER JOIN " & If(TableType = TableTypeEnum.MAINTAINABLE, "CPE_Deliverables", "CPE_ST_Deliverables") & _
                                        " as DELV with (NoLock) ON RO.RewardOptionID = DELV.RewardOptionID " & _
                                        " INNER JOIN " & If(TableType = TableTypeEnum.MAINTAINABLE, "CPE_Discounts ", "CPE_ST_Discounts ") & _
                                        " as DISC with (NoLock) on DISC.DiscountID = DELV.OutputID "

                    ProductGroupIDs2 = " and DELV.Deleted = 0 and DISC.Deleted = 0 " & _
                                       " and (DISC.DiscountedProductGroupID In (select items from dbo.Split(@ProductGroupIDs, ','))" & _
                                       " OR DISC.ExcludedProductGroupID In (select items from dbo.Split(@ProductGroupIDs, ',')))"
                Else
                    JoinProductGroup2 = " INNER JOIN " & If(TableType = TableTypeEnum.MAINTAINABLE, "CPE_Discounts ", "CPE_ST_Discounts ") & _
                                        " as DISC with (NoLock) on DISC.DiscountID = DELV.OutputID "

                    ProductGroupIDs2 = " and DISC.Deleted = 0 " & _
                                       " and (DISC.DiscountedProductGroupID In (select items from dbo.Split(@ProductGroupIDs, ','))" & _
                                       " OR DISC.ExcludedProductGroupID In (select items from dbo.Split(@ProductGroupIDs, ',')))"
                End If
                DBParam_ProductGroupIDs = GetProductGroupIDs(ExtProductID, ProductTypeID)
            End If
        End If

        If ExtLocationCode <> "" Then

            If (OfferEngineID = ENGINE_ID.CPE Or OfferEngineID = ENGINE_ID.UE) Then
                JoinOfferLocations = " INNER JOIN OfferLocations ON INC.IncentiveID = OfferLocations.OfferID "
            End If

            LocationGroupIDs = " and OfferLocations.LocationGroupID In (select items from dbo.Split(@LocationGroupIDs, ',')) and OfferLocations.Deleted = 0"
            DBParam_LocationGroupIDs = GetLocationGroupIDs(ExtLocationCode)

        End If

        Try
            If (MyCommon.LRTadoConn.State = ConnectionState.Closed) Then MyCommon.Open_LogixRT()

            Select Case OfferEngineID

                Case ENGINE_ID.CPE, ENGINE_ID.UE
                    MyCommon.QueryStr = "select Distinct " & TopLimit & " INC.IncentiveID as OfferID from " & If(TableType = TableTypeEnum.MAINTAINABLE, "CPE_Incentives", "CPE_ST_Incentives") &
                                        " as INC with (NoLock) inner join " & If(TableType = TableTypeEnum.MAINTAINABLE, "CPE_RewardOptions", "CPE_ST_RewardOptions") &
                                        " as RO with (NoLock) ON INC.IncentiveID = RO.IncentiveID " & JoinOfferLocations &
                                        JoinCustomerGroup1 & JoinProductGroup1 &
                                        " where " & DateRange & "INC.EngineID = @EngineID " & ExtIntIdNotZero & " and INC.Deleted=0 and INC.IsTemplate=0 and RO.Deleted=0 and INC.StatusFlag <> 11 and  INC.StatusFlag <> 12 " &
                                        LocationGroupIDs & CustomerGroupIDs1 & ProductGroupIDs1

                    If (JoinCustomerGroup1 <> "" Or JoinProductGroup1 <> "") Then
                        UnionCustProd = " UNION " &
                                        "select Distinct " & TopLimit & " INC.IncentiveID as OfferID from " & If(TableType = TableTypeEnum.MAINTAINABLE, "CPE_Incentives", "CPE_ST_Incentives") &
                                        " as INC with (NoLock) inner join " & If(TableType = TableTypeEnum.MAINTAINABLE, "CPE_RewardOptions", "CPE_ST_RewardOptions") &
                                        " as RO with (NoLock) ON INC.IncentiveID = RO.IncentiveID " & JoinOfferLocations &
                                        JoinCustomerGroup2 & JoinProductGroup2 &
                                        " where " & DateRange & "INC.EngineID = @EngineID" & ExtIntIdNotZero & " and INC.Deleted=0 and INC.IsTemplate=0 and RO.Deleted=0 and INC.StatusFlag <> 11 and  INC.StatusFlag <> 12 " &
                                        LocationGroupIDs & CustomerGroupIDs2 & ProductGroupIDs2

                        MyCommon.QueryStr = MyCommon.QueryStr & UnionCustProd
                    End If

            End Select

            MyCommon.DBParameters.Add("@CustomerGroupIDs", SqlDbType.NVarChar, Int32.MaxValue).Value = DBParam_CustomerGroupIDs
            MyCommon.DBParameters.Add("@ProductGroupIDs", SqlDbType.NVarChar, Int32.MaxValue).Value = DBParam_ProductGroupIDs
            MyCommon.DBParameters.Add("@LocationGroupIDs", SqlDbType.NVarChar, Int32.MaxValue).Value = DBParam_LocationGroupIDs
            MyCommon.DBParameters.Add("@InboundCRMEngineID", SqlDbType.Int).Value = DBParam_InboundCRMEngineID
            MyCommon.DBParameters.Add("@TransactionDate", SqlDbType.NVarChar, Int32.MaxValue).Value = DBParam_TransactionDate
            MyCommon.DBParameters.Add("@EngineID", SqlDbType.Int).Value = OfferEngineID
            dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            If (dt.Rows.Count > 0) Then
                Select Case OfferEngineID

                    Case ENGINE_ID.CPE
                        For Each row In dt.Rows
                            OfferXml = MyExportCpe.GetOfferXMLAuxilary(MyCommon.NZ(row.Item("OfferID"), 0), IncludeAuxilaryDetailsCustProdLoc, IncludeAuxilaryOthers)
                            If SbOffersXml.Length = 0 Then
                                SbOffersXml.Append(OfferXml)
                            Else
                                n = OfferXml.IndexOf("<Offer>")
                                SbOffersXml.Append(OfferXml.Substring(n, (OfferXml.Length - n)))
                            End If
                            If OfferXml.Length > 0 Then offersCount = offersCount + 1
                        Next

                    Case ENGINE_ID.UE
                        For Each row In dt.Rows
                            OfferXml = m_ExportXMLUE.GetOfferXMLAuxilary(MyCommon.NZ(row.Item("OfferID"), 0), IncludeAuxilaryDetailsCustProdLoc, IncludeAuxilaryOthers)
                            If SbOffersXml.Length = 0 Then
                                SbOffersXml.Append(OfferXml)
                            Else
                                n = OfferXml.IndexOf("<Offer>")
                                SbOffersXml.Append(OfferXml.Substring(n, (OfferXml.Length - n)))
                            End If
                            If OfferXml.Length > 0 Then offersCount = offersCount + 1
                        Next
                    Case Else

                End Select
                If (MyCommon.LRTadoConn.State = Data.ConnectionState.Closed) Then MyCommon.Close_LogixRT()
                If (MyCommon.LXSadoConn.State = ConnectionState.Closed) Then MyCommon.Open_LogixXS()
            End If

        Catch ex As Exception
            Throw ex
        End Try

        MaxOffers = offersCount

        'for test number of offers can the system handle
        'Dim str = "dt.Rows.Count = " & dt.Rows.Count & "   offersCount = " & offersCount.ToString & " From: " & startTime.ToString & " To: " & Date.Now.ToString
        'SbOffersXml.Append(str)
        Return SbOffersXml.ToString
    End Function

    Private Function OfferListWithFiltersXML(ByVal TransactionDate As String, ByVal ExtCardID As String, ByVal CardTypeID As Integer, ByVal ExtProductID As String, ByVal ProductTypeID As Integer, ByVal ExtLocationCode As String, ByVal OfferEngineID As Integer, ByVal ExtInterfaceID As Integer, ByRef MaxOffers As Integer) As String
        Dim OfferXml As String = ""
        Dim DateRange As String = ""
        Dim ExtIntIdNotZero As String = ""
        Dim TopLimit As String = ""
        Dim SbOffersXml As New StringBuilder()
        Dim dt As DataTable = Nothing
        Dim row As DataRow
        Dim MyExport As New Copient.ExportXml
        Dim MyExportCpe As New Copient.ExportXmlCPE
        Dim n, offersCount As Integer

        Dim JoinOfferLocations As String = ""
        Dim LocationGroupIDs As String = ""
        Dim JoinCustomerGroup1 As String = ""
        Dim CustomerGroupIDs1 As String = ""
        Dim JoinCustomerGroup2 As String = ""
        Dim CustomerGroupIDs2 As String = ""
        Dim JoinProductGroup1 As String = ""
        Dim ProductGroupIDs1 As String = ""
        Dim JoinProductGroup2 As String = ""
        Dim ProductGroupIDs2 As String = ""
        Dim UnionCustProd As String = ""

        Dim DBParam_CustomerGroupIDs As String = ""
        Dim DBParam_ProductGroupIDs As String = ""
        Dim DBParam_LocationGroupIDs As String = ""
        Dim DBParam_InboundCRMEngineID As Integer = -1
        Dim DBParam_TransactionDate As String = ""

        Dim TempFileName As String
        TempFileName = System.AppDomain.CurrentDomain.BaseDirectory() & "Connectors\TempOfferFile" & Date.Now.ToString("yyyyMMdd") & OfferEngineID.ToString & ".txt"

        If ExtInterfaceID > 0 Then
            If OfferEngineID = ENGINE_ID.CM Then
                ExtIntIdNotZero = " and O.InboundCRMEngineID = @InboundCRMEngineID "
            Else
                ExtIntIdNotZero = " and INC.InboundCRMEngineID = @InboundCRMEngineID "
            End If
            DBParam_InboundCRMEngineID = ExtInterfaceID
        End If

        If MaxOffers > 0 Then
            TopLimit = " TOP (" & MaxOffers & ")"
        End If

        If TransactionDate <> "" Then
            If OfferEngineID = ENGINE_ID.CM Then
                DateRange = " @TransactionDate between O.ProdStartDate and O.ProdEndDate and "
            Else
                DateRange = " @TransactionDate between INC.StartDate and INC.EndDate and "
            End If
            DBParam_TransactionDate = TransactionDate
        End If

        If ExtCardID <> "" And CardTypeID <> -1 Then
            If OfferEngineID = ENGINE_ID.CM Then
                JoinCustomerGroup1 = " inner join " & If(TableType = TableTypeEnum.MAINTAINABLE, "OfferConditions", "CM_ST_OfferConditions") & _
                                     " as OC with (NoLock) on O.OfferID = OC.offerID and OC.Deleted = 0 "

                CustomerGroupIDs1 = " and (OC.ConditionTypeID=1 and " & _
                                    "( OC.LinkID In (select items from dbo.Split(@CustomerGroupIDs, ',')) OR " & _
                                    "OC.ExcludedID In (select items from dbo.Split(@CustomerGroupIDs, ',')))) "

                JoinCustomerGroup2 = " inner join " & If(TableType = TableTypeEnum.MAINTAINABLE, "OfferRewards", "CM_ST_OfferRewards") & _
                                     " as ORW with (NoLock) ON O.OfferID = ORW.OfferID and ORW.Deleted=0 " & _
                                     " LEFT JOIN " & If(TableType = TableTypeEnum.MAINTAINABLE, "RewardCustomerGroupTiers", "CM_ST_RewardCustomerGroupTiers") & _
                                     " as RCG with (NoLock) ON ORW.RewardID = RCG.RewardID "

                CustomerGroupIDs2 = " and ORW.RewardTypeID In (5, 6) " & _
                                    " and RCG.CustomerGroupID In (select items from dbo.Split(@CustomerGroupIDs, ',')) "

            ElseIf (OfferEngineID = ENGINE_ID.CPE Or OfferEngineID = ENGINE_ID.UE) Then
                JoinCustomerGroup1 = " LEFT JOIN " & If(TableType = TableTypeEnum.MAINTAINABLE, "CPE_IncentiveCustomerGroups", "CPE_ST_IncentiveCustomerGroups") & _
                                     " as ICG with (NoLock) ON RO.RewardOptionID = ICG.RewardOptionID "
                CustomerGroupIDs1 = " and (ICG.Deleted = 0 and ICG.CustomerGroupID In (select items from dbo.Split(@CustomerGroupIDs, ',')))"

                JoinCustomerGroup2 = " INNER JOIN " & If(TableType = TableTypeEnum.MAINTAINABLE, "CPE_Deliverables", "CPE_ST_Deliverables") & _
                                     " as DELV with (NoLock) ON RO.RewardOptionID = DELV.RewardOptionID "
                CustomerGroupIDs2 = " and (DELV.Deleted = 0 and DELV.DeliverableTypeId IN (5,6) " & _
                                    " and DELV.OutputID In (select items from dbo.Split(@CustomerGroupIDs, ',')))"

            End If
            DBParam_CustomerGroupIDs = GetCustomerGroupIDs(ExtCardID, CardTypeID)
        End If

        If ExtProductID <> "" And ProductTypeID <> -1 Then
            If OfferEngineID = ENGINE_ID.CM Then
                If JoinCustomerGroup1 = "" Then
                    JoinProductGroup1 = " inner join " & If(TableType = TableTypeEnum.MAINTAINABLE, "OfferConditions", "CM_ST_OfferConditions") & _
                                        " as OC with (NoLock) on O.OfferID = OC.offerID and OC.Deleted = 0 "

                    ProductGroupIDs1 = " and (OC.ConditionTypeID=2 and (OC.LinkID In (select items from dbo.Split(@ProductGroupIDs, ','))" & _
                                       " OR OC.ExcludedID In (select items from dbo.Split(@ProductGroupIDs, ','))))"
                Else
                    ProductGroupIDs1 = " or (OC.ConditionTypeID=2 and (OC.LinkID In (select items from dbo.Split(@ProductGroupIDs, ','))" & _
                                       " OR OC.ExcludedID In (select items from dbo.Split(@ProductGroupIDs, ',')))) group by O.OfferID "
                End If

                If JoinCustomerGroup2 = "" Then
                    JoinProductGroup2 = " INNER JOIN " & If(TableType = TableTypeEnum.MAINTAINABLE, "OfferRewards", "CM_ST_OfferRewards") & _
                                        " as ORW with (NoLock) ON O.OfferID = ORW.OfferID and ORW.Deleted = 0 "
                End If

                ProductGroupIDs2 = " and (ORW.ProductGroupID In (select items from dbo.Split(@ProductGroupIDs, ','))" & _
                                   " OR ORW.ExcludedProdGroupID In (select items from dbo.Split(@ProductGroupIDs, ',')))"

            ElseIf (OfferEngineID = ENGINE_ID.CPE Or OfferEngineID = ENGINE_ID.UE) Then
                JoinProductGroup1 = " RIGHT OUTER JOIN " & If(TableType = TableTypeEnum.MAINTAINABLE, "CPE_IncentiveProductGroups", "CPE_ST_IncentiveProductGroups") & _
                                    " as IPG with (NoLock) ON RO.RewardOptionID = IPG.RewardOptionID "
                ProductGroupIDs1 = " and IPG.Deleted = 0 and (IPG.ProductGroupID In (select items from dbo.Split(@ProductGroupIDs, ',')))"

                If JoinCustomerGroup2 = "" Then
                    JoinProductGroup2 = " INNER JOIN " & If(TableType = TableTypeEnum.MAINTAINABLE, "CPE_Deliverables", "CPE_ST_Deliverables") & _
                                        " as DELV with (NoLock) ON RO.RewardOptionID = DELV.RewardOptionID " & _
                                        " INNER JOIN " & If(TableType = TableTypeEnum.MAINTAINABLE, "CPE_Discounts ", "CPE_ST_Discounts ") & _
                                        " as DISC with (NoLock) on DISC.DiscountID = DELV.OutputID "

                    ProductGroupIDs2 = " and DELV.Deleted = 0 and DISC.Deleted = 0 " & _
                                       " and (DISC.DiscountedProductGroupID In (select items from dbo.Split(@ProductGroupIDs, ','))" & _
                                       " OR DISC.ExcludedProductGroupID In (select items from dbo.Split(@ProductGroupIDs, ',')))"
                Else
                    JoinProductGroup2 = " INNER JOIN " & If(TableType = TableTypeEnum.MAINTAINABLE, "CPE_Discounts ", "CPE_ST_Discounts ") & _
                                        " as DISC with (NoLock) on DISC.DiscountID = DELV.OutputID "

                    ProductGroupIDs2 = " and DISC.Deleted = 0 " & _
                                       " and (DISC.DiscountedProductGroupID In (select items from dbo.Split(@ProductGroupIDs, ','))" & _
                                       " OR DISC.ExcludedProductGroupID In (select items from dbo.Split(@ProductGroupIDs, ',')))"
                End If

            End If
            DBParam_ProductGroupIDs = GetProductGroupIDs(ExtProductID, ProductTypeID)
        End If

        If ExtLocationCode <> "" Then

            If OfferEngineID = ENGINE_ID.CM Then
                JoinOfferLocations = " INNER JOIN OfferLocations ON O.OfferID = OfferLocations.OfferID "
            Else
                JoinOfferLocations = " INNER JOIN OfferLocations ON INC.IncentiveID = OfferLocations.OfferID "
            End If

            LocationGroupIDs = " and OfferLocations.LocationGroupID In (select items from dbo.Split(@LocationGroupIDs, ',')) and OfferLocations.Deleted = 0"
            DBParam_LocationGroupIDs = GetLocationGroupIDs(ExtLocationCode)

        End If

        If (MyCommon.LRTadoConn.State = ConnectionState.Closed) Then MyCommon.Open_LogixRT()

        Select Case OfferEngineID
            Case ENGINE_ID.CM
                If (CustomerGroupIDs1 <> "" And ProductGroupIDs1 <> "") Then
                    MyCommon.QueryStr = "select  Q1.OfferID from " & _
                                        "(select " & TopLimit & " O.OfferID, count(O.OfferID) as Occurence from " & If(TableType = TableTypeEnum.MAINTAINABLE, "Offers", "CM_ST_Offers") & _
                                        " as O with (NoLock) " & JoinOfferLocations & _
                                        JoinCustomerGroup1 & JoinProductGroup1 & _
                                        " where " & DateRange & "O.EngineID= @EngineID" & ExtIntIdNotZero & " and O.Deleted=0 and O.IsTemplate=0 " & _
                                        LocationGroupIDs & CustomerGroupIDs1 & ProductGroupIDs1 & _
                                        ") Q1 where Q1.Occurence > 1"
                Else
                    MyCommon.QueryStr = "select Distinct " & TopLimit & " O.OfferID from " & If(TableType = TableTypeEnum.MAINTAINABLE, "Offers", "CM_ST_Offers") & _
                                        " as O with (NoLock) " & JoinOfferLocations & _
                                        JoinCustomerGroup1 & JoinProductGroup1 & _
                                        " where " & DateRange & "O.EngineID= @EngineID" & ExtIntIdNotZero & " and O.Deleted=0 and O.IsTemplate=0 " & _
                                        LocationGroupIDs & CustomerGroupIDs1 & ProductGroupIDs1
                End If

                If (JoinCustomerGroup1 <> "" Or JoinProductGroup1 <> "") Then
                    UnionCustProd = " UNION " & _
                                    "select Distinct " & TopLimit & " O.OfferID from " & If(TableType = TableTypeEnum.MAINTAINABLE, "Offers", "CM_ST_Offers") & _
                                    " as O with (NoLock) " & JoinOfferLocations & _
                                    JoinCustomerGroup2 & JoinProductGroup2 & _
                                    " where " & DateRange & "O.EngineID = @EngineID" & ExtIntIdNotZero & " and O.Deleted=0 and O.IsTemplate=0 " & _
                                    LocationGroupIDs & CustomerGroupIDs2 & ProductGroupIDs2
                    MyCommon.QueryStr = MyCommon.QueryStr & UnionCustProd
                End If

            Case ENGINE_ID.CPE, ENGINE_ID.UE
                MyCommon.QueryStr = "select Distinct " & TopLimit & " INC.IncentiveID as OfferID from " & If(TableType = TableTypeEnum.MAINTAINABLE, "CPE_Incentives", "CPE_ST_Incentives") &
                                    " as INC with (NoLock) inner join " & If(TableType = TableTypeEnum.MAINTAINABLE, "CPE_RewardOptions", "CPE_ST_RewardOptions") &
                                    " as RO with (NoLock) ON INC.IncentiveID = RO.IncentiveID " & JoinOfferLocations &
                                    JoinCustomerGroup1 & JoinProductGroup1 &
                                    " where " & DateRange & "INC.EngineID = @EngineID" & ExtIntIdNotZero & " and INC.Deleted=0 and INC.IsTemplate=0 and RO.Deleted=0 and INC.StatusFlag <> 11 and  INC.StatusFlag <> 12" &
                                    LocationGroupIDs & CustomerGroupIDs1 & ProductGroupIDs1

                If (JoinCustomerGroup1 <> "" Or JoinProductGroup1 <> "") Then
                    UnionCustProd = " UNION " &
                                   "select Distinct " & TopLimit & " INC.IncentiveID as OfferID from " & If(TableType = TableTypeEnum.MAINTAINABLE, "CPE_Incentives", "CPE_ST_Incentives") &
                                   " as INC with (NoLock) inner join " & If(TableType = TableTypeEnum.MAINTAINABLE, "CPE_RewardOptions", "CPE_ST_RewardOptions") &
                                   " as RO with (NoLock) ON INC.IncentiveID = RO.IncentiveID " & JoinOfferLocations &
                                   JoinCustomerGroup2 & JoinProductGroup2 &
                                   " where " & DateRange & "INC.EngineID = @EngineID" & ExtIntIdNotZero & " and INC.Deleted=0 and INC.IsTemplate=0 and RO.Deleted=0 and INC.StatusFlag <> 11 and  INC.StatusFlag <> 12" &
                                   LocationGroupIDs & CustomerGroupIDs2 & ProductGroupIDs2
                    MyCommon.QueryStr = MyCommon.QueryStr & UnionCustProd
                End If

        End Select

        MyCommon.DBParameters.Add("@CustomerGroupIDs", SqlDbType.NVarChar, Int32.MaxValue).Value = DBParam_CustomerGroupIDs
        MyCommon.DBParameters.Add("@ProductGroupIDs", SqlDbType.NVarChar, Int32.MaxValue).Value = DBParam_ProductGroupIDs
        MyCommon.DBParameters.Add("@LocationGroupIDs", SqlDbType.NVarChar, Int32.MaxValue).Value = DBParam_LocationGroupIDs
        MyCommon.DBParameters.Add("@InboundCRMEngineID", SqlDbType.Int).Value = DBParam_InboundCRMEngineID
        MyCommon.DBParameters.Add("@TransactionDate", SqlDbType.NVarChar, Int32.MaxValue).Value = DBParam_TransactionDate
        MyCommon.DBParameters.Add("@EngineID", SqlDbType.Int).Value = OfferEngineID
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        If (dt.Rows.Count > 0) Then
            Select Case OfferEngineID
                Case ENGINE_ID.CM
                    For Each row In dt.Rows
                        If MyExport.GenOfferXml(MyCommon.NZ(row.Item("OfferID"), 0).ToString, TempFileName, True, 0) Then
                            OfferXml = File.ReadAllText(TempFileName)
                            If SbOffersXml.Length = 0 Then
                                SbOffersXml.Append(OfferXml)
                            Else
                                n = OfferXml.IndexOf("<PromoMaint")
                                SbOffersXml.Append(OfferXml.Substring(n, (OfferXml.Length - n)))
                            End If
                            File.Delete(TempFileName)
                        End If
                        If OfferXml.Length > 0 Then offersCount = offersCount + 1
                    Next

                Case ENGINE_ID.CPE
                    For Each row In dt.Rows
                        OfferXml = MyExportCpe.GetOfferXML(MyCommon.NZ(row.Item("OfferID"), 0))
                        If SbOffersXml.Length = 0 Then
                            SbOffersXml.Append(OfferXml)
                        Else
                            n = OfferXml.IndexOf("<Offer>")
                            SbOffersXml.Append(OfferXml.Substring(n, (OfferXml.Length - n)))
                        End If
                        If OfferXml.Length > 0 Then offersCount = offersCount + 1
                    Next

                Case ENGINE_ID.UE
                    For Each row In dt.Rows
                        OfferXml = m_ExportXMLUE.GetOfferXML(MyCommon.NZ(row.Item("OfferID"), 0))
                        If SbOffersXml.Length = 0 Then
                            SbOffersXml.Append(OfferXml)
                        Else
                            n = OfferXml.IndexOf("<Offer>")
                            SbOffersXml.Append(OfferXml.Substring(n, (OfferXml.Length - n)))
                        End If
                        If OfferXml.Length > 0 Then offersCount = offersCount + 1
                    Next
                Case Else

            End Select
            If (MyCommon.LRTadoConn.State = ConnectionState.Closed) Then MyCommon.Close_LogixRT()
            If (MyCommon.LXSadoConn.State = ConnectionState.Closed) Then MyCommon.Open_LogixXS()
        End If

        MaxOffers = offersCount

        'for test number of offers can the system handle
        'Dim str = "dt.Rows.Count = " & dt.Rows.Count & "   offersCount = " & offersCount.ToString & " From: " & startTime.ToString & " To: " & Date.Now.ToString
        'SbOffersXml.Append(str)
        Return SbOffersXml.ToString
    End Function

    Private Function IsStringBoolean(ByVal StrBool As String) As Boolean
        If StrBool.ToLower = "false" Or StrBool.ToLower = "true" Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function GetCustomerGroupID(ByVal ExtCardID As String, ByVal CardTypeID As Integer, ByVal TableName As String) As String
        Dim dt As DataTable
        Dim row As DataRow
        Dim CustomerGroupIDCondition As String = ""
        Dim CustGroupIDEqualTo As String = " OR " & TableName & ".CustomerGroupID = "

        If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

        MyCommon.QueryStr = "Select GroupMembership.CustomerGroupID from GroupMembership INNER JOIN CardIDs " & _
                            "ON GroupMembership.CustomerPK = CardIDs.CustomerPK Where CardIDs.ExtCardID = @ExtCardID" & _
                            " and CardIDs.CardTypeID = @CardTypeID "
        MyCommon.DBParameters.Add("@ExtCardID", SqlDbType.NVarChar).Value = ExtCardID.ConvertBlankIfNothing
        MyCommon.DBParameters.Add("@CardTypeID", SqlDbType.Int).Value = CardTypeID
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixXS)
        If (dt.Rows.Count > 0) Then
            For Each row In dt.Rows
                CustomerGroupIDCondition = CustomerGroupIDCondition & CustGroupIDEqualTo & MyCommon.NZ(row.Item("CustomerGroupID"), -1)
            Next
        End If
        'AppendToLog(ExtInterfaceID.ToString, "GetCustomerGroupID", "ExtCardID = " & ExtCardID, "QueryStr=" & MyCommon.QueryStr, 0, "Debug...")
        'AppendToLog(ExtInterfaceID.ToString, "GetCustomerGroupID", "ExtCardID = " & ExtCardID, "CustomerGroupIDCondition=" & CustomerGroupIDCondition, 0, "Debug...")
        Return CustomerGroupIDCondition
    End Function

    Private Function GetProductGroupID(ByVal ExtProductID As String, ByVal ProductTypeID As Integer, ByVal TableName As String) As String
        Dim dt As DataTable
        Dim row As DataRow
        Dim ProductGroupIDCondition As String = ""
        Dim ProdGroupIDEqualTo As String = " OR " & TableName & ".ProductGroupID = "

        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
        MyCommon.QueryStr = "Select ProdGroupItems.ProductGroupID from ProdGroupItems INNER JOIN Products " & _
                            "ON ProdGroupItems.ProductID = Products.ProductID Where Products.ExtProductID = @ExtProductID and Products.ProductTypeID = @ProductTypeID "
        MyCommon.DBParameters.Add("@ExtProductID", SqlDbType.NVarChar).Value = ExtProductID.ConvertBlankIfNothing
        MyCommon.DBParameters.Add("@ProductTypeID", SqlDbType.Int).Value = ProductTypeID
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        If (dt.Rows.Count > 0) Then
            For Each row In dt.Rows
                ProductGroupIDCondition = ProductGroupIDCondition & ProdGroupIDEqualTo & MyCommon.NZ(row.Item("ProductGroupID"), -1)
            Next
        End If

        Return ProductGroupIDCondition
    End Function

    Private Function GetLocationGroupID(ByVal ExtLocationCode As String) As String
        Dim dt As DataTable
        Dim row As DataRow
        Dim IsFirstRow As Boolean = True
        Dim LocationGroupIDCondition As String = ""
        Dim LocGroupIDEqualTo As String = " OR OfferLocations.LocationGroupID = "

        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
        MyCommon.QueryStr = "Select LocGroupItems.LocationGroupID from LocGroupItems INNER JOIN Locations " & _
                            "ON LocGroupItems.LocationID = Locations.LocationID Where Locations.ExtLocationCode = @ExtLocationCode"
        MyCommon.DBParameters.Add("@ExtLocationCode", SqlDbType.NVarChar).Value = ExtLocationCode.ConvertBlankIfNothing
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        If (dt.Rows.Count > 0) Then
            For Each row In dt.Rows
                If IsFirstRow Then
                    LocationGroupIDCondition = " and ( OfferLocations.LocationGroupID = " & MyCommon.NZ(row.Item("LocationGroupID"), -1)
                    IsFirstRow = False
                Else
                    LocationGroupIDCondition = LocationGroupIDCondition & LocGroupIDEqualTo & MyCommon.NZ(row.Item("LocationGroupID"), -1)
                End If
            Next
        End If
        'AppendToLog(ExtInterfaceID.ToString, "GetLocationGroupID", "ExtLocationCode = " & ExtLocationCode, "QueryStr=" & MyCommon.QueryStr, 0, "Debug...")
        'AppendToLog(ExtInterfaceID.ToString, "GetLocationGroupID", "ExtLocationCode = " & ExtLocationCode, "LocationGroupIDCondition=" & LocationGroupIDCondition, 0, "Debug...")
        Return LocationGroupIDCondition
    End Function

    Private Function GetCustomerGroupIDs(ByVal ExtCardID As String, ByVal CardTypeID As Integer) As String
        Dim dt As DataTable
        Dim row As DataRow

        Dim CustomerGroupIDs As String = GetAllCustomersAndCardholersGroupIDs()

        If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
        Dim ConnectorsExtCardId As String = MyCommon.Pad_ExtCardID(ExtCardID, CardTypeID)
        MyCommon.QueryStr = "Select Distinct GroupMembership.CustomerGroupID from GroupMembership INNER JOIN CardIDs " & _
                            "ON GroupMembership.CustomerPK = CardIDs.CustomerPK and GroupMembership.Deleted=0 " & _
                            "Where CardIDs.ExtCardID = @ExtCardID and CardIDs.CardTypeID = @CardTypeID"
        MyCommon.DBParameters.Add("@ExtCardID", SqlDbType.NVarChar).Value = MyCryptLib.SQL_StringEncrypt(ConnectorsExtCardId, True)
        MyCommon.DBParameters.Add("@CardTypeID", SqlDbType.Int).Value = CardTypeID
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixXS)
        If (dt.Rows.Count > 0) Then
            For Each row In dt.Rows
                CustomerGroupIDs = CustomerGroupIDs & ", " & MyCommon.NZ(row.Item("CustomerGroupID"), -1)
            Next
        End If

        AppendToLog(ExtInterfaceID.ToString, "GetCustomerGroupIDs", "GetCustomerGroupIDs = " & CustomerGroupIDs, "GetCustomerGroupIDs=" & CustomerGroupIDs, 0, "Debug...")
        AppendToLog(ExtInterfaceID.ToString, "GetCustomerGroupIDs", "ExtCardID = " & ExtCardID, "QueryStr=" & MyCommon.QueryStr, 0, "Debug...")
        Return CustomerGroupIDs
    End Function

    Private Function GetAllCustomersAndCardholersGroupIDs() As String
        Dim AllCustomersAndCardholersGroupIDs As String = ""
        Dim dt As DataTable
        Dim row As DataRow
        Dim IsFirstRow As Boolean = True

        MyCommon.QueryStr = "select CustomerGroupID from CustomerGroups with (NoLock) where (AnyCustomer=1 or AnyCardholder=1) and Deleted=0"
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        If dt.Rows.Count > 0 Then
            For Each row In dt.Rows
                If IsFirstRow Then
                    AllCustomersAndCardholersGroupIDs = MyCommon.NZ(row.Item("CustomerGroupID"), -1).ToString
                    IsFirstRow = False
                Else
                    AllCustomersAndCardholersGroupIDs = AllCustomersAndCardholersGroupIDs & ", " & MyCommon.NZ(row.Item("CustomerGroupID"), -1).ToString
                End If
            Next
        End If
        Return AllCustomersAndCardholersGroupIDs
    End Function

    Private Function GetProductGroupIDs(ByVal ExtProductID As String, ByVal ProductTypeID As Integer) As String
        Dim dt As DataTable
        Dim row As DataRow
        Dim ProductGroupIDs As String = GetAnyProductGroupID().ToString

        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
        MyCommon.QueryStr = "Select Distinct ProdGroupItems.ProductGroupID from ProdGroupItems INNER JOIN Products " & _
                            "ON ProdGroupItems.ProductID = Products.ProductID and ProdGroupItems.Deleted = 0 " & _
                            "Where Products.ExtProductID = @ExtProductID and Products.ProductTypeID = @ProductTypeID "
        MyCommon.DBParameters.Add("@ExtProductID", SqlDbType.NVarChar).Value = ExtProductID.ConvertBlankIfNothing
        MyCommon.DBParameters.Add("@ProductTypeID", SqlDbType.Int).Value = ProductTypeID
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        If (dt.Rows.Count > 0) Then
            For Each row In dt.Rows
                ProductGroupIDs = ProductGroupIDs & ", " & MyCommon.NZ(row.Item("ProductGroupID"), -1).ToString
            Next
        End If

        Return ProductGroupIDs
    End Function

    Private Function GetAnyProductGroupID() As Integer
        Dim AnyProductGroupID As Integer = -1
        Dim dt As DataTable
        MyCommon.QueryStr = "select ProductGroupID from ProductGroups with (NoLock) where AnyProduct=1 and Deleted=0"
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        If dt.Rows.Count > 0 Then
            AnyProductGroupID = MyCommon.NZ(dt.Rows(0).Item("ProductGroupID"), 0)
        End If
        Return AnyProductGroupID
    End Function

    Private Function GetLocationGroupIDs(ByVal ExtLocationCode As String) As String
        Dim dt As DataTable
        Dim row As DataRow
        Dim LocationGroupIDs As String = GetAllLocationGroupID().ToString

        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
        MyCommon.QueryStr = "Select Distinct LocGroupItems.LocationGroupID from LocGroupItems INNER JOIN Locations " & _
                            "ON LocGroupItems.LocationID = Locations.LocationID and LocGroupItems.Deleted=0 Where Locations.ExtLocationCode = @ExtLocationCode "
        MyCommon.DBParameters.Add("@ExtLocationCode", SqlDbType.NVarChar).Value = ExtLocationCode.ConvertBlankIfNothing
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        If (dt.Rows.Count > 0) Then
            For Each row In dt.Rows
                LocationGroupIDs = LocationGroupIDs & ", " & MyCommon.NZ(row.Item("LocationGroupID"), -1).ToString
            Next
        End If
        Return LocationGroupIDs
    End Function

    Private Function GetAllLocationGroupID() As Integer
        Dim AllLocGroupID As Integer = -1
        Dim dt As DataTable
        MyCommon.QueryStr = "select LocationGroupID from LocationGroups with (NoLock) where AllLocations=1 and Deleted=0"
        dt = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
        If dt.Rows.Count > 0 Then
            AllLocGroupID = MyCommon.NZ(dt.Rows(0).Item("LocationGroupID"), 0)
        End If
        Return AllLocGroupID
    End Function

End Class