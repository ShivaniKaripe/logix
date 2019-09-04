Imports System
Imports System.Web
Imports System.Web.Services
Imports System.Xml
Imports System.Data
Imports System.IO
Imports System.Collections.Generic

Imports Copient.CustomerLookup
Imports Copient

<WebService(Namespace:="http://www.copienttech.com/KioskAd/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class KioskAd
  Inherits System.Web.Services.WebService
  ' $Id: KioskAd.asmx 64228 2013-04-08 17:54:36Z ns185105 $
  ' version:7.3.1.138972.Official Build (SUSDAY10202)

  Public Const USE_DEFAULT_CARD_TYPE_ID As Integer = -1
  Private Const CONNECTOR_ID As Integer = 49
  Public Const CPE_OPERATE_AT_ENTERPRISE As Integer = 91

  Private LogFile As String = "KioskWSLog." & Date.Now.ToString("yyyyMMdd") & ".txt"
    Private MyCommon As New Copient.CommonInc
    Private MyCryptLib As New Copient.CryptLib

  Public Enum CustomerTypes As Integer
    CARDHOLDER = 0
    HOUSEHOLD = 1
    CAM = 2
  End Enum

  Private Enum PassThruRewardType As Integer
    SSA_Coupon = 4
    SSA_Targeted_Ad = 5
  End Enum

  Private Enum TargetedOfferType As Integer
    Activity = 0
    Coupon = 1
    Survey = 2
    Maintenance = 3
  End Enum

  Private Enum TargetedOfferPriority As Integer
    None = 0
    Low = 1
    Medium = 5
    High = 10
  End Enum

  Private Enum CouponStatusCodes As Integer
    INVALID_LOCATION = 13
    INVALID_COUPONLOCATION = 14
    INVALID_COUPON = 15
    VALIDATION_ATTEMPTED = 16
    REDEMPTION_ATTEMPTED = 17
    INVALID_MEMBER = 18
    INVALID_STRING = 19
    INVALID_CUSTOMERID = 20
    APPLICATION_EXCEPTION = 9999
  End Enum

  Public Const OPERATION_TAG_LIMIT As Integer = 6

#Region "ExcludeOffer"

  <WebMethod()> _
  Public Function ExcludeOffer(ByVal CardID As String, ByVal CardTypeID As Integer, ByVal OfferID As Integer) As Excluded
    Dim excluded As New Excluded
    Dim errorMessage1 As String = ""
    Dim errorMessage As String = ""

    Try
      If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
      If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

      CardID = transformCard(CardID, CardTypeID.ToString(), MyCommon, errorMessage1)
      Dim CustomerPK As Long = GetCustomerPK(CardID, CardTypeID, False, errorMessage)
      Dim OfferROID As Long = 0

	  If(Not String.IsNullOrEmpty(errorMessage1)) Then
		    errorMessage =errorMessage1
      ElseIf (MyCommon.Fetch_CPE_SystemOption(CPE_OPERATE_AT_ENTERPRISE) <> "1") Then
        errorMessage = "CPE is not configured to operate at the enterprise level. This functionality is not available."
      ElseIf (Not IsValidCardType(CardTypeID, errorMessage)) Then
        errorMessage = "Card type " & CardTypeID & " is an invalid card type. "
      ElseIf (CustomerPK = 0) Then
        errorMessage = "CardID " & CardID & " of type " & CardTypeID & " cannot be found. "
      ElseIf (Not DoesCPEOfferExist(OfferID, errorMessage)) Then
        errorMessage = "Offer #" & OfferID & " does not exist. "
      Else
        OfferROID = GetOfferROID(OfferID, errorMessage)
        MyCommon.QueryStr = "select CustomerGroupID from CPE_IncentiveCustomerGroups with (NoLock) where RewardOptionID=@OfferROID and ExcludedUsers=0 and Deleted=0;"
        MyCommon.DBParameters.Add("@OfferROID", SqlDbType.BigInt).Value = OfferROID
        Dim customerGroupDT As DataTable = MyCommon.ExecuteQuery(DataBases.LogixRT)
        If (customerGroupDT.Rows.Count > 0) Then
          Dim inGroup As Boolean = False
          For Each row As DataRow In customerGroupDT.Rows
            If (CustomerIsInGroup(CustomerPK, MyCommon.NZ(row.Item("CustomerGroupID"), 0), errorMessage)) Then inGroup = True
          Next
          If inGroup Then
            ExcludeCustomerFromOffer(CustomerPK, OfferID, OfferROID, errorMessage)  'Must exclude customer by putting them into the excluded customer group
          Else
            errorMessage = "CardID " & CardID & " is not eligible for offer " & OfferID & "."
          End If
        Else
          errorMessage = "Offer " & OfferID & " does not have a customer condition. "
        End If
      End If
      If (errorMessage = "") Then
        'excluded.Message = ""
        excluded.Success = True
      Else
        excluded.Success = False
        excluded.Message = errorMessage
      End If
    Catch ex As Exception
      MyCommon.Write_Log(LogFile, "Application exception: " & ex.ToString(), True)
      excluded.Success = False
      excluded.Message = errorMessage & " " & "Application exception: " & ex.ToString()
    Finally
      If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
      If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
    End Try

    Return excluded
  End Function

  Private Function DoesCPEOfferExist(ByVal OfferID As Integer, ByRef ErrorMessage As String) As Boolean
    Dim OfferExists As Boolean = False
    Dim dt As DataTable = Nothing

    Try
      MyCommon.QueryStr = "select IncentiveID from CPE_ST_Incentives with (NoLock) " & _
                          "where IncentiveID=@OfferID and EngineID in (2,3,6,9) and Deleted=0;"
      MyCommon.DBParameters.Add("@OfferID", SqlDbType.Int).Value = OfferID
      dt = MyCommon.ExecuteQuery(DataBases.LogixRT)
      If (dt.Rows.Count > 0) Then
        OfferExists = True
      End If
    Catch ex As Exception
      ErrorMessage += ex.ToString() & vbCrLf
    End Try

    Return OfferExists
  End Function

  Private Function ExcludeCustomerFromOffer(ByVal CustomerPK As Long, ByVal OfferID As Long, ByVal ROID As Long, ByRef ErrorMessage As String) As Boolean
    Dim Excluded As Boolean = False
    Dim dt As DataTable
    Dim ExcludedCustomerGroupID As Long = 0

    Try
      MyCommon.QueryStr = "select CustomerGroupID from CPE_ST_IncentiveCustomerGroups with (NoLock) " & _
                          "where RewardOptionID=@ROID and ExcludedUsers=1 and Deleted=0;"
      MyCommon.DBParameters.Add("@ROID", SqlDbType.BigInt).Value = ROID
      dt = MyCommon.ExecuteQuery(DataBases.LogixRT)
      If dt.Rows.Count = 0 Then
        'There's no existing excluded group in use by the offer, so return an error, ...
        ErrorMessage = "Offer " & OfferID & " does not have an excluded customer group."
      Else
        'An excluded group is already associated to the offer, so use that.
        ExcludedCustomerGroupID = MyCommon.NZ(dt.Rows(0).Item("CustomerGroupID"), 0)
        If CustomerIsInGroup(CustomerPK, ExcludedCustomerGroupID, ErrorMessage) Then
          'Customer's already in the excluded group, so no need to continue
          ErrorMessage = "Customer is already excluded from offer " & OfferID & " because of membership in customer group " & ExcludedCustomerGroupID & "."
        Else
          If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
          'Add the customer to the excluded group.
          MyCommon.QueryStr = "Insert into GroupMembership with (RowLock) (CustomerGroupID, CustomerPK, Manual, Deleted, CMOAStatusFlag, TCRMAStatusFlag, CPEStatusFlag) values (@ExcludedCustomerGroupID, @CustomerPK, 1, 0, 2, 2,-5);"
          MyCommon.DBParameters.Add("@ExcludedCustomerGroupID", SqlDbType.BigInt).Value = ExcludedCustomerGroupID
          MyCommon.DBParameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
          MyCommon.ExecuteNonQuery(DataBases.LogixXS)
          MyCommon.Activity_Log(4, ExcludedCustomerGroupID, 1, Copient.PhraseLib.Lookup("history.cgroup-add", 1))
          Excluded = True
        End If
      End If
    Catch ex As Exception
      ErrorMessage += ex.ToString() & vbCrLf
    End Try

    Return Excluded
  End Function

#End Region

#Region "GetOfferCoupon"

  Private Function GetPassThruCoupon(ByVal TargetedOffer As Integer, ByVal RewardType As PassThruRewardType, ByRef ErrorMessage As String) As DataTable
    If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()

    MyCommon.QueryStr = "select RO.IncentiveID,  " & _
                        "    PT.ParamName,       " & _
                        "    TV.Value,           " & _
                        "    TV.PTPKID " & _
                        "from CPE_RewardOptions RO " & _
                        "inner join CPE_ST_Deliverables D on RO.RewardOptionID=D.RewardOptionID " & _
                        "inner join CPE_ST_PassThrus Pass on D.DeliverableID=Pass.DeliverableID " & _
                        "inner join PassThruTierValues TV on Pass.PKID=TV.PTPKID " & _
                        "inner join PassThruPresTags PT on TV.PassThruPresTagID=PT.PassThruPresTagID " & _
                        "where Pass.Deleted = 0 And Pass.PassThruRewardID =@RewardType And RO.IncentiveID = @TargetedOffer;"
    MyCommon.DBParameters.Add("@RewardType", SqlDbType.Int).Value = RewardType
    MyCommon.DBParameters.Add("@TargetedOffer", SqlDbType.Int).Value = TargetedOffer
    Dim offerDT As DataTable = MyCommon.ExecuteQuery(DataBases.LogixRT)
    Dim rewardDT As DataTable = New DataTable()
    Try

      If (offerDT.Rows.Count > 0) Then
        'Create a table to use to put values into xml
        rewardDT.Rows.Add(rewardDT.NewRow())

        'Only use the first PassThruTierPKID since there can be more than one PassThruReward of any given type

        Dim colName As String = "IncentiveID"
        rewardDT.Columns.Add(colName, GetType(Integer))
        rewardDT.Rows(0).Item(colName) = MyCommon.NZ(offerDT.Rows(0).Item("IncentiveID"), 0)

        Dim passThruTierPK As Integer = MyCommon.NZ(offerDT.Rows(0).Item("PTPKID"), 0)
        For Each rewardRow As DataRow In offerDT.Select("PTPKID = " & passThruTierPK)
          colName = MyCommon.NZ(rewardRow("ParamName"), "")
          colName = colName.Replace(" ", "")
          rewardDT.Columns.Add(colName, GetType(String))
          rewardDT.Rows(0).Item(colName) = MyCommon.NZ(rewardRow("Value"), "")
        Next
        For Each dc As DataColumn In rewardDT.Columns
          If (dc.ColumnName = "EffectiveDate") Then
            If (rewardDT.Rows(0).Item("EffectiveDate") = "") Then
              rewardDT.Rows(0).Item("EffectiveDate") = DateTime.Now.ToString("yyyy-MM-dd")
            End If
          End If
        Next
      Else
        rewardDT.Dispose()
        ErrorMessage += "No coupon values were found for Targeted Offer #" & TargetedOffer.ToString() & ". "
        'MyCommon.Write_Log(LogFile, ErrorMessage, True)
      End If

    Catch ex As Exception
      rewardDT.Dispose()
      'ErrorMessage += "Application exception: " & ex.ToString()
      MyCommon.Write_Log(LogFile, "Application exception: " & ex.ToString(), True)
    End Try
    If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()

    Return rewardDT
  End Function

  Private Function WriteTargetedOfferCoupon_XML(ByVal OfferDT As DataTable, ByRef ErrorMessage As String) As XmlDocument
    'Write to XML
    Dim xmlSettings As XmlWriterSettings = New XmlWriterSettings()
    xmlSettings.Indent = True
    xmlSettings.Encoding = System.Text.UTF8Encoding.UTF8
    Dim stringWriter As StringWriter = New StringWriter()
    Dim xmlWriter As XmlWriter = xmlWriter.Create(stringWriter, xmlSettings)
    Dim xmlDoc As New XmlDocument
    If (OfferDT.Columns.Count > 0 AndAlso OfferDT.Rows.Count > 0) Then
      Try
        xmlWriter.WriteStartDocument()
        xmlWriter.WriteStartElement("TargetedOfferCoupon")

        For Each dc As DataColumn In OfferDT.Columns

          Dim colname As String = dc.ColumnName
          Dim colval As String = MyCommon.NZ(OfferDT.Rows(0).Item(colname), String.Empty)
          xmlWriter.WriteElementString(colname, colval)

        Next

        xmlWriter.WriteEndElement() 'End TargetedOfferCoupon element
        xmlWriter.WriteEndDocument()
        xmlWriter.Flush()
        xmlWriter.Close()
        xmlDoc.PreserveWhitespace = True
        xmlDoc.LoadXml(stringWriter.ToString())
      Catch exArg As ArgumentException

        ErrorMessage += "Argument exception " & exArg.ToString() & ". "
        MyCommon.Write_Log(LogFile, "Argument exception " & exArg.ToString() & ". ", True)

      Catch ex As Exception

        ErrorMessage += "Application exception: " & ex.ToString() & vbCrLf
        MyCommon.Write_Log(LogFile, ErrorMessage, True)

      End Try
    Else

      ErrorMessage += "No SSA Coupon Data was found. "
      MyCommon.Write_Log(LogFile, ErrorMessage, True)

    End If
    Return xmlDoc
  End Function

  <WebMethod()> _
  Public Function GetOfferCoupon(ByVal TargetedOffer As Integer) As String
    'Provides the coupon information needed to present a member with a targeted coupon.
    Dim xmlDoc As New XmlDocument
    Dim errorMessage As String = ""

    Dim xmlDT As DataTable = New DataTable()
    xmlDT = GetPassThruCoupon(TargetedOffer, PassThruRewardType.SSA_Coupon, errorMessage)
    If (xmlDT.Rows.Count > 0) Then xmlDoc = WriteTargetedOfferCoupon_XML(xmlDT, errorMessage)
    Dim returnString As String = ""
    If (errorMessage = "") Then
      returnString = FormatXmlString(xmlDoc.OuterXml)
    Else
      returnString = ResponseXML("GetOfferCoupon", errorMessage, False)
    End If
    Return returnString

  End Function

  Private Function call_pt_Generate_UPC(ByVal inputBarcode As Int64, ByVal location As String, ByVal SVProgramID As Int64, ByVal effectiveDate As DateTime, ByVal expireDate As DateTime, _
      ByVal customerPK As Int64, ByVal redemptionRestrictionID As Integer, ByVal issuingTransactionID As String, ByVal issueDate As DateTime, ByVal validLocation As Int64, _
     ByVal ROID As Int64) As String

    Try

      MyCommon.QueryStr = "dbo.pt_Generate_UPC"
      MyCommon.Open_LogixXS()
      MyCommon.Open_LXSsp()

      MyCommon.LXSsp.Parameters.Add("@InputBarCode", SqlDbType.BigInt).Value = inputBarcode
      MyCommon.LXSsp.Parameters.Add("@LocationID", SqlDbType.NVarChar, 20).Value = location
      MyCommon.LXSsp.Parameters.Add("@SVProgramID", SqlDbType.BigInt).Value = SVProgramID
      MyCommon.LXSsp.Parameters.Add("@EffectiveDate", SqlDbType.DateTime).Value = effectiveDate
      MyCommon.LXSsp.Parameters.Add("@ExpireDate", SqlDbType.DateTime).Value = expireDate
      MyCommon.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = customerPK
      MyCommon.LXSsp.Parameters.Add("@RedemptionRestrictionID", SqlDbType.Int).Value = redemptionRestrictionID
      MyCommon.LXSsp.Parameters.Add("@IssuingTransactionID", SqlDbType.NVarChar,128).Value = issuingTransactionID
      MyCommon.LXSsp.Parameters.Add("@IssueDate", SqlDbType.DateTime).Value = issueDate
      MyCommon.LXSsp.Parameters.Add("@ValidLocation", SqlDbType.BigInt).Value = validLocation
      MyCommon.LXSsp.Parameters.Add("@RewardOptionID", SqlDbType.BigInt).Value = ROID
      MyCommon.LXSsp.Parameters.Add("@UPCCode", SqlDbType.Char, 12).Direction = ParameterDirection.Output
      MyCommon.LXSsp.Parameters.Add("@retval", SqlDbType.Int).Direction = ParameterDirection.ReturnValue

      MyCommon.LXSsp.ExecuteNonQuery()

      If (MyCommon.LXSsp.Parameters("@UPCCode").Value Is Nothing Or IsDBNull(MyCommon.LXSsp.Parameters("@UPCCode").Value)) Then
        ' this should never be reached; if pt_Generate_UPC can't return a valid barcode, it should throw an error.
        Throw New ApplicationException(String.Format("No valid barcode was found! {0}", MyCommon.LXSsp.Parameters("@retval").Value))
      End If

      Return MyCommon.LXSsp.Parameters("@UPCCode").Value

    Finally

      MyCommon.Close_LXSsp()

    End Try

  End Function

  Private Function getLocationIDFromExtLocationId(ByVal extLocationID As String) As Int64
    Dim errmsg As String = String.Empty
    Dim locationid As Int64 = GetLocationID(extLocationID, errmsg)
    If Not String.IsNullOrEmpty(errmsg) Then
      Throw New ApplicationException(String.Format("An error occurred in validating the location code: {0}", errmsg))
    End If
    Return locationid

  End Function

  Private Function getRedemptionRestrictionIDFromSVProgram(ByVal SVProgramID As Int64) As Integer

    MyCommon.QueryStr = "SELECT TOP 1 [RedemptionRestrictionID] FROM [StoredValuePrograms] WHERE [SVProgramID] = @SVProgramID;"
    MyCommon.DBParameters.Add("@SVProgramID", SqlDbType.BigInt).Value = SVProgramID
    Using dt As DataTable = MyCommon.ExecuteQuery(DataBases.LogixRT)
      If dt.Rows.Count > 0 Then
        Return MyCommon.NZ(dt.Rows(0).Item("RedemptionRestrictionID"), 0)
      End If
    End Using
    Return 0
  End Function

  ' See svn://svn.copienttech.com:443/Logix/trunk/internal_documentation/requirements/SpeedwayUniqueuseandredemptionitems.doc sec 1.2.4
  <WebMethod()> _
  Public Function GetCouponBarcode(ByVal baseBarcode As Int64, ByVal extLocationID As String, ByVal SVProgramID As Int64, _
        ByVal effectiveDate As String, ByVal expirationDate As DateTime, ByVal customerCardNumber As String, ByVal customerCardType As Int16, _
        ByVal issuingTransactionID As String, ByVal issuedate As String, _
        ByVal ROID As Int64) As String
  Dim ErrorMessage As String = ""    'added
    Dim Result As String = ""
    'Dim locationid As Int64 = getLocationIDFromExtLocationId(extLocationID)  //commented
    'added
    Dim errmsg As String = String.Empty
    Dim locationid As Int64 = GetLocationID(extLocationID, errmsg)
    If Not String.IsNullOrEmpty(errmsg) Then
      Try
        Throw New ApplicationException(String.Format("An error occurred in validating the location code: {0}", errmsg))
      Catch ex As System.Exception
        Result = ex.Message
      End Try
      Return Result
    End If
    Dim effectiveTimestamp As DateTime = DateTime.Now

    If effectiveDate Is Nothing OrElse effectiveDate.Trim = "" Then
      Try
        Throw New ApplicationException(String.Format("EffectiveDate was not provided."))
      Catch ex As System.Exception
        Result = ex.Message
      End Try
      Return Result
    Else
      If (Not String.IsNullOrEmpty(HttpUtility.HtmlDecode(effectiveDate))) Then
        Try
          effectiveTimestamp = DateTime.Parse(HttpUtility.HtmlDecode(effectiveDate))
        Catch ex As Exception
          Result = "EffectiveDate provided is not a valid date format."
          Return Result
        End Try
      End If
    End If

    'Dim effectiveTimestamp As DateTime = DateTime.Now
    'If (Not String.IsNullOrEmpty(HttpUtility.HtmlDecode(effectiveDate))) Then
    '    effectiveTimestamp = DateTime.Parse(HttpUtility.HtmlDecode(effectiveDate))
    'End If

    customerCardNumber = transformCard(customerCardNumber, customerCardType.ToString(), MyCommon, errorMessage)
        If (Not String.IsNullOrEmpty(errorMessage)) Then
           Try
            Throw New ApplicationException(errorMessage)
          Catch ex As Exception
            Result = ex.Message
            Return Result
          End Try
         End If
    
    Dim custlookup As Copient.CustomerLookup = New Copient.CustomerLookup(MyCommon)
	Dim retcode As Copient.CustomerAbstract.RETURN_CODE = 0
	Dim custpk As Int64 = custlookup.GetCustomerPK(customerCardNumber, customerCardType, retcode)
	If (retcode <> RETURN_CODE.OK) Then
      Try
        Throw New ApplicationException(String.Format("Customer Card {0} / {1} could not be verified", customerCardNumber, customerCardType))
      Catch ex As Exception
        Result = ex.Message
      End Try
      Return Result
    End If

    Dim couponIssueTimestamp As DateTime = DateTime.Now
    If issuedate Is Nothing OrElse issuedate.Trim = "" Then
      Try
        Throw New ApplicationException(String.Format("IssueDate was not provided."))

      Catch ex As Exception
        Result = ex.Message
      End Try
      Return Result
    Else
      If (Not String.IsNullOrEmpty(HttpUtility.HtmlDecode(issuedate))) Then
        Try
          couponIssueTimestamp = DateTime.Parse(HttpUtility.HtmlDecode(issuedate))
        Catch ex As Exception
          Result = "IssueDate provided is not a valid date format."
          Return Result
        End Try
      End If
    End If

    'If (Not String.IsNullOrEmpty(HttpUtility.HtmlDecode(issuedate))) Then
    '    couponIssueTimestamp = DateTime.Parse(HttpUtility.HtmlDecode(issuedate))
    'End If

    Dim redemptionRestrictionID As Integer = getRedemptionRestrictionIDFromSVProgram(SVProgramID)

    'Return call_pt_Generate_UPC( baseBarcode, locationid, SVProgramID, effectiveTimestamp, expirationDate, custpk, redemptionRestrictionID, issuingTransactionID, couponIssueTimestamp, locationid )
    Return call_pt_Generate_UPC(baseBarcode, extLocationID, SVProgramID, effectiveTimestamp, expirationDate, custpk, redemptionRestrictionID, issuingTransactionID, couponIssueTimestamp, locationid, ROID)

  End Function

  Private Function call_pa_StoredValue_ExpirationDate(ByVal svprogramid As Int64, ByVal offerid As Integer, ByVal issuedate As DateTime) As DateTime

    MyCommon.QueryStr = "dbo.pa_StoredValue_ExpirationDate"
    MyCommon.Open_LogixRT()
    MyCommon.Open_LRTsp()

    MyCommon.LRTsp.Parameters.Add("@SVProgramID", SqlDbType.BigInt).Value = svprogramid
    MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.Int).Value = offerid
    MyCommon.LRTsp.Parameters.Add("@issueDateTime", SqlDbType.DateTime).Value = issuedate

    MyCommon.LRTsp.Parameters.Add("@expiresAt", SqlDbType.DateTime).Direction = ParameterDirection.Output

    MyCommon.LRTsp.ExecuteNonQuery()

    Dim expiredate As DateTime = MyCommon.LRTsp.Parameters("@expiresAt").Value
    MyCommon.Close_LRTsp()

    Return expiredate

  End Function

  ' See svn://svn.copienttech.com:443/Logix/trunk/internal_documentation/requirements/SpeedwayUniqueuseandredemptionitems.doc sec 1.2.3
  <WebMethod()> _
  Public Function GetCouponExpirationDate(ByVal SVProgramID As Int64, ByVal OfferID As Integer, ByVal issueDate As String) As DateTime

    Dim couponIssueTimestamp As DateTime = DateTime.Now

    If (Not String.IsNullOrEmpty(HttpUtility.HtmlDecode(issueDate))) Then
      couponIssueTimestamp = DateTime.Parse(HttpUtility.HtmlDecode(issueDate))
    End If

    Return call_pa_StoredValue_ExpirationDate(SVProgramID, OfferID, couponIssueTimestamp)

  End Function

#End Region

#Region "GetMemberOffers"

  Public Function GetTargetedOffers(ByRef ErrorMessage As String) As DataTable
    Dim offersDT As New DataTable
    Try
      If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
      'Get Offers and customer groups that are targeted web offers
      'Offers dates are limited by the production end date being greater than or equal to today.
      MyCommon.QueryStr = "select OID.OfferID,CG.CustomerGroupID,CG.ExcludedUsers,I.IncentiveName as Name,I.Description,I.StartDate,I.EndDate from OfferIDs OID " & _
                          "inner join CPE_ST_RewardOptions RO on OID.OfferID=RO.IncentiveID " & _
                          "inner join CPE_ST_Incentives I on OID.OfferID=I.IncentiveID " & _
                          "inner join CPE_IncentiveCustomerGroups CG on RO.RewardOptionID=CG.RewardOptionID " & _
                          "where OID.EngineID = 2 And OID.EngineSubTypeID = 0 And CG.Deleted = 0 And RO.Deleted = 0 And I.Deleted = 0 " & _
                          "and I.EndDate >= dateadd(dd, datediff(dd, 0, getdate()), 0) and I.StartDate <= dateadd(dd, datediff(dd, 0, getdate()), 0)" & _
                          "order by OID.OfferID"
      offersDT = MyCommon.LRT_Select()
    Catch ex As Exception
      offersDT.Dispose()
      ErrorMessage += "Application exception: " & ex.ToString()
      MyCommon.Write_Log(LogFile, ErrorMessage, True)
    Finally
      If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
    End Try
    Return offersDT
  End Function

  Public Function CustomerMembership(ByVal CardID As String, ByVal CardTypeID As Integer, ByRef ErrorMessage As String) As List(Of Integer)
    Dim groupList As New List(Of Integer)
    Try
      If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
      MyCommon.QueryStr = "select distinct GM.CustomerGroupID from CardIDs C " & _
                          "inner Join GroupMembership GM on C.CustomerPK=GM.CustomerPK " & _
                          "where C.ExtCardID=@CardID and C.CardTypeID=@CardTypeID and C.CardStatusID = 1 and GM.Deleted=0;"
            MyCommon.DBParameters.Add("@CardID", SqlDbType.NVarChar, 400).Value = MyCryptLib.SQL_StringEncrypt(CardID, True)
            MyCommon.DBParameters.Add("@CardTypeID", SqlDbType.Int).Value = CardTypeID
            Dim customerDT As DataTable = MyCommon.ExecuteQuery(DataBases.LogixXS)
            groupList.Add(1) 'Add all Customers
            groupList.Add(2) 'Add all card holders
            If (customerDT.Rows.Count > 0) Then
                Dim groupRow As DataRow
                'groupList.Add(1) 'Add all Customers
                'groupList.Add(2) 'Add all card holders
                For Each groupRow In customerDT.Rows
                    groupList.Add(MyCommon.NZ(groupRow.Item("CustomerGroupID"), 0))
                Next
                'Else
                '  ErrorMessage += "No group membership found for Customer #" & CardID & " of type ID " & CardTypeID & ". "
                '  MyCommon.Write_Log(LogFile, ErrorMessage, True)
            End If
        Catch ex As Exception
            groupList.Clear()
            ErrorMessage += "Application exception: " & ex.ToString()
            MyCommon.Write_Log(LogFile, ErrorMessage, True)
        Finally
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try

        Return groupList
    End Function

    Public Function WriteTargetedOfferList_XML(ByVal OfferDT As DataTable, ByRef ErrorMessage As String) As XmlDocument
        Dim xmlSettings As XmlWriterSettings = New XmlWriterSettings()
        xmlSettings.Indent = True
        xmlSettings.Encoding = System.Text.UTF8Encoding.UTF8
        Dim stringWriter As StringWriter = New StringWriter()
        Dim xmlWriter As XmlWriter = xmlWriter.Create(stringWriter, xmlSettings)
        Dim xmlDoc As New XmlDocument
        If (OfferDT.Columns.Count > 0 AndAlso OfferDT.Rows.Count > 0) Then
            Try
                xmlWriter.WriteStartDocument()
                xmlWriter.WriteStartElement("TargetedOffers")
                For Each offerRow As DataRow In OfferDT.Rows
                    xmlWriter.WriteStartElement("TargetedOffer")
                    xmlWriter.WriteElementString("ID", MyCommon.NZ(offerRow.Item("OfferID"), ""))
                    xmlWriter.WriteElementString("TargetedOfferType", MyCommon.NZ(offerRow.Item("OfferType"), ""))
                    xmlWriter.WriteElementString("Title", MyCommon.NZ(offerRow.Item("Name"), ""))
                    xmlWriter.WriteElementString("Description", MyCommon.NZ(offerRow.Item("Description"), ""))
                    xmlWriter.WriteElementString("Start", MyCommon.NZ(offerRow.Item("StartDate"), ""))
                    xmlWriter.WriteElementString("End", MyCommon.NZ(offerRow.Item("EndDate"), ""))
                    xmlWriter.WriteElementString("SmallImageName", MyCommon.NZ(offerRow.Item("SmallImageName"), ""))
                    xmlWriter.WriteElementString("LargeImageName", MyCommon.NZ(offerRow.Item("LargeImageName"), ""))
                    xmlWriter.WriteElementString("Invitation", MyCommon.NZ(offerRow.Item("Invitation"), ""))
                    xmlWriter.WriteElementString("TargetedOfferPriority", MyCommon.NZ(offerRow.Item("Priority"), ""))
                    xmlWriter.WriteElementString("Points", MyCommon.NZ(offerRow.Item("Points"), ""))
                    xmlWriter.WriteElementString("ProgramID", MyCommon.NZ(offerRow.Item("ProgramID"), ""))
                    xmlWriter.WriteEndElement() 'End TargetedOffer element
                Next
                xmlWriter.WriteEndElement() 'End TargetedOffers element

                xmlWriter.WriteEndDocument()
                xmlWriter.Flush()
                xmlWriter.Close()
                xmlDoc.PreserveWhitespace = True
                xmlDoc.LoadXml(stringWriter.ToString())
            Catch exArg As ArgumentException
                'ErrorMessage += "Argument exception " & exArg.ToString() & ". "
                MyCommon.Write_Log(LogFile, "Argument exception " & exArg.ToString() & ". ", True)
            Catch ex As Exception
                ErrorMessage += "Application exception: " & ex.ToString() & vbCrLf
                MyCommon.Write_Log(LogFile, ErrorMessage, True)
            End Try
        Else
            ErrorMessage += "No SSA Coupon Data was found. "
            MyCommon.Write_Log(LogFile, ErrorMessage, True)
        End If
        Return xmlDoc
    End Function

    Private Sub AddPassThru(ByRef OfferDT As DataTable)
        'Add pass thru columns to offer data table
        OfferDT.Columns.Add("OfferType", GetType(String))
        OfferDT.Columns.Add("SmallImageName", GetType(String))
        OfferDT.Columns.Add("LargeImageName", GetType(String))
        OfferDT.Columns.Add("Priority", GetType(String))
        OfferDT.Columns.Add("Invitation", GetType(String))
        OfferDT.Columns.Add("Points", GetType(String))
        OfferDT.Columns.Add("ProgramID", GetType(String))

        Dim passthruDT As New DataTable
        Dim removeDT As DataTable = OfferDT.Clone()
        For Each offerRow As DataRow In OfferDT.Rows
            Dim TempErrorMessage As String = ""
            passthruDT = New DataTable
            passthruDT = GetPassThruCoupon(offerRow.Item("OfferID"), PassThruRewardType.SSA_Targeted_Ad, TempErrorMessage)
            If (passthruDT.Rows.Count > 0) Then
                Dim offerType As String = ""
                Select Case MyCommon.NZ(passthruDT.Rows(0).Item("Offertype"), -1)
                    Case TargetedOfferType.Activity
                        offerType = "Activity"
                    Case TargetedOfferType.Coupon
                        offerType = "Coupon"
                    Case TargetedOfferType.Maintenance
                        offerType = "Maintenance"
                    Case TargetedOfferType.Survey
                        offerType = "Survey"
                End Select
                offerRow.Item("Offertype") = offerType
                offerRow.Item("SmallImageName") = MyCommon.NZ(passthruDT.Rows(0).Item("Smallimagename"), "")
                offerRow.Item("LargeImageName") = MyCommon.NZ(passthruDT.Rows(0).Item("Largeimagename"), "")
                Dim offerPriority As String = ""
                Select Case MyCommon.NZ(passthruDT.Rows(0).Item("Priority"), 0)
                    Case TargetedOfferPriority.High
                        offerPriority = "High"
                    Case TargetedOfferPriority.Medium
                        offerPriority = "Medium"
                    Case TargetedOfferPriority.Low
                        offerPriority = "Low"
                    Case TargetedOfferPriority.None
                        offerPriority = "None"
                End Select
                offerRow.Item("Priority") = offerPriority
                offerRow.Item("Invitation") = MyCommon.NZ(passthruDT.Rows(0).Item("Invitation"), "")
                offerRow.Item("Points") = MyCommon.NZ(passthruDT.Rows(0).Item("Points"), "")
                offerRow.Item("ProgramID") = MyCommon.NZ(passthruDT.Rows(0).Item("ProgramID"), "")
            Else
                removeDT.ImportRow(offerRow)
            End If
        Next
        'Remove rows that do not have a SSA Targeted Ad pass thru reward
        For Each removeRow As DataRow In removeDT.Rows
            For Each tempRow As DataRow In OfferDT.Select("OfferID=" & removeRow.Item("OfferID"), "OfferID desc")
                OfferDT.Rows.Remove(tempRow)
            Next
        Next
    End Sub

    <WebMethod()> _
    Public Function GetMemberOffers(ByVal CardID As String, ByVal CardTypeID As Integer) As String
        'Provides the TargetedOffer list for which the member is eligible.
        Dim ErrorMessage As String = ""
        Dim LogErrorMessage As String = ""
        CardID = transformCard(CardID, CardTypeID.ToString(), MyCommon, ErrorMessage)
        Dim xmlDoc As New XmlDocument
        Try
            Dim xmlDT As DataTable
            xmlDT = GetTargetedOffers(ErrorMessage)

            If (xmlDT.Rows.Count > 0) Then
                Dim customerGroups As List(Of Integer) = CustomerMembership(CardID, CardTypeID, ErrorMessage)
                Dim foundDT As DataTable = xmlDT.Clone()
                For Each foundRow As DataRow In xmlDT.Rows
                    For Each groupID As Integer In customerGroups
                        If (foundRow.Item("CustomerGroupID") = groupID) Then
                            foundDT.ImportRow(foundRow)
                        End If
                    Next
                Next
                xmlDT.Clear()
                xmlDT = foundDT.Copy()
                foundDT.Dispose()

                Dim includedDT As DataTable = xmlDT.Clone()
                Dim removeOffers As New List(Of Integer)
                For Each excludedRow As DataRow In xmlDT.Select("ExcludedUsers=1")
                    removeOffers.Add(MyCommon.NZ(excludedRow.Item("OfferID"), 0))
                Next

                If (removeOffers.Count > 0) Then
                    Dim lastOfferID As Integer = 0
                    Dim keepOffer As Boolean = True
                    For Each keepRow As DataRow In xmlDT.Rows
                        keepOffer = True
                        For Each offerID As Integer In removeOffers
                            If (keepRow.Item("OfferID") = offerID) Then keepOffer = False
                        Next
                        If (keepOffer AndAlso keepRow.Item("OfferID") <> lastOfferID) Then
                            includedDT.ImportRow(keepRow)
                            lastOfferID = keepRow.Item("OfferID")
                        End If
                    Next

                    xmlDT.Clear()
                    xmlDT = includedDT.Copy()
                    includedDT.Dispose()
                End If

                AddPassThru(xmlDT)

                If (xmlDT.Rows.Count > 0) Then xmlDoc = WriteTargetedOfferList_XML(xmlDT, ErrorMessage)
            Else
                ErrorMessage += "GetMemberOffers: No offers were found for CustomerID " & CardID.ToString() & " of type ID " & CardTypeID.ToString() & ". "
                LogErrorMessage += "GetMemberOffers: No offers were found for CustomerID " & MaskHelper.MaskCard(CardID.ToString(), CardTypeID.ToString()) & " of type ID " & CardTypeID.ToString() & ". "
                MyCommon.Write_Log(LogFile, LogErrorMessage, True)
            End If
        Catch ex As Exception
            LogErrorMessage += "GetMemberOffers: Application exception: " & ex.ToString()
            MyCommon.Write_Log(LogFile, LogErrorMessage, True)
        End Try

        Dim returnString As String = ""
        If (ErrorMessage = "") Then
            returnString = FormatXmlString(xmlDoc.OuterXml)
        Else
            returnString = ResponseXML("GetMemberOffers", ErrorMessage, False)
        End If
        Return returnString

    End Function
#End Region

    Private Function GetOfferROID(ByVal OfferID As Long, ByRef ErrorMessage As String) As Long
        Dim ROID As Long = 0
        Dim dt As DataTable

        Try
            MyCommon.QueryStr = "select RewardOptionID from CPE_ST_RewardOptions with (NoLock) where IncentiveID=@OfferID and Deleted=0;"
            MyCommon.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
            dt = MyCommon.ExecuteQuery(DataBases.LogixRT)
            If dt.Rows.Count > 0 Then
                ROID = MyCommon.NZ(dt.Rows(0).Item("RewardOptionID"), 0)
            End If
        Catch ex As Exception
            ErrorMessage += ex.ToString() & vbCrLf
        End Try

        Return ROID
    End Function

    Private Function GetLocationID(ByVal ExtLocationCode As String, ByRef ErrorMessage As String) As Long
        Dim LocationID As Long = 0
        'Dim LocationID As Long
        Dim dt As DataTable
        If ExtLocationCode Is Nothing OrElse ExtLocationCode.Trim = "" Then
            ErrorMessage = "Failure: ExtLocation code is not provided."
        Else
            Try
                If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
                MyCommon.QueryStr = "select LocationID from Locations with (NoLock) where ExtLocationCode=@ExtLocationCode and Deleted=0;"
                MyCommon.DBParameters.Add("@ExtLocationCode", SqlDbType.NVarChar, 20).Value = ExtLocationCode
                dt = MyCommon.ExecuteQuery(DataBases.LogixRT)
                If dt.Rows.Count > 0 Then
                    LocationID = MyCommon.NZ(dt.Rows(0).Item("LocationID"), 0)
                Else
                    ErrorMessage = "Failure: ExtLocationID not found"
                End If
            Catch ex As Exception
                ErrorMessage += ex.ToString() & vbCrLf
            End Try
        End If

        Return LocationID
    End Function

    Private Function IsValidCardType(ByVal CardTypeID As Integer, ByRef ErrorMessage As String) As Boolean
        Dim ValidType As Boolean = False
        Dim dt As DataTable

        Try
            MyCommon.QueryStr = "select CardTypeID from CardTypes with (NoLock) where CardTypeID=@CardTypeID;"
            MyCommon.DBParameters.Add("@CardTypeID", SqlDbType.Int).Value = CardTypeID
            dt = MyCommon.ExecuteQuery(DataBases.LogixXS)
            ValidType = (dt.Rows.Count > 0)
        Catch ex As Exception
            ErrorMessage += ex.ToString() & vbCrLf
        End Try

        Return ValidType
    End Function

    Private Function GetCustomerPK(ByVal ExtCardID As String, ByVal CardTypeID As Integer, ByVal CreateIfNotFound As Boolean, ByRef ErrorMessage As String) As Long
        Dim CustomerPK As Long = 0
        Dim dt As DataTable
        Dim dtPref As DataTable
        Dim bAllowBinPref As Boolean = False
        Dim objbinPrefix As New Copient.CustomizedCustomerInquiry(MyCommon.Get_Install_Path() & "/AgentFiles/CustomizedCustomerInquiryCard.config")
        Dim prefix As String = ""
        Dim binID As String = ""
        Dim prefID As Integer = 0
        Dim prefValue As String = ""
        Dim UpdateCode As Integer = 0

        Try
            If CardTypeID = USE_DEFAULT_CARD_TYPE_ID Then
                CardTypeID = MyCommon.Extract_Val(MyCommon.Fetch_SystemOption(30))
            End If
            If IsValidCardType(CardTypeID, ErrorMessage) Then
                If CreateIfNotFound Then
                    MyCommon.QueryStr = "dbo.pa_EOC_GetOrCreateCustomer"
                    MyCommon.Open_LXSsp()
                    MyCommon.LXSsp.Parameters.Add("@ExtCardID", SqlDbType.NVarChar, 400).Value = MyCryptLib.SQL_StringEncrypt(ExtCardID, True)
                    MyCommon.LXSsp.Parameters.Add("@CardTypeID", SqlDbType.Int).Value = CardTypeID
                    MyCommon.LXSsp.Parameters.Add("@CustomerTypeID", SqlDbType.Int).Value = CardTypeID
                    MyCommon.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Direction = ParameterDirection.Output
                    MyCommon.LXSsp.Parameters.Add("@ExtCardIDOriginal", SqlDbType.NVarChar, 400).Value = MyCryptLib.SQL_StringEncrypt(ExtCardID)
                    MyCommon.LXSsp.ExecuteNonQuery()
                    CustomerPK = MyCommon.LXSsp.Parameters("@CustomerPK").Value
                    MyCommon.Close_LXSsp()

                    'Associating the card to the bin preference
                    If MyCommon.PMRTadoConn.State <> ConnectionState.Open Then MyCommon.Open_PrefManRT()
                    MyCommon.QueryStr = "select OptionValue  from SystemOptions with (NoLock) where OptionID = 43;"
                    dtPref = MyCommon.PMRT_Select()
                    If dtPref.Rows.Count > 0 Then
                        If Not IsDBNull(dtPref.Rows(0)("OptionValue")) Then
                            bAllowBinPref = IIf(dtPref.Rows(0)("OptionValue") = "1", True, False)
                        End If
                    End If
                    If bAllowBinPref AndAlso CustomerPK > 0 Then
                        prefix = ExtCardID.Substring(0, 7)
                        binID = objbinPrefix.GetBinNumberFromPrefix(prefix)
                        If binID.Length < 2 Then
                            binID = "0" + binID
                        End If
                        MyCommon.QueryStr = "select PreferenceID, PreferenceValue from PreferenceBinRangeMap with (NoLock) where BinRange = '" & binID & "' and BinPrefix = '" & prefix & "';"
                        dtPref = MyCommon.PMRT_Select()
                        If dtPref.Rows.Count > 0 Then
                            For Each row As DataRow In dtPref.Rows
                                prefID = MyCommon.NZ(row.Item("PreferenceID"), 0)
                                prefValue = MyCommon.NZ(row.Item("PreferenceValue"), "")
                                If (prefID > 0 AndAlso Not String.IsNullOrEmpty(prefValue)) Then
                                    MyCommon.QueryStr = "dbo.pm_CustomerPreference_Update"
                                    MyCommon.Open_LXSsp()
                                    MyCommon.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
                                    MyCommon.LXSsp.Parameters.Add("@PreferenceID", SqlDbType.BigInt).Value = prefID
                                    MyCommon.LXSsp.Parameters.Add("@Value", SqlDbType.NVarChar, 250).Value = Left(prefValue, 250)
                                    MyCommon.LXSsp.Parameters.Add("@UpdateTime", SqlDbType.DateTimeOffset).Value = Date.Now()
                                    MyCommon.LXSsp.Parameters.Add("@LastChannelID", SqlDbType.Int).Value = 3
                                    MyCommon.LXSsp.Parameters.Add("@UpdateCode", SqlDbType.Int).Direction = ParameterDirection.Output
                                    MyCommon.LXSsp.ExecuteNonQuery()
                                    UpdateCode = MyCommon.LXSsp.Parameters("@UpdateCode").Value
                                    If UpdateCode < 0 Then
                                        MyCommon.Write_Log(LogFile, "Failed to insert/update preference to customer", True)
                                    End If
                                    MyCommon.Close_LXSsp()
                                End If
                            Next
                        End If
                    End If
                Else
                    MyCommon.QueryStr = "select CustomerPK from CardIDs with (NoLock) where CardTypeID=@CardTypeID and ExtCardID=@ExtCardID;"
                    MyCommon.DBParameters.Add("@CardTypeID", SqlDbType.Int).Value = CardTypeID
                    MyCommon.DBParameters.Add("@ExtCardID", SqlDbType.NVarChar, 400).Value = MyCryptLib.SQL_StringEncrypt(MyCommon.Parse_Quotes(ExtCardID), True)
                    dt = MyCommon.ExecuteQuery(DataBases.LogixXS)
                    If dt.Rows.Count > 0 Then
                        CustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)
                    End If
                End If
            End If
        Catch ex As Exception
            ErrorMessage += ex.ToString() & vbCrLf
        End Try

        Return CustomerPK
    End Function

    Private Function GetCustomerTypeID(ByVal CustomerPK As Long, ByRef ErrorMessage As String) As Integer
        Dim CustomerTypeID As Integer = -1
        Dim dt As DataTable

        Try
            MyCommon.QueryStr = "select CustomerTypeID from Customers with (NoLock) where CustomerPK=@CustomerPK;"
            MyCommon.DBParameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
            dt = MyCommon.ExecuteQuery(DataBases.LogixXS)
            If dt.Rows.Count > 0 Then
                CustomerTypeID = MyCommon.NZ(dt.Rows(0).Item("CustomerTypeID"), 0)
            End If
        Catch ex As Exception
            ErrorMessage += ex.ToString() & vbCrLf
        End Try

        Return CustomerTypeID
    End Function

    Private Function CustomerIsInGroup(ByVal CustomerPK As Long, ByVal CustomerGroupID As Long, ByRef ErrorMessage As String) As Boolean
        Dim InGroup As Boolean = False
        Dim dt As DataTable
        Dim NewCardholdersGroupID As Long = 0

        Try
            MyCommon.QueryStr = "select CustomerGroupID from CustomerGroups with (NoLock) " & _
                                "where NewCardholders=1 and Deleted=0;"
            dt = MyCommon.LRT_Select
            If dt.Rows.Count > 0 Then
                NewCardholdersGroupID = MyCommon.NZ(dt.Rows(0).Item("CustomerGroupID"), 0)
            End If

            If (CustomerGroupID = NewCardholdersGroupID) Then
                'New Cardholders group -- we will assume the customer is extant, not new, and so not a part of this group.
            ElseIf (CustomerGroupID = 1) Or (CustomerGroupID = 2) Then
                'Any Customer or Any Cardholder -- any cardholder will always be present in these.
                InGroup = True
            Else
                MyCommon.QueryStr = "select MembershipID from GroupMembership with (NoLock) " & _
                                    "where CustomerPK=@CustomerPK and CustomerGroupID=@CustomerGroupID and Deleted=0;"
                MyCommon.DBParameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
                MyCommon.DBParameters.Add("@CustomerGroupID", SqlDbType.BigInt).Value = CustomerGroupID
                dt = MyCommon.ExecuteQuery(DataBases.LogixXS)
                If dt.Rows.Count > 0 Then
                    InGroup = True
                End If
            End If
        Catch ex As Exception
            ErrorMessage += ex.ToString() & vbCrLf
        End Try

        Return InGroup
    End Function

    Private Function FormatXmlString(ByVal XmlString As String) As String
        Dim FormattedXml As String = XmlString
        Dim Lines(-1) As String
        Dim i As Integer
        Dim sb As New StringBuilder()

        If FormattedXml IsNot Nothing Then
            Lines = FormattedXml.Split(ControlChars.CrLf.ToCharArray)
            For i = 0 To Lines.GetUpperBound(0)
                If Lines(i).Trim <> "" Then
                    sb.Append(Lines(i))
                    sb.Append(ControlChars.CrLf)
                End If
            Next
            FormattedXml = sb.ToString
            sb = Nothing

            '' remove empty tags
            'FormattedXml = Regex.Replace(FormattedXml, "(<[^>]+/>)", String.Empty)

            FormattedXml = FormattedXml.Replace("encoding=""utf-16""", "encoding=""utf-8""")
        End If

        Return FormattedXml
    End Function

    Private Function ResponseXML(ByVal MethodName As String, ByVal Message As String, Optional ByVal Success As Boolean = True) As String
        Dim xmlSettings As XmlWriterSettings = New XmlWriterSettings()
        xmlSettings.Indent = True
        xmlSettings.IndentChars = ControlChars.Tab
        xmlSettings.Encoding = System.Text.UTF8Encoding.UTF8
        xmlSettings.NewLineChars = ControlChars.CrLf
        xmlSettings.NewLineHandling = NewLineHandling.Replace
        Dim stringWriter As StringWriter = New StringWriter()
        Dim xmlWriter As XmlWriter = xmlWriter.Create(stringWriter, xmlSettings)

        xmlWriter.WriteStartDocument()
        xmlWriter.WriteStartElement("KioskAd")
        xmlWriter.WriteStartElement(MethodName)

        xmlWriter.WriteAttributeString("success", Success.ToString().ToLower())

        If (Success) Then
            If (Not Message = "") Then xmlWriter.WriteAttributeString("message", Message)
        Else
            xmlWriter.WriteStartElement("ERROR")
            xmlWriter.WriteAttributeString("message", Message)
            xmlWriter.WriteEndElement() 'ERROR
        End If
        xmlWriter.WriteEndElement() 'MethodName

        xmlWriter.WriteEndElement() 'KioskAd
        xmlWriter.WriteEndDocument()
        xmlWriter.Flush()
        xmlWriter.Close()

        Dim returnXML As String = stringWriter.ToString()
        ' workaround for problem where encoding is always set to utf-16 no matter
        ' what you set for the encoding in the XMLWriterSettings.Encoding
        If (returnXML IsNot Nothing) Then
            returnXML = returnXML.Replace("encoding=""utf-16""", "encoding=""utf-8""")
        End If

        Return returnXML
    End Function

    Private Function shouldDoCustomizedCustomerInquiry(ByRef MyCommon As Copient.CommonInc) As Boolean
        Const USE_CUSTOMIZED_CUSTOMER_INQUIRY As Integer = 107
        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
        Return (MyCommon.Fetch_SystemOption(USE_CUSTOMIZED_CUSTOMER_INQUIRY) = 1)
        If MyCommon.LRTadoConn.State = ConnectionState.Open Then MyCommon.Close_LogixRT()
    End Function

    Private Function transformCard(ByVal card As String, ByVal cardType As String, ByRef MyCommon As Copient.CommonInc, ByRef ErrorMessage As String) As String
        Const STANDARD_CUSTOMER_CARD_TYPEID As String = "0"
        If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()

        If (cardType = STANDARD_CUSTOMER_CARD_TYPEID AndAlso Not isEmpty(card) AndAlso shouldDoCustomizedCustomerInquiry(MyCommon)) Then
            Try
                Return validateCard(card, MyCommon, ErrorMessage)
            Catch ex As Exception
                ErrorMessage += ex.ToString() & vbCrLf
            End Try
        End If
        If MyCommon.LXSadoConn.State = ConnectionState.Open Then MyCommon.Close_LogixXS()
        Return card
    End Function

    Private Function validateCard(ByVal card As String, ByRef MyCommon As Copient.CommonInc, ByRef ErrorMessage As String) As String
        Const PHYSICAL_CARD_LENGTH As Integer = 12
        Const MEMBER_ID_LENGTH As Integer = 15
        card = Trim(card)

        Dim cardConverter As New Copient.CustomizedCustomerInquiry(MyCommon.Get_Install_Path() & "/AgentFiles/CustomizedCustomerInquiryCard.config")
        If (card.Length = PHYSICAL_CARD_LENGTH And IsNumeric(card)) Then
            card = cardConverter.getMemberIdFromCardNumber(card)
        ElseIf (card.Length = MEMBER_ID_LENGTH And IsNumeric(card)) Then
            Dim physical_card As String = cardConverter.getCardNumberFromMemberId(Long.Parse(card))
        Else
            Throw New ArgumentException(String.Format("{0} ({1})", Copient.PhraseLib.Lookup("term.invalid-cust-specific-card-number", 1), card))
        End If

        Return card
    End Function

    Private Function isEmpty(ByVal s As String) As Boolean
        Return s Is Nothing OrElse s.Trim.Length < 1
    End Function

    <WebMethod()> _
    Public Function RedemptionItemsCategories() As RedemptionItemCategories
        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
        MyCommon.QueryStr = "Select OfferCategoryID,Description,SortOrder,IconFileName from OfferCategories where Deleted=0;"
        Dim offerCategoryDT As DataTable = MyCommon.LRT_Select()
        Dim offerCategoryCollection As New RedemptionItemCategories
        ReDim offerCategoryCollection.Categories(offerCategoryDT.Rows.Count - 1)
        If (offerCategoryDT.Rows.Count > 0) Then
            Dim rowCount As Integer = 0
            For Each categoryRow As DataRow In offerCategoryDT.Rows
                offerCategoryCollection.Categories(rowCount) = New OfferCategory(MyCommon.NZ(categoryRow.Item("OfferCategoryID"), 0), MyCommon.NZ(categoryRow.Item("Description"), ""), _
                  MyCommon.NZ(categoryRow.Item("SortOrder"), 0), MyCommon.NZ(categoryRow.Item("IconFileName"), ""))
                rowCount += 1
            Next
        End If
        If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
        Return offerCategoryCollection
    End Function

    Public Function FindCPEAllTerminal() As Integer
        MyCommon.QueryStr = "select TerminalTypeID from TerminalTypes where EngineID=2 and AnyTerminal=1;"
        Dim terminalDT As DataTable = MyCommon.LRT_Select()
        If (terminalDT.Rows.Count > 0) Then
            Dim termID As Integer = MyCommon.NZ(terminalDT.Rows(0).Item("TerminalTypeID"), 0)
            If (termID <> 0) Then
                Return termID
            Else
                Throw New ApplicationException("The TerminalTypeID for ""All CPE terminals"" is NULL in the TerminalTypes table.")
            End If
        Else
            Throw New ApplicationException("The terminal ""All CPE terminals"" was not found in the TerminalTypes table.")
        End If
    End Function

    Public Sub TerminalExists(ByVal TerminalTypeID As Integer)
        MyCommon.QueryStr = "select TerminalTypeID from TerminalTypes where TerminalTypeID=@TerminalTypeID;"
        MyCommon.DBParameters.Add("@TerminalTypeID", SqlDbType.Int).Value = TerminalTypeID
        Dim terminalDT As DataTable = MyCommon.ExecuteQuery(DataBases.LogixRT)
        If (terminalDT.Rows.Count > 0) Then
            Dim termID As Integer = MyCommon.NZ(terminalDT.Rows(0).Item("TerminalTypeID"), 0)
            If termID = 0 Then Throw New ArgumentException("Terminal (" & TerminalTypeID & ") does not exist.")
        Else
            Throw New ArgumentException("Terminal (" & TerminalTypeID & ") does not exist.")
        End If
    End Sub

    Public Sub LocationExists(ByVal ExtLocationCode As String)
        MyCommon.QueryStr = "select ExtLocationCode from Locations where ExtLocationCode=@ExtLocationCode;"
        MyCommon.DBParameters.Add("@ExtLocationCode", SqlDbType.NVarChar, 20).Value = ExtLocationCode
        Dim locDT As DataTable = MyCommon.ExecuteQuery(DataBases.LogixRT)
        If (locDT.Rows.Count > 0) Then
            Dim locCode As String = MyCommon.NZ(locDT.Rows(0).Item("ExtLocationCode"), "")
            If locCode = "" Then Throw New ArgumentException("Location (" & ExtLocationCode & ") does not exist.")
        Else
            Throw New ArgumentException("Location (" & ExtLocationCode & ") does not exist.")
        End If
    End Sub

    Public Function GetLocationGroupID(ByVal ExtLocationCode As String) As String
        LocationExists(ExtLocationCode)
        MyCommon.QueryStr = "select LGI.LocationGroupID from Locations L with (NoLock) " & _
                            "inner join LocGroupItems LGI with (NoLock) on L.LocationID=LGI.LocationID " & _
                            "where L.ExtLocationCode=@ExtLocationCode and L.Deleted=0 AND LGI.Deleted=0;"
        MyCommon.DBParameters.Add("@ExtLocationCode", SqlDbType.NVarChar, 20).Value = ExtLocationCode
        Dim locGroupDT As DataTable = MyCommon.ExecuteQuery(DataBases.LogixRT)
        Dim locGroupID As String = ""
        For Each row As DataRow In locGroupDT.Rows
            locGroupID &= MyCommon.NZ(row.Item("LocationGroupID"), "") & ","
        Next
        If locGroupID <> "" Then
            locGroupID = Left(locGroupID, Len(locGroupID) - 1)
        End If
        Return locGroupID
    End Function

    <WebMethod()> _
    Public Function RedemptionItemsAvaliable(ByVal CostCenter As String, ByVal TerminalID As Integer) As RedemptionItemsAvailable
        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
        Const RedemptionText_Index As Integer = 0
        Const PointsValue_Index As Integer = 1
        Const ImageName_Index As Integer = 2
        Const MessageText_Index As Integer = 3
        Const ProgramID_Index As Integer = 4

        Const CPE_ALL_STORES As Integer = 1
        Dim locationGroupID As String = ""
        Try
            locationGroupID = GetLocationGroupID(CostCenter)
        Catch ex As Exception
            Throw New CostCenterException(ex.Message)
        End Try

        Try
            TerminalExists(TerminalID)
        Catch ex As Exception
            Throw New TerminalIDException(ex.Message)
        End Try

        Dim availableItemCollection As New RedemptionItemsAvailable
        Try
            'Get offers that have SSA Redemption Item rewards
            Dim DBParam_FindCPEAllTerminal As Integer = FindCPEAllTerminal()
            If String.IsNullOrWhiteSpace(locationGroupID) Then
                MyCommon.QueryStr = "select CD.DeliverableID,CD.RewardOptionID,OutputID from CPE_ST_Deliverables CD with (NoLock)" & _
                                   "inner join CPE_ST_RewardOptions RO with (NoLock) on RO.RewardOptionID=CD.RewardOptionID " & _
                                   "inner join CPE_ST_Incentives I with (NoLock) on I.IncentiveID=RO.IncentiveID " & _
                                   "inner join CPE_ST_OfferTerminals OT with (NoLock) on OT.OfferID=RO.IncentiveID " & _
                                   "inner join OfferLocations OL with (NoLock) on OL.OfferID=RO.IncentiveID " & _
                                   "where exists (select PKID from CPE_ST_PassThrus where PassThruRewardID=(select PassThruRewardID from passthrurewards where Name='SSA Redemption Item') and CD.OutputID=CPE_ST_PassThrus.PKID) " & _
                                   "and DeliverableTypeID=12 and CD.Deleted=0 and I.EngineID=2 and I.EngineSubTypeID=0 " & _
                                   "and OT.TerminalTypeID in (@FindCPEAllTerminal, @TerminalID) and OL.Deleted=0 and OL.LocationGroupID in (@CPE_ALL_STORES) " & _
                                   "and I.StartDate <= GETDATE() and I.EndDate >= CONVERT(date,GETDATE());"
            Else
                MyCommon.QueryStr = "select CD.DeliverableID,CD.RewardOptionID,OutputID from CPE_ST_Deliverables CD with (NoLock)" & _
                                    "inner join CPE_ST_RewardOptions RO with (NoLock) on RO.RewardOptionID=CD.RewardOptionID " & _
                                    "inner join CPE_ST_Incentives I with (NoLock) on I.IncentiveID=RO.IncentiveID " & _
                                    "inner join CPE_ST_OfferTerminals OT with (NoLock) on OT.OfferID=RO.IncentiveID " & _
                                    "inner join OfferLocations OL with (NoLock) on OL.OfferID=RO.IncentiveID  AND OL.OfferID NOT IN (SELECT OfferID FROM OfferLocations WHERE Excluded = 1 AND LocationGroupID in (" & locationGroupID & " )) " & _
                                    "where exists (select PKID from CPE_ST_PassThrus where PassThruRewardID=(select PassThruRewardID from passthrurewards where Name='SSA Redemption Item') and CD.OutputID=CPE_ST_PassThrus.PKID) " & _
                                    "and DeliverableTypeID=12 and CD.Deleted=0 and I.EngineID=2 and I.EngineSubTypeID=0 " & _
                                    "and OT.TerminalTypeID in (@FindCPEAllTerminal, @TerminalID) and OL.Deleted=0 and OL.LocationGroupID in (@CPE_ALL_STORES, " & locationGroupID & ") " & _
                                    "and I.StartDate <= GETDATE() and I.EndDate >= CONVERT(date,GETDATE());"
            End If
            MyCommon.DBParameters.Add("@FindCPEAllTerminal", SqlDbType.Int).Value = DBParam_FindCPEAllTerminal
            MyCommon.DBParameters.Add("@TerminalID", SqlDbType.Int).Value = TerminalID
            MyCommon.DBParameters.Add("@CPE_ALL_STORES", SqlDbType.Int).Value = CPE_ALL_STORES

            Dim offerDT As DataTable = MyCommon.ExecuteQuery(DataBases.LogixRT)

            If (offerDT.Rows.Count > 0) Then
                ReDim availableItemCollection.Items(offerDT.Rows.Count - 1)
                Dim rowCount As Integer = 0
                Dim passThruDT As DataTable
                Dim incentiveDT As DataTable
                For Each offerRow As DataRow In offerDT.Rows
                    'Find the incentive information for each offer
                    MyCommon.QueryStr = "select I.IncentiveID,I.IncentiveName,I.CreatedDate,I.LastUpdate,I.StartDate,I.EndDate,I.PromoClassID from CPE_ST_Incentives I with (NoLock) " & _
                                        "inner join CPE_ST_RewardOptions RO on RO.IncentiveID=I.IncentiveID " & _
                                        "where RO.RewardOptionID=@RewardOptionID;"
                    MyCommon.DBParameters.Add("@RewardOptionID", SqlDbType.BigInt).Value = Convert.ToInt64(MyCommon.NZ(offerRow.Item("RewardOptionID"), 0))
                    incentiveDT = MyCommon.ExecuteQuery(DataBases.LogixRT)
                    availableItemCollection.Items(rowCount) = New RedemptionItem()
                    If (incentiveDT.Rows.Count > 0) Then
                        availableItemCollection.Items(rowCount).LoadOffer(MyCommon.NZ(incentiveDT.Rows(0).Item("IncentiveID"), 0), MyCommon.NZ(incentiveDT.Rows(0).Item("IncentiveName"), ""), incentiveDT.Rows(0).Item("CreatedDate"), _
                                                                          incentiveDT.Rows(0).Item("LastUpdate"), incentiveDT.Rows(0).Item("StartDate"), incentiveDT.Rows(0).Item("EndDate"), MyCommon.NZ(incentiveDT.Rows(0).Item("PromoClassID"), 0))
                    End If
                    'Find the Pass thru tier values for each offer
                    MyCommon.QueryStr = "Select PTTV.Value from PassThruTierValues PTTV with (NoLock) " & _
                                "inner join CPE_ST_PassThruTiers PTT with (NoLock) on PTTV.PTPKID=PTT.PTPKID " & _
                                "inner join CPE_ST_PassThrus PT with (NoLock) on PTT.PTPKID=PT.PKID " & _
                                "where PT.DeliverableID=@DeliverableID and PTT.PTPKID=@OutputID;"
                    MyCommon.DBParameters.Add("@DeliverableID", SqlDbType.Int).Value = Convert.ToInt32(MyCommon.NZ(offerRow.Item("DeliverableID"), 0))
                    MyCommon.DBParameters.Add("@OutputID", SqlDbType.Int).Value = Convert.ToInt32(MyCommon.NZ(offerRow.Item("OutputID"), 0))
                    passThruDT = MyCommon.ExecuteQuery(DataBases.LogixRT)
                    'There should be only 5 columns returned for the SSA Redemption Item pass thru. Even if the fields are left empty they should still be there.
                    If (passThruDT.Rows.Count = 5) Then
                        availableItemCollection.Items(rowCount).LoadPassThru(MyCommon.NZ(passThruDT.Rows(RedemptionText_Index).Item("Value"), ""), MyCommon.NZ(passThruDT.Rows(PointsValue_Index).Item("Value"), ""), _
                        MyCommon.NZ(passThruDT.Rows(ImageName_Index).Item("Value"), ""), MyCommon.NZ(passThruDT.Rows(MessageText_Index).Item("Value"), ""), MyCommon.NZ(passThruDT.Rows(ProgramID_Index).Item("Value"), ""))
                    End If
                    rowCount += 1
                Next
            End If
        Catch ex As Exception
            Throw New RedemptionItemsAvailableException(ex.ToString())
        End Try

        Return availableItemCollection
    End Function

    'Validate a coupon
    <WebMethod()> _
    Public Function ValidateCoupon(ByVal LocationCode As String, ByVal CustomerID As String, _
                                    ByVal CustomerTypeID As Integer, ByVal Coupon As String, _
                                    ByVal TransactionID As String, ByVal CurrentTime As String) As System.Data.DataSet
        Dim dt As System.Data.DataTable
        Dim dt2 As System.Data.DataTable
        Dim dtMemberRedemptionId As System.Data.DataTable
        Dim dtCouponValidation As System.Data.DataTable
        Dim dtGetCustomerPK As System.Data.DataTable
        Dim dtStatus As System.Data.DataTable
        Dim row As System.Data.DataRow
        Dim drValidation As DataRow
        Dim RetCode As CouponStatusCodes
        Dim RetMessage As String = ""
        Dim bOpenedXS, bOpenedRT, bOpenedWH As Boolean
        Dim ResultSet As New System.Data.DataSet("ValidateCoupon")
        Dim CurrentTimeVar As DateTime
        Dim MemRedemID As Integer = -1
        Dim ErrorMessage As String = ""

        If CurrentTime = "" Then
            CurrentTimeVar = Now
        Else
            CurrentTimeVar = Date.ParseExact(CurrentTime, "yyyy-MM-ddTHH:mm:ss", Nothing) 'System.Globalization.CultureInfo.InvariantCulture)
        End If

        'Initialize the status table, which will report the success or failure of the CustomerDetails operation
        dtStatus = New DataTable
        dtStatus.TableName = "ValidationStatus"
        dtStatus.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        Try
            bOpenedRT = False
            bOpenedXS = False
            bOpenedWH = False
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                MyCommon.Open_LogixRT()
                bOpenedRT = True
            End If
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then
                MyCommon.Open_LogixXS()
                bOpenedXS = True
            End If
            If MyCommon.LWHadoConn.State = ConnectionState.Closed Then
                bOpenedWH = True
                MyCommon.Open_LogixWH()
            End If

            CustomerID = transformCard(CustomerID, CustomerTypeID.ToString(), MyCommon, ErrorMessage)

            MyCommon.QueryStr = "select top 1 CustomerPK from CardIDs with (NoLock) where ExtCardID=@CustomerID AND CardTypeID=@CustomerTypeID;"
            MyCommon.DBParameters.Add("@CustomerID", SqlDbType.NVarChar, 400).Value = MyCryptLib.SQL_StringEncrypt(CustomerID, True)
            MyCommon.DBParameters.Add("@CustomerTypeID", SqlDbType.Int).Value = CustomerTypeID
            dtGetCustomerPK = MyCommon.ExecuteQuery(DataBases.LogixXS)

            MyCommon.QueryStr = "select locationid from locations where ExtLocationCode = @LocationCode;"
            MyCommon.DBParameters.Add("@LocationCode", SqlDbType.NVarChar, 20).Value = LocationCode
            dt = MyCommon.ExecuteQuery(DataBases.LogixRT)

            MyCommon.QueryStr = "select ValidLocation, RedemptionRestrictionID, SVProgramID from BarcodeDetails where Barcode = @Coupon;"
            MyCommon.DBParameters.Add("@Coupon", SqlDbType.NVarChar, 14).Value = Coupon
            dt2 = MyCommon.ExecuteQuery(DataBases.LogixXS)

            If dt2.Rows.Count = 0 Then
            Else
                MyCommon.QueryStr = "select MemberRedemptionId from StoredValuePrograms with (NoLock) where SVProgramID = @SVProgramID;"
                MyCommon.DBParameters.Add("@SVProgramID", SqlDbType.BigInt).Value = Convert.ToInt64(MyCommon.NZ(dt2.Rows(0).Item("SVProgramID"), -1))
                dtMemberRedemptionId = MyCommon.ExecuteQuery(DataBases.LogixRT)
                If dtMemberRedemptionId.Rows.Count > 0 Then
                    MemRedemID = MyCommon.NZ(dtMemberRedemptionId.Rows(0).Item("MemberRedemptionID"), -1)
                Else
                    MemRedemID = -1
                End If
            End If

            If (Not String.ISNullOrEmpty(ErrorMessage)) Then
                row = dtStatus.NewRow()
                row.Item("StatusCode") = CouponStatusCodes.INVALID_CUSTOMERID
                row.Item("Description") = ErrorMessage
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
            ElseIf LocationCode = "" Or CustomerID = "" Or Coupon = "" Then
                'One of the string fields were submitted empty
                row = dtStatus.NewRow()
                row.Item("StatusCode") = CouponStatusCodes.INVALID_STRING
                row.Item("Description") = "Failure: An empty value was passed for location, customer, or coupon."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
            ElseIf dt.Rows.Count = 0 Then
                'ExtLocationCode not found
                row = dtStatus.NewRow()
                row.Item("StatusCode") = CouponStatusCodes.INVALID_LOCATION
                row.Item("Description") = "Failure: Invalid ExtLocationCode."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
            ElseIf dt2.Rows.Count = 0 Then
                'Coupon not found
                row = dtStatus.NewRow()
                row.Item("StatusCode") = CouponStatusCodes.INVALID_COUPON
                row.Item("Description") = "Failure: Coupon Number not found."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
            ElseIf Not (CustomerID = "0" AndAlso CustomerTypeID = 0) AndAlso dtGetCustomerPK.Rows.Count = 0 AndAlso MemRedemID <> 0 Then 'If Member Redemption Restriction ID (MemRedemID) is 0, any customer can redeem this coupon
                'Customer not found
                row = dtStatus.NewRow()
                row.Item("StatusCode") = CouponStatusCodes.INVALID_CUSTOMERID
                row.Item("Description") = "Customer: " & CustomerID & " with CustomerType: " & CustomerTypeID & " not found."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
            Else

                MyCommon.Open_LogixXS()
                MyCommon.QueryStr = "dbo.pt_ValidateMember"
                MyCommon.Open_LXSsp()
                MyCommon.LXSsp.Parameters.Add("@MemberRedemptionID", SqlDbType.Int).Value = MemRedemID
                MyCommon.LXSsp.Parameters.Add("@UniqueBarCode", SqlDbType.NVarChar, 14).Value = Coupon
                MyCommon.LXSsp.Parameters.Add("@ExtCardID", SqlDbType.NVarChar, 400).Value = MyCryptLib.SQL_StringEncrypt(CustomerID, True)
                MyCommon.LXSsp.Parameters.Add("@CardTypeID", SqlDbType.Int).Value = CustomerTypeID
                MyCommon.LXSsp.Parameters.Add("@ValidMember", SqlDbType.Int).Direction = ParameterDirection.Output
                MyCommon.LXSsp.ExecuteNonQuery()

                If MyCommon.LXSsp.Parameters("@ValidMember").Value = 1 Then

                    MyCommon.QueryStr = "dbo.pt_ValidateLocation"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@ValidLocationID", SqlDbType.BigInt).Value = MyCommon.NZ(dt2.Rows(0).Item("ValidLocation"), 0)
                    MyCommon.LRTsp.Parameters.Add("@RedemptionRestrictionID", SqlDbType.Int).Value = MyCommon.NZ(dt2.Rows(0).Item("RedemptionRestrictionID"), 0)
                    MyCommon.LRTsp.Parameters.Add("@ExtLocationCode", SqlDbType.NVarChar, 20).Value = LocationCode
                    MyCommon.LRTsp.Parameters.Add("@ValidLocation", SqlDbType.Bit).Direction = ParameterDirection.Output
                    MyCommon.LRTsp.ExecuteNonQuery()

                    If MyCommon.LRTsp.Parameters("@ValidLocation").Value = True Then
                        'Coupon validation attempted for this location and coupon
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = CouponStatusCodes.VALIDATION_ATTEMPTED
                        row.Item("Description") = "Success."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()

                        MyCommon.QueryStr = "dbo.pt_ValidateCoupon"
                        MyCommon.Open_LXSsp()
                        MyCommon.LXSsp.Parameters.Add("@UniqueBarCode", SqlDbType.NVarChar, 12).Value = Coupon
                        MyCommon.LXSsp.Parameters.Add("@posDateTime", SqlDbType.DateTime).Value = CurrentTimeVar
                        MyCommon.LXSsp.Parameters.Add("@ExtCardID", SqlDbType.NVarChar, 400).Value = MyCryptLib.SQL_StringEncrypt(CustomerID, True)
                        MyCommon.LXSsp.Parameters.Add("@CardTypeID", SqlDbType.Int).Value = CustomerTypeID
                        MyCommon.LXSsp.Parameters.Add("@ValidLocation", SqlDbType.Bit).Value = 1
                        MyCommon.LXSsp.Parameters.Add("@CouponReasonCodeDesc", SqlDbType.NVarChar, 1000).Direction = ParameterDirection.Output
                        MyCommon.LXSsp.Parameters.Add("@Valid", SqlDbType.Bit).Direction = ParameterDirection.Output
                        MyCommon.LXSsp.ExecuteNonQuery()

                        'Record whether or not the coupon passed validation
                        dtCouponValidation = New DataTable
                        dtCouponValidation.TableName = "CouponValidation"
                        dtCouponValidation.Columns.Add("Valid", System.Type.GetType("System.Boolean"))
                        dtCouponValidation.Columns.Add("CouponReasonCodeDesc", System.Type.GetType("System.String"))

                        drValidation = dtCouponValidation.NewRow()
                        drValidation.Item("CouponReasonCodeDesc") = MyCommon.NZ(MyCommon.LXSsp.Parameters("@CouponReasonCodeDesc").Value, "")
                        drValidation.Item("Valid") = MyCommon.NZ(MyCommon.LXSsp.Parameters("@Valid").Value, 0)
                        dtCouponValidation.Rows.Add(drValidation)
                        dtCouponValidation.AcceptChanges()

                        ResultSet.Tables.Add(dtStatus.Copy())
                        ResultSet.Tables.Add(dtCouponValidation.Copy())

                        'Report Coupon Validation Errors
                        If (ResultSet.Tables.Item("CouponValidation").Rows(0).Item("Valid") IsNot Nothing) AndAlso _
                           (ResultSet.Tables.Item("CouponValidation").Rows(0).Item("Valid") <> 1) Then
                            MyCommon.QueryStr = Nothing
                            MyCommon.QueryStr = "dbo.pt_CouponError_Insert"
                            MyCommon.Open_LWHsp()
                            MyCommon.LWHsp.Parameters.Add("@Barcode", SqlDbType.NVarChar, 14).Value = Coupon
                            MyCommon.LWHsp.Parameters.Add("@Description", SqlDbType.NVarChar, 255).Value = ResultSet.Tables.Item("CouponValidation").Rows(0).Item("CouponReasonCodeDesc")
                            MyCommon.LWHsp.Parameters.Add("@ErrorCode", SqlDbType.BigInt).Value = ResultSet.Tables.Item("CouponValidation").Rows(0).Item("Valid")
                            MyCommon.LWHsp.Parameters.Add("@ErrorDate", SqlDbType.DateTime).Value = Now
                            MyCommon.LWHsp.Parameters.Add("@LocationCode", SqlDbType.NVarChar, 20).Value = LocationCode
                            MyCommon.LWHsp.Parameters.Add("@TransNum", SqlDbType.NVarChar, 128).Value = TransactionID.ToString  'A Transaction ID of 0 means this validation occurred independent of a transaction
                            MyCommon.LWHsp.ExecuteNonQuery()
                            MyCommon.Close_LWHsp()
                        End If

                    Else
                        'Coupon not valid for this location
                        row = dtStatus.NewRow()
                        row.Item("StatusCode") = CouponStatusCodes.INVALID_COUPONLOCATION
                        row.Item("Description") = "Coupon does not pass Location Validation."
                        dtStatus.Rows.Add(row)
                        dtStatus.AcceptChanges()
                        ResultSet.Tables.Add(dtStatus.Copy())
                    End If
                Else
                    'Coupon not valid for this member
                    If MyCommon.LXSsp.Parameters("@ValidMember").Value = 2 Then
                        RetCode = CouponStatusCodes.INVALID_MEMBER
                        RetMessage = "Coupon does not pass Member Validation as MemberRedemptionID is not provided"
                    ElseIf MyCommon.LXSsp.Parameters("@ValidMember").Value = 3 Then
                        RetCode = CouponStatusCodes.INVALID_MEMBER
                        RetMessage = "Coupon does not pass Member Validation as the provided customer has not printed the coupon"
                    ElseIf MyCommon.LXSsp.Parameters("@ValidMember").Value = 4 Then
                        RetCode = CouponStatusCodes.INVALID_MEMBER
                        RetMessage = "Coupon does not pass Member Validation as the customer's card is not active"
                    End If
                    row = dtStatus.NewRow()
                    row.Item("StatusCode") = RetCode
                    row.Item("Description") = RetMessage
                    dtStatus.Rows.Add(row)
                    dtStatus.AcceptChanges()
                    ResultSet.Tables.Add(dtStatus.Copy())
                End If
            End If
        Catch ex As Exception
            row = dtStatus.NewRow()
            row.Item("StatusCode") = CouponStatusCodes.APPLICATION_EXCEPTION
            row.Item("Description") = "Failure: Application " & ex.ToString
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
        Finally
            'Report Status errors
            If (ResultSet.Tables.Item("ValidationStatus").Rows(0).Item("StatusCode") IsNot Nothing) AndAlso _
               (ResultSet.Tables.Item("ValidationStatus").Rows(0).Item("StatusCode") <> 16) Then
                MyCommon.QueryStr = Nothing
                MyCommon.QueryStr = "dbo.pt_CouponError_Insert"
                MyCommon.Open_LWHsp()
                MyCommon.LWHsp.Parameters.Add("@Barcode", SqlDbType.NVarChar, 14).Value = Coupon
                MyCommon.LWHsp.Parameters.Add("@Description", SqlDbType.NVarChar, 255).Value = ResultSet.Tables.Item("ValidationStatus").Rows(0).Item("Description")
                MyCommon.LWHsp.Parameters.Add("@ErrorCode", SqlDbType.BigInt).Value = ResultSet.Tables.Item("ValidationStatus").Rows(0).Item("StatusCode")
                MyCommon.LWHsp.Parameters.Add("@ErrorDate", SqlDbType.DateTime).Value = Now
                MyCommon.LWHsp.Parameters.Add("@LocationCode", SqlDbType.NVarChar, 20).Value = LocationCode
                MyCommon.LWHsp.Parameters.Add("@TransNum", SqlDbType.NVarChar, 128).Value = TransactionID.ToString 'A Transaction ID of 0 means this validation occurred independent of a transaction
                MyCommon.LWHsp.ExecuteNonQuery()
                MyCommon.Close_LWHsp()
            End If

            If MyCommon.LRTadoConn.State <> ConnectionState.Closed AndAlso bOpenedRT Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed AndAlso bOpenedXS Then MyCommon.Close_LogixXS()
            If MyCommon.LWHadoConn.State <> ConnectionState.Closed AndAlso bOpenedWH Then MyCommon.Close_LogixWH()
            MyCommon.Close_LWHsp()
            MyCommon.Close_LRTsp()
            MyCommon.Close_LXSsp()
        End Try

        Return ResultSet
    End Function

    'Redeem a coupon
    <WebMethod()> _
    Public Function RedeemCoupon(ByVal LocationCode As String, ByVal CustomerID As String, _
                                  ByVal CustomerTypeID As Integer, ByVal Coupon As String, _
                                  ByVal TransactionID As String, ByVal CurrentTime As String) As System.Data.DataSet
        Dim dt As System.Data.DataTable = New System.Data.DataTable()
        Dim dt2 As System.Data.DataTable = New System.Data.DataTable()
        Dim dtCouponRedemption As System.Data.DataTable = New System.Data.DataTable()
        Dim dtStatus As System.Data.DataTable = New System.Data.DataTable()
        Dim row As System.Data.DataRow = Nothing
        Dim drRedemption As System.Data.DataRow = Nothing
        Dim bOpenedXS, bOpenedRT, bOpenedWH As Boolean
        Dim ResultSet As New System.Data.DataSet("CouponRedemption")
        Dim ResultSet2 As New System.Data.DataSet("ValidateCoupon")
        Dim ResultSet3 As New System.Data.DataSet("Combined")
        Dim CurrentTimeVar As DateTime
        Dim ErrorMessage As String = ""

        If CurrentTime = "" Then
            CurrentTimeVar = Now
        Else
            Try
                CurrentTimeVar = Date.ParseExact(CurrentTime, "yyyy-MM-ddTHH:mm:ss", Nothing) 'System.Globalization.CultureInfo.InvariantCulture)
            Catch ex As Exception
                CurrentTimeVar = Now
            End Try
        End If

        'Initialize the status table, which will report the success or failure of the CustomerDetails operation
        dtStatus.TableName = "RedemptionStatus"
        dtStatus.Columns.Add("StatusCode2", System.Type.GetType("System.Int32"))
        dtStatus.Columns.Add("Description", System.Type.GetType("System.String"))

        Try
            'Initialize connections
            bOpenedRT = False
            bOpenedXS = False
            bOpenedWH = False
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then
                bOpenedRT = True
                MyCommon.Open_LogixRT()
            End If
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then
                bOpenedXS = True
                MyCommon.Open_LogixXS()
            End If
            If MyCommon.LWHadoConn.State = ConnectionState.Closed Then
                bOpenedWH = True
                MyCommon.Open_LogixWH()
            End If

            MyCommon.QueryStr = "select locationid from locations where ExtLocationCode = @LocationCode;"
            MyCommon.DBParameters.Add("@LocationCode", SqlDbType.NVarChar, 20).Value = LocationCode
            dt = MyCommon.ExecuteQuery(DataBases.LogixRT)

            MyCommon.QueryStr = "select Barcode from BarcodeDetails where Barcode = @Coupon;"
            MyCommon.DBParameters.Add("@Coupon", SqlDbType.NVarChar, 14).Value = Coupon
            dt2 = MyCommon.ExecuteQuery(DataBases.LogixXS)


            CustomerID = transformCard(CustomerID, CustomerTypeID.ToString(), MyCommon, ErrorMessage)
            If (Not String.IsNullOrEmpty(ErrorMessage)) Then
                row = dtStatus.NewRow()
                row.Item("StatusCode2") = CouponStatusCodes.INVALID_CUSTOMERID
                row.Item("Description") = ErrorMessage
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
            ElseIf dt.Rows.Count = 0 Then
                'ExtLocationCode not found
                row = dtStatus.NewRow()
                row.Item("StatusCode2") = CouponStatusCodes.INVALID_LOCATION
                row.Item("Description") = "Failure: Invalid ExtLocationCode."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
            ElseIf dt2.Rows.Count = 0 Then
                'Coupon not found
                row = dtStatus.NewRow()
                row.Item("StatusCode2") = CouponStatusCodes.INVALID_COUPON
                row.Item("Description") = "Failure: Coupon Number not found."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()
                ResultSet.Tables.Add(dtStatus.Copy())
            Else
                'Perform a validation of this coupon so that we can log if anything is wrong with the coupon being redeemed
                ResultSet2 = ValidateCoupon(LocationCode, CustomerID, CustomerTypeID, Coupon, TransactionID, CurrentTimeVar.ToString("yyyy-MM-ddTHH:mm:ss"))

                MyCommon.QueryStr = "dbo.pt_RedeemCoupon"
                MyCommon.Open_LogixXS()
                MyCommon.Open_LXSsp()
                MyCommon.LXSsp.Parameters.Add("@UniqueBarCode", SqlDbType.NVarChar, 12).Value = Coupon
                MyCommon.LXSsp.Parameters.Add("@TransEndDate", SqlDbType.DateTime).Value = CurrentTimeVar
                MyCommon.LXSsp.Parameters.Add("@ExtCardID", SqlDbType.NVarChar, 400).Value = MyCryptLib.SQL_StringEncrypt(CustomerID, True)
                MyCommon.LXSsp.Parameters.Add("@CardTypeID", SqlDbType.Int).Value = CustomerTypeID
                MyCommon.LXSsp.Parameters.Add("@LocationID", SqlDbType.NVarChar, 20).Value = LocationCode 'MyCommon.NZ(dt.Rows(0).Item("LocationID"), 0)
                MyCommon.LXSsp.Parameters.Add("@TransactionID", SqlDbType.NVarChar, 128).Value = TransactionID
                MyCommon.LXSsp.Parameters.Add("@StatusCode", SqlDbType.Int).Direction = ParameterDirection.Output
                MyCommon.LXSsp.ExecuteNonQuery()
                MyCommon.Close_LogixXS()

                'Coupon Redemption attempted for this location and coupon
                row = dtStatus.NewRow()
                row.Item("StatusCode2") = CouponStatusCodes.REDEMPTION_ATTEMPTED
                row.Item("Description") = "Success."
                dtStatus.Rows.Add(row)
                dtStatus.AcceptChanges()

                'Record whether or not the coupon passed Redemption
                dtCouponRedemption.TableName = "CouponRedemption"
                dtCouponRedemption.Columns.Add("StatusCode", System.Type.GetType("System.Int32"))

                drRedemption = dtCouponRedemption.NewRow()
                drRedemption.Item("StatusCode") = MyCommon.NZ(MyCommon.LXSsp.Parameters("@StatusCode").Value, 0)
                dtCouponRedemption.Rows.Add(drRedemption)
                dtCouponRedemption.AcceptChanges()

                MyCommon.QueryStr = Nothing

                ResultSet.Tables.Add(dtStatus.Copy())
                ResultSet.Tables.Add(dtCouponRedemption.Copy())
                'Report Coupon Redemption Errors
                If (ResultSet.Tables.Item("CouponRedemption").Rows(0).Item("StatusCode") IsNot Nothing) AndAlso _
                   (ResultSet.Tables.Item("CouponRedemption").Rows(0).Item("StatusCode") <> 1) Then
                    MyCommon.QueryStr = Nothing
                    MyCommon.QueryStr = "dbo.pt_CouponError_Insert"
                    MyCommon.Open_LWHsp()
                    MyCommon.LWHsp.Parameters.Add("@Barcode", SqlDbType.NVarChar, 14).Value = Coupon
                    MyCommon.LWHsp.Parameters.Add("@Description", SqlDbType.NVarChar, 255).Value = "Entered Coupon has already been redeemed or cannot be found"
                    MyCommon.LWHsp.Parameters.Add("@ErrorCode", SqlDbType.BigInt).Value = ResultSet.Tables.Item("CouponRedemption").Rows(0).Item("StatusCode")
                    MyCommon.LWHsp.Parameters.Add("@ErrorDate", SqlDbType.DateTime).Value = Now
                    MyCommon.LWHsp.Parameters.Add("@LocationCode", SqlDbType.NVarChar, 20).Value = LocationCode
                    MyCommon.LWHsp.Parameters.Add("@TransNum", SqlDbType.NVarChar, 128).Value = TransactionID.ToString
                    MyCommon.LWHsp.ExecuteNonQuery()
                    MyCommon.Close_LWHsp()
                End If
            End If

        Catch ex As Exception
            row = dtStatus.NewRow()
            row.Item("StatusCode2") = CouponStatusCodes.APPLICATION_EXCEPTION
            row.Item("Description") = "Failure: Application " & ex.ToString
            dtStatus.Rows.Add(row)
            dtStatus.AcceptChanges()
            ResultSet.Tables.Add(dtStatus.Copy())
        Finally
            'Report Status errors
            If (ResultSet.Tables.Item("RedemptionStatus").Rows(0).Item("StatusCode2") IsNot Nothing) AndAlso _
              (ResultSet.Tables.Item("RedemptionStatus").Rows(0).Item("StatusCode2") <> 17) Then
                MyCommon.QueryStr = Nothing
                MyCommon.QueryStr = "dbo.pt_CouponError_Insert"
                MyCommon.Open_LWHsp()
                MyCommon.LWHsp.Parameters.Add("@Barcode", SqlDbType.NVarChar, 14).Value = Coupon
                MyCommon.LWHsp.Parameters.Add("@Description", SqlDbType.NVarChar, 255).Value = ResultSet.Tables.Item("RedemptionStatus").Rows(0).Item("Description")
                MyCommon.LWHsp.Parameters.Add("@ErrorCode", SqlDbType.BigInt).Value = ResultSet.Tables.Item("RedemptionStatus").Rows(0).Item("StatusCode2")
                MyCommon.LWHsp.Parameters.Add("@ErrorDate", SqlDbType.DateTime).Value = Now
                MyCommon.LWHsp.Parameters.Add("@LocationCode", SqlDbType.NVarChar, 20).Value = LocationCode
                MyCommon.LWHsp.Parameters.Add("@TransNum", SqlDbType.NVarChar, 128).Value = TransactionID.ToString
                MyCommon.LWHsp.ExecuteNonQuery()
                MyCommon.Close_LWHsp()
            End If

            If MyCommon.LRTadoConn.State <> ConnectionState.Closed AndAlso bOpenedRT Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed AndAlso bOpenedXS Then MyCommon.Close_LogixXS()
            If MyCommon.LWHadoConn.State <> ConnectionState.Closed AndAlso bOpenedWH Then MyCommon.Close_LogixWH()
            MyCommon.Close_LWHsp()
            MyCommon.Close_LRTsp()
            MyCommon.Close_LXSsp()
        End Try

        'Return the appropriate tables
        If ResultSet2 IsNot Nothing Then
            ResultSet3.Merge(ResultSet, True, System.Data.MissingSchemaAction.Add)
            ResultSet3.Merge(ResultSet2, True, System.Data.MissingSchemaAction.Add)
            Return ResultSet3
        Else
            Return ResultSet
        End If

    End Function

End Class

Public Class OfferCategory
  Public CategoryID As Integer
  Public DisplayText As String
  Public SortOrder As Integer
  Public IconFileName As String

  Public Sub New()
    CategoryID = 0
    DisplayText = ""
    SortOrder = 0
    IconFileName = ""
  End Sub

  Public Sub New(ByVal NewCategoryID As Integer, ByVal NewDisplayText As String, ByVal NewSortOrder As Integer, ByVal NewIconFileName As String)
    CategoryID = NewCategoryID
    DisplayText = NewDisplayText
    SortOrder = NewSortOrder
    IconFileName = NewIconFileName
  End Sub

End Class

Public Class RedemptionItemCategories

  Public Categories() As OfferCategory

End Class

Public Class RedemptionItem
  Public SamID As Integer
  Public Name As String
  Public RedemptionText As String
  Public PointsValue As String
  Public DateCreated As Nullable(Of DateTime)
  Public LastUpdate As Nullable(Of DateTime)
  Public StartDate As Nullable(Of DateTime)
  Public EndDate As Nullable(Of DateTime)
  Public CategoryID As Integer
  Public ImageName As String
  Public MessageText As String
  Public ProgramID As String

  Public Sub New()
    SamID = 0
    Name = ""
    RedemptionText = ""
    PointsValue = 0
    DateCreated = Nothing
    LastUpdate = Nothing
    StartDate = Nothing
    EndDate = Nothing
    CategoryID = 0
    ImageName = ""
    MessageText = ""
    ProgramID = 0
  End Sub

  Public Sub LoadOffer(ByVal NewSamID As Integer, ByVal NewName As String, ByVal NewDateCreated As DateTime, ByVal NewLastUpdate As DateTime, ByVal NewStartDate As DateTime, ByVal NewEndDate As DateTime, ByVal NewCategoryID As Integer)
    SamID = NewSamID
    Name = NewName
    DateCreated = NewDateCreated
    LastUpdate = NewLastUpdate
    StartDate = NewStartDate
    EndDate = NewEndDate
    CategoryID = NewCategoryID
  End Sub

  Public Sub LoadPassThru(ByVal NewRedemptionText As String, ByVal NewPointsValue As String, ByVal NewImageName As String, ByVal NewMessageText As String, ByVal NewProgramID As String)
    RedemptionText = NewRedemptionText
    PointsValue = NewPointsValue
    ImageName = NewImageName
    MessageText = NewMessageText
    ProgramID = NewProgramID
  End Sub

End Class

Public Class RedemptionItemsAvailable
  Public Items() As RedemptionItem
End Class

Public Class Excluded
  Public Message As String
  Public Success As Boolean

  Public Sub New()
    Message = ""
    Success = False
  End Sub
End Class

<Serializable()> _
Public Class RedemptionItemsAvailableException
  Inherits ApplicationException
  Public Sub New()
  End Sub

  Public Sub New(ByVal message As String)
    MyBase.New(message)
  End Sub

  Protected Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, _
    ByVal context As System.Runtime.Serialization.StreamingContext)
    MyBase.New(info, context)
  End Sub

End Class

Public Class CostCenterException
  Inherits ArgumentException
  Public Sub New(ByVal message As String)
    MyBase.New(message)
  End Sub
End Class

Public Class TerminalIDException
  Inherits ArgumentException
  Public Sub New(ByVal message As String)
    MyBase.New(message)
  End Sub
End Class