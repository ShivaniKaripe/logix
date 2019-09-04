<%-- version:7.3.1.138972 --%>
<%@ WebService Language="VB" Class="CouponRedemption" %>

Imports System
Imports System.Web
Imports System.Web.Services
Imports System.Xml.Serialization
Imports System.Xml.Schema
Imports System.Xml
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Collections.Generic
Imports System.Collections.Specialized.NameValueCollection
Imports System.Threading.Thread

<WebService(Namespace:="http://www.copienttech.com/CouponRedemption/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class CouponRedemption
  Inherits System.Web.Services.WebService

    Private MyCommon As New Copient.CommonInc
    Private MyCryptLib As New Copient.CryptLib
  'Private MyLog As New Copient.Log4Net 
  'Private WSthread As Integer = System.Threading.Thread.CurrentThread.ManagedThreadId
    
  Private LogFile As String = "CouponRedemptionWSLog." & Date.Now.ToString("yyyyMMdd") & ".txt"
  Private Const CONNECTOR_ID As Integer = 59

  Private Sub InitApp()
    MyCommon.AppName = "CouponRedemption.asmx"

    Try
    Catch eXmlSch As XmlSchemaException
    Catch ex As Exception

    End Try
  End Sub

  Private Function IsValidGUID(ByVal GUID As String, ByVal MethodName As String) As Boolean
    Dim IsValid As Boolean = False
    Dim ConnInc As New Copient.ConnectorInc
    Dim MsgBuf As New StringBuilder()
    Try
      Using MyCommon.LRTadoConn
        If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
        IsValid = ConnInc.IsValidConnectorGUID(MyCommon, CONNECTOR_ID, GUID)
      End Using
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
      Copient.Logger.Write_Log(LogFile, MsgBuf.ToString, True)
    Catch ex As Exception
      ' ignore
    End Try

    Return IsValid
  End Function
	
  Function buildCouponList(ByVal barcodes As DataTable, ByVal ExtCardID As String) As DataSet
	
    Dim CouponList As New System.Data.DataSet("ActiveCoupons")
    Dim CouponDST As DataTable = New DataTable("Coupon")
    Dim CouponInfoDST As DataTable
		
    If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
    CouponDST.Columns.Add("Type", Type.GetType("System.String"))
    CouponDST.Columns.Add("Reward", Type.GetType("System.String"))
    CouponDST.Columns.Add("Desc", Type.GetType("System.String"))
    CouponDST.Columns.Add("Expires", Type.GetType("System.DateTime"))
    CouponDST.Columns.Add("CsrInfo", Type.GetType("System.String"))
    CouponDST.Columns.Add("BarCode", Type.GetType("System.String"))
    CouponDST.Columns.Add("Effective", Type.GetType("System.DateTime"))
    CouponDST.Columns.Add("TriggerUPC", Type.GetType("System.String"))
    CouponDST.Columns.Add("Restriction", Type.GetType("System.String"))
    CouponDST.Columns.Add("RedemptionText", Type.GetType("System.String"))
    CouponDST.Columns.Add("EarnedLocation", Type.GetType("System.String"))
    CouponDST.Columns.Add("EarningMember", Type.GetType("System.String"))
    CouponDST.Columns.Add("SVProgramID", Type.GetType("System.String"))
    CouponDST.Columns.Add("ImageName", Type.GetType("System.String"))
		
    For Each BarcodeRow In barcodes.Rows
      If BarcodeRow.Item("RewardOptionID") Is DBNull.Value Then
        MyCommon.Write_Log(LogFile, "No offer Connected with this barcode: " & BarcodeRow.Item("Barcode"), True)
      Else
        MyCommon.QueryStr = "select CouponType.value 'Type',  " & _
          "RewardName.Value 'Reward', " & _
          "Descrip.Value 'Desc', " & _
          "CSRInfo.Value 'CSRInfo', " & _
          "TriggerUPC.Value 'TriggerUPC', " & _
          "Restriction.Value 'Restriction', " & _
          "RedemptionText.Value 'RedemptionText', " & _
          "ImageName.Value 'ImageName' " & _
         "from CPE_Deliverables as D " & _
         "inner join PassThrus as PT  " & _
          "on D.DeliverableID = PT.DeliverableID  " & _
         "inner join PassThruTierValues as CouponType  " & _
          "on CouponType.PTPKID = PT.PKID " & _
         "inner join PassthruTierValues as RewardName  " & _
          "on CouponType.PTPKID = RewardName.PTPKID " & _
         "inner join PassThruTierValues as Descrip " & _
          "on Descrip.PTPKID = CouponType.PTPKID " & _
         "inner join PassThruTierValues as CSRInfo " & _
          "on CSRInfo.PTPKID = CouponType.PTPKID " & _
         "inner join PassThruTierValues as TriggerUPC " & _
          "on TriggerUPC.PTPKID = CouponType.PTPKID " & _
         "inner join PassThruTierValues as Restriction " & _
          "on Restriction.PTPKID = CouponType.PTPKID " & _
         "inner join PassThruTierValues as RedemptionText " & _
          "on RedemptionText.PTPKID = CouponType.PTPKID " & _
         "inner join PassThruTierValues as ImageName " & _
          "on ImageName.PTPKID = CouponType.PTPKID " & _
         "where CouponType.PassThruPresTagID = 20 and  " & _
          "RewardName.PassThruPresTagID = 11 and  " & _
          "Descrip.PassThruPresTagID = 12 and  " & _
          "CSRInfo.PassThruPresTagID = 17 and " & _
          "TriggerUPC.PassThruPresTagID = 15 and " & _
          "Restriction.PassThruPresTagID = 18 and   " & _
          "RedemptionText.PassThruPresTagID = 19 and   " & _
          "ImageName.PassThruPresTagID = 23 and   " & _
          "PT.PassThruRewardID = 4 and  " & _
          "D.RewardOptionID = " & MyCommon.NZ(BarcodeRow.Item("RewardOptionID"), -1)

			
        CouponInfoDST = MyCommon.LRT_Select
				
        If CouponInfoDST.Rows.Count > 0 Then
          For Each CouponRow In CouponInfoDST.Rows
            CouponDST.Rows.Add(CouponRow.Item("Type"), CouponRow.Item("Reward"), CouponRow.Item("Desc"), BarcodeRow.Item("ExpirationDate"), CouponRow.Item("CSRInfo"),
                  BarcodeRow.Item("barcode"), BarcodeRow.Item("EffectiveDate"), CouponRow.Item("TriggerUPC"), CouponRow.Item("Restriction"), CouponRow.Item("RedemptionText"),
                  BarcodeRow.Item("IssuingCostCenter"), ExtCardID, BarcodeRow.Item("SVProgramID"), CouponRow.Item("ImageName"))

          Next
        End If
      End If

    Next
    CouponList.Tables.Add(CouponDST)
    Return CouponList
  End Function
		
  <WebMethod()> _
  Function GetMemberCoupons(ByVal GUID As String, ByVal ExtCardID As String, ByVal CardTypeID As String, ByVal CurrentDate As String) As DataSet
    Dim BarcodeDST As DataTable
    Dim CouponList As DataSet
    Dim CardTypeIDint As Integer
    Dim tempDate As DateTime
    Dim CustomerPK As Int64
		
    If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
    If IsValidGUID(GUID, "GetMemberCoupons") Then
      If DateTime.TryParse(CurrentDate, tempDate) = False Then
        Copient.Logger.Write_Log(LogFile, "CurrentDate could not be parsed", True)
        Return GenerateErrorXML("CurrentDate could not be parsed")
      End If
      If ExtCardID = "" Then
        Copient.Logger.Write_Log(LogFile, "Error. ExtCardID is blank", True)
        Return GenerateErrorXML("Error. ExtCardID is blank")
      End If
      If CardTypeID = "" Or Not Integer.TryParse(CardTypeID, CardTypeIDint) Then
        Copient.Logger.Write_Log(LogFile, "Error. Invalid CardType", True)
        Return GenerateErrorXML("Error. Invalid CardType")
      End If
			
      Dim ErrorMessage As String = ""
      ExtCardID = transformCard(ExtCardID, CardTypeIDint.ToString(), MyCommon, ErrorMessage)
      If (Not String.IsNullOrEmpty(ErrorMessage)) Then
        Copient.Logger.Write_Log(LogFile, ErrorMessage, True)
        Return GenerateErrorXML(ErrorMessage)
      End If
      ExtCardID = MyCommon.Pad_ExtCardID(ExtCardID, CardTypeIDint)
            MyCommon.QueryStr = "select top 1 CustomerPK from CardIDs where ExtCardID = '" & MyCryptLib.SQL_StringEncrypt(ExtCardID, True) & "' and CardTypeID = " & CardTypeIDint
      BarcodeDST = MyCommon.LXS_Select
      If BarcodeDST.Rows.Count > 0 Then
        CustomerPK = BarcodeDST.Rows(0).Item("CustomerPK")
      Else
        CustomerPK = -1
      End If
			
      MyCommon.QueryStr = "select  Barcode, CustomerPK, SVProgramID, ExpirationDate, EffectiveDate, RewardOptionID,IssuingCostCenter from BarcodeDetails where  ExpirationDate > '" & tempDate.ToString("s") & "' and RedeemedDate is null  and CustomerPK = " & CustomerPK
      BarcodeDST = MyCommon.LXS_Select

      If BarcodeDST.Rows.Count > 0 Then
        CouponList = buildCouponList(BarcodeDST, ExtCardID)
      Else
        Copient.Logger.Write_Log(LogFile, "Could not find any active barcodes for customer: " & ExtCardID & " with card type: " & CardTypeIDint, True)
        Return GenerateErrorXML("Could not find any active barcodes for customer: " & ExtCardID & " with card type: " & CardTypeIDint)
      End If
      Return CouponList
    Else
      Copient.Logger.Write_Log(LogFile, "Could not validate GUID", True)
      Return GenerateErrorXML("Could not validate GUID")
    End If
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
        ErrorMessage += ex.Message & vbCrLf
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
  Function GetCouponRedemptionHistory(ByVal GUID As String, ByVal ExtCardID As String, ByVal CardTypeID As String, ByVal StartDate As String, ByVal EndDate As String) As DataSet
    Dim ExpiredDST As DataTable
    Dim RedeemedDST As DataTable
    Dim CouponList As DataSet
    Dim CardTypeIDint As Integer
    Dim tempStartDate As DateTime
    Dim tempEndDate As DateTime
    If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
    If IsValidGUID(GUID, "GetCouponRedemptionHistory") Then
      If DateTime.TryParse(StartDate, tempStartDate) = False Then
        Copient.Logger.Write_Log(LogFile, "Start Date could not be parsed", True)
        Return GenerateErrorXML("Start Date could not be parsed")
      End If
      If DateTime.TryParse(EndDate, tempEndDate) = False Then
        Copient.Logger.Write_Log(LogFile, "End Date could not be parsed", True)
        Return GenerateErrorXML("End Date could not be parsed")
      End If
      If ExtCardID = "" Then
        Copient.Logger.Write_Log(LogFile, "Error. ExtCardID is blank", True)
        Return GenerateErrorXML("Error. ExtCardID is blank")
      End If
      If CardTypeID = "" Or Not Integer.TryParse(CardTypeID, CardTypeIDint) Then
        Copient.Logger.Write_Log(LogFile, "Error. Invalid CardType", True)
        Return GenerateErrorXML("Error. Invalid CardType")
      End If
      If tempStartDate > tempEndDate Then
        Copient.Logger.Write_Log(LogFile, "Start Date must be before End Date", True)
        Return GenerateErrorXML("Start Date must be before End Date")
      End If
			
      Dim ErrorMessage As String = ""
      ExtCardID = transformCard(ExtCardID, CardTypeIDint.ToString(), MyCommon, ErrorMessage)
      If (Not String.IsNullOrEmpty(ErrorMessage)) Then
        Copient.Logger.Write_Log(LogFile, ErrorMessage, True)
        Return GenerateErrorXML(ErrorMessage)
      End If
            
      ExtCardID = MyCommon.Pad_ExtCardID(ExtCardID, CardTypeIDint)
			
            MyCommon.QueryStr = "select Barcode, ExpirationDate, RewardOptionID, RedeemedDate, RedeemedLocationID, RedeemingTransactionID from BarcodeDetails where" & _
                     "((RedeemedDate<'" & tempEndDate.ToString("s") & "' and RedeemedDate >'" & tempStartDate.ToString("s") & "') OR " & _
                     "(RedeemedLocationID is NULL and ExpirationDate< '" & tempEndDate.ToString("s") & "' and ExpirationDate >'" & tempStartDate.ToString("s") & "'))" & _
                     "and CustomerPK = (select top 1 CustomerPK from CardIDs where ExtCardID = '" & MyCryptLib.SQL_StringEncrypt(ExtCardID, True) & "' and CardTypeID = " & CardTypeIDint & ")"
      ExpiredDST = MyCommon.LXS_Select


      If ExpiredDST.Rows.Count > 0 Then
        CouponList = BuildCouponHistoryList(ExpiredDST, EndDate)
      Else
        MyCommon.Write_Log(LogFile, "Could not find any expired or redeemed barcodes for customer: " & ExtCardID & " with card type: " & CardTypeIDint, True)
        Return GenerateErrorXML("Could not find any expired or redeemed barcodes for customer: " & ExtCardID & " with card type: " & CardTypeIDint)
      End If
			
      Return CouponList
    Else
      Copient.Logger.Write_Log(LogFile, "Could not validate GUID", True)

      Return GenerateErrorXML("Could not validate GUID")
    End If
  End Function
	
  Function BuildCouponHistoryList(ByVal barcodes As DataTable, ByVal EndDate As String) As DataSet
    Dim CouponList As New System.Data.DataSet("CouponHistory")
    Dim CouponDST As DataTable = New DataTable("Coupon")
    Dim CouponInfoDST As DataTable
    Dim dst As DataTable
    Dim RewardName As String
		
		
    If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
    CouponDST.Columns.Add("Barcode", Type.GetType("System.String"))
    CouponDST.Columns.Add("RewardName", Type.GetType("System.String"))
    CouponDST.Columns.Add("ExpirationDate", Type.GetType("System.DateTime"))
    CouponDST.Columns.Add("RedeemedDate", Type.GetType("System.DateTime"))
    CouponDST.Columns.Add("RedeemedLocationID", Type.GetType("System.String"))
    CouponDST.Columns.Add("RedeemingTransactionID", Type.GetType("System.String"))
		
    For Each barcodeRow In barcodes.Rows
      If barcodeRow.Item("RewardOptionID") Is DBNull.Value Then
        MyCommon.Write_Log(LogFile, "No offer Connected with this barcode: " & barcodeRow.Item("Barcode"), True)
      Else
        MyCommon.QueryStr = "select RewardName.Value as 'RewardName' from CPE_Deliverables as D " & _
                 "inner join PassThrus as PT on D.DeliverableID = PT.DeliverableID " & _
                 "inner join PassthruTierValues as RewardName on RewardName.PTPKID =PT.PKID " & _
                 "where RewardName.PassThruPresTagID = 11 and " & _
                 "PT.PassThruRewardID = 4 and 	D.RewardOptionID = " & barcodeRow.Item("RewardOptionID")
        dst = MyCommon.LRT_Select
        For Each row In dst.Rows
          If dst.Rows.Count > 0 Then
            RewardName = row.Item("RewardName")

          Else
            RewardName = "NULL"
          End If

          If barcodeRow.Item("RedeemedDate") Is DBNull.Value Then
            CouponDST.Rows.Add(barcodeRow.item("Barcode"), RewardName, barcodeRow.Item("ExpirationDate"), DBNull.Value, DBNull.Value, DBNull.Value)
          Else
            CouponDST.Rows.Add(barcodeRow.Item("Barcode"), RewardName, DBNull.Value, MyCommon.NZ(barcodeRow.Item("RedeemedDate"), DBNull.Value), MyCommon.NZ(barcodeRow.Item("RedeemedLocationID"), DBNull.Value),
                    MyCommon.NZ(barcodeRow.Item("RedeemingTransactionID"), DBNull.Value))
          End If

        Next
      End If
    Next
    CouponList.Tables.Add(CouponDST)
    Return CouponList
  End Function
	
  Function GenerateErrorXML(ByVal errorString As String) As DataSet
    Dim ErrorXML As New System.Data.DataSet("Error")
    Dim errorTable As DataTable = New DataTable("Error")
    errorTable.Columns.Add("ErrorMessage", Type.GetType("System.String"))

    errorTable.Rows.Add(errorString)
    ErrorXML.Tables.Add(errorTable)
    Return ErrorXML
  End Function

  Protected Overrides Sub Finalize()
    MyBase.Finalize()
  End Sub
    
End Class

