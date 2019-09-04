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
Imports Copient.CryptLib

<WebService(Namespace:="http://www.copienttech.com/KioskWeb/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class Service
    Inherits System.Web.Services.WebService
    ' version:7.3.1.138972.Official Build (SUSDAY10202)
  
    Private MyCommon As New Copient.CommonInc
    Private MyAltID As New Copient.AlternateID
    Private MyCryptLib As New Copient.CryptLib
  
    Public Enum StatusCodes As Integer
        SUCCESS = 0
        INVALID_GUID = 1
        INVALID_LOCATION = 2
        INVALID_CARD = 3
        ALT_ID_NOT_FOUND = 4
        BANNER_ID_NOT_FOUND = 5
        APPLICATION_EXCEPTION = 9999
    End Enum
  
  
    <WebMethod()> _
    Public Function GetAccumulationBalance(ByVal GUID As String, ByVal CardID As String, ByVal CardTypeID As String, ByVal AccumulationType As Integer) As XmlDocument
        Dim RetXmlDoc As New XmlDocument
        Dim ms As New MemoryStream
        Dim Writer As New XmlTextWriter(ms, System.Text.Encoding.UTF8)
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = ""
        Dim XsdType As Integer = 0
        Dim BatchGUID As String = System.Guid.NewGuid.ToString
        Dim dt As DataTable
        Dim dtCGIDs As DataTable
        Dim dtOffers As DataTable
        Dim dtBals As DataTable
        Dim row As DataRow
        Dim row2 As DataRow
        Dim attrib As XmlAttribute = Nothing
        Dim CGIDs As String = "-1"
        Dim CustomerPK As Long = 0
        Dim HHPK As Long = 0
        Dim CustomerGroupIDs As String = ""
        Dim cgXml As String = ""
        Dim HHEnable As Boolean = False
        Dim HHCustomerPKs As String = "-1"
        Dim IDLength As Integer = 0
        Dim i As Integer = 0
        Dim j As Integer = 0
        'added
        Dim iCardTypeID As Integer = -1
        'If Not String.IsNullOrEmpty(CardTypeID) AndAlso Not CardTypeID.Trim = String.Empty AndAlso IsNumeric(CardTypeID) Then iCardTypeID = Convert.ToInt32(CardTypeID)
       
    
        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
      
            'Validate the request
            If Not IsValidGUID(GUID) Then
                RetCode = StatusCodes.INVALID_GUID
                RetMsg = "GUID " & GUID & " is not valid for the KioskWeb web service."
            End If
      
            'Pad, then validate card
            'added
            If IsValidCustomerCard(CardID, CardTypeID, RetCode, RetMsg) Then
                If Not String.IsNullOrEmpty(CardTypeID) AndAlso Not CardTypeID.Trim = String.Empty AndAlso IsNumeric(CardTypeID) Then iCardTypeID = Convert.ToInt32(CardTypeID)
                If CardTypeID = "0" Or CardTypeID = "1" Then
                    CardID = MyCommon.Pad_ExtCardID(CardID, Convert.ToInt32(CardTypeID))
                End If
            End If
          
            MyCommon.QueryStr = "select CustomerPK from CardIDs with (NoLock) where ExtCardID='" & MyCryptLib.SQL_StringEncrypt(CardID, True) & "' and CardTypeID=" & iCardTypeID & ";"
            dt = MyCommon.LXS_Select
            If dt.Rows.Count = 0 Then
                If RetCode = StatusCodes.SUCCESS AndAlso RetMsg = "" Then
                    'RetCode = StatusCodes.INVALID_CARD
                    'RetMsg = "Card " & CardID & " of type " & CardTypeID & " not found in Logix."
                    RetCode = StatusCodes.INVALID_CARD
                    RetMsg = "Card " & CardID & " of type " & CardTypeID & " not found in Logix."
                Else
                    RetCode = RetCode
                    RetMsg = RetMsg
                End If
            Else
                CustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)
            End If
      
            If iCardTypeID = 1 Then
                'Card is a household, so find all the customers within it.
                MyCommon.QueryStr = "select CustomerPK from Customers with (NoLock) where HHPK=" & CustomerPK & ";"
                dt = MyCommon.LXS_Select
                If (dt.Rows.Count > 0) Then
                    HHCustomerPKs = ""
                    For Each row In dt.Rows
                        j += 1
                        HHCustomerPKs &= MyCommon.NZ(row.Item("CustomerPK"), 0) & IIf(j < dt.Rows.Count, ",", "")
                    Next
                End If
            ElseIf iCardTypeID = 0 Then
                'Card is a customer, then get its HHPK (if any)
                MyCommon.QueryStr = "select HHPK from Customers with (NoLock) where CustomerPK=" & CustomerPK & ";"
                dt = MyCommon.LXS_Select
                If (dt.Rows.Count > 0) Then
                    HHPK = MyCommon.NZ(dt.Rows(0).Item("HHPK"), 0)
                End If
            End If
      
            If RetCode = StatusCodes.SUCCESS Then
        
                '1) Create a list of the customer groups of which the customer is a member, to be used in next step.
                MyCommon.QueryStr = "select distinct CustomerGroupID from GroupMembership with (NoLock) where CustomerPK in (" & CustomerPK & ", " & HHPK & ") and Deleted=0;"
                dtCGIDs = MyCommon.LXS_Select()
                If dtCGIDs.Rows.Count > 0 Then
                    CGIDs = ""
                    For Each row In dtCGIDs.Rows
                        i += 1
                        CGIDs &= MyCommon.NZ(row.Item("CustomerGroupID"), 0) & IIf(i < dtCGIDs.Rows.Count, ",", "")
                    Next
                End If
        
                '2) Find all associated offer information and put it in a hashtable.
                MyCommon.QueryStr = "select I.IncentiveID, I.IncentiveName, RO.RewardOptionID, RO.HHEnable, IPG.QtyUnitType " & _
                                    "from CPE_Incentives as I with (NoLock) " & _
                                    "inner join CPE_RewardOptions as RO with (NoLock) on RO.IncentiveID=I.IncentiveID " & _
                                    "inner join CPE_IncentiveProductGroups as IPG with (NoLock) on RO.RewardOptionID=IPG.RewardOptionID " & _
                                    "where I.Deleted=0 and I.IsTemplate=0 and I.DisplayOnWebKiosk=1 " & _
                                    "and IsNull(IPG.AccumMin,0) > 0 and DateAdd(d,1,I.EndDate)>=GETDATE() "
                If AccumulationType = 0 Then
                    MyCommon.QueryStr &= "and IPG.QtyUnitType in (1,2) "
                Else
                    MyCommon.QueryStr &= "and IPG.QtyUnitType=" & AccumulationType & " "
                End If
                MyCommon.QueryStr &= "and RO.RewardOptionID in (" & _
                                    "  select distinct RewardOptionID from CPE_IncentiveCustomerGroups where CustomerGroupID in (1,2," & CGIDs & ")" & _
                                    ");"
                dtOffers = MyCommon.LRT_Select

                Writer.Formatting = Formatting.Indented
                Writer.Indentation = 4
        
                Writer.WriteStartDocument()
                Writer.WriteStartElement("Offers")
        
                '3) If there are any offers returned, get the accumulation balance for each one and write it out.
                If dtOffers.Rows.Count > 0 Then
                    For Each row In dtOffers.Rows
                        HHEnable = MyCommon.NZ(dtOffers.Rows(0).Item("HHEnable"), False)
                        If MyCommon.NZ(row.Item("QtyUnitType"), 1) = 1 Then
                            MyCommon.QueryStr = "select " & MyCommon.NZ(row.Item("IncentiveID"), 0) & " as OfferID, '" & MyCommon.NZ(row.Item("IncentiveName"), "") & "' as OfferName, " & _
                                                "isnull(sum(RA.QtyPurchased), 0) as AccumulationBalance from CPE_RewardAccumulation as RA with (NoLock) " & _
                                                "where RewardOptionID=" & MyCommon.NZ(row.Item("RewardOptionID"), 0)
                        Else
                            MyCommon.QueryStr = "select " & MyCommon.NZ(row.Item("IncentiveID"), 0) & " as OfferID, '" & MyCommon.NZ(row.Item("IncentiveName"), "") & "' as OfferName, " & _
                                                "isnull(sum(RA.TotalPrice), 0) as AccumulationBalance from CPE_RewardAccumulation as RA with (NoLock) " & _
                                                "where RewardOptionID=" & MyCommon.NZ(row.Item("RewardOptionID"), 0)
                        End If
                        If iCardTypeID = 1 Then
                            If HHCustomerPKs <> "" Then
                                MyCommon.QueryStr &= " and CustomerPK in (" & CustomerPK & "," & HHCustomerPKs & ");"
                            Else
                                MyCommon.QueryStr &= " and CustomerPK=" & CustomerPK & ";"
                            End If
                        Else
                            If HHPK > 0 AndAlso HHEnable Then
                                MyCommon.QueryStr &= " and CustomerPK in (" & CustomerPK & "," & HHPK & ");"
                            Else
                                MyCommon.QueryStr &= " and CustomerPK=" & CustomerPK & ";"
                            End If
                        End If
                        dtBals = MyCommon.LXS_Select
                        If dtBals.Rows.Count > 0 Then
                            For Each row2 In dtBals.Rows
                                WriteKioskWebAccums(Writer, row2)
                            Next
                        End If
                    Next
                End If
        
                Writer.WriteEndElement() ' end Offers
                Writer.WriteEndDocument()
                Writer.Flush()
        
                ms.Seek(0, SeekOrigin.Begin)
                RetXmlDoc.Load(ms)
        
                attrib = RetXmlDoc.CreateAttribute("returnCode")
                attrib.Value = "SUCCESS"
                attrib = RetXmlDoc.SelectSingleNode("//Offers").Attributes.Append(attrib)
            Else
                RetXmlDoc = New XmlDocument()
                RetXmlDoc.LoadXml(GetErrorXML(RetCode, RetMsg))
            End If

        Catch ex As Exception
            MyCommon.Write_Log("KioskWeb.txt", "Exception: " & ex.ToString, True)
            RetXmlDoc = New XmlDocument()
            RetCode = StatusCodes.APPLICATION_EXCEPTION
            RetMsg = ex.ToString
            RetXmlDoc.LoadXml(GetErrorXML(RetCode, RetMsg))
        Finally
            If Not Writer Is Nothing Then
                Writer.Close()
            End If
            If Not ms Is Nothing Then
                ms.Close()
                ms.Dispose()
                ms = Nothing
            End If
      
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try
    
        Return RetXmlDoc
    End Function
  
    <WebMethod()> _
    Public Function GetWebKioskTargetedOffers(ByVal GUID As String, ByVal CardID As String, ByVal CardTypeID As String, ByVal LocationCode As String) As XmlDocument
        'RT#2629--SSA FIS §2.3.9
        Dim RetXmlDoc As New XmlDocument
        Dim ms As New MemoryStream
        Dim Writer As New XmlTextWriter(ms, System.Text.Encoding.UTF8)
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = ""
        Dim XsdType As Integer = 0
        Dim BatchGUID As String = System.Guid.NewGuid.ToString
        Dim dt As DataTable
        Dim dtOffers As DataTable
        Dim row As DataRow
        Dim attrib As XmlAttribute = Nothing
        Dim LocationID As Long = 0
        Dim CustomerPK As Long = 0
        Dim CustomerGroupIDs As String = ""
        Dim cgXml As String = ""
        Dim rowCount As Integer
        'added
        Dim iCardTypeID As Integer = -1
       
    
        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
      
            ' validate the request
            If Not IsValidGUID(GUID) Then
                RetCode = StatusCodes.INVALID_GUID
                RetMsg = "GUID " & GUID & " is not valid for the KioskWeb web service."
            End If
      
      ' validate location
      If RetCode = StatusCodes.SUCCESS AndAlso RetMsg = "" Then
        LocationID = GetLocationID(LocationCode)
        If LocationID <= 0 Then
          RetCode = StatusCodes.INVALID_LOCATION
          RetMsg = "Location Code " & LocationCode & " not found in Logix."
        End If
      End If
      'validate,then pad card
      If RetCode = StatusCodes.SUCCESS AndAlso RetMsg = "" Then
        If IsValidCustomerCard(CardID, CardTypeID, RetCode, RetMsg) Then
          If Not String.IsNullOrEmpty(CardTypeID) AndAlso Not CardTypeID.Trim = String.Empty AndAlso IsNumeric(CardTypeID) Then iCardTypeID = Convert.ToInt32(CardTypeID)
          If CardTypeID = "0" Or CardTypeID = "1" Then
            CardID = MyCommon.Pad_ExtCardID(CardID, iCardTypeID) 'Convert.ToInt32(CardTypeID)
          End If
        End If
      End If
      ' validate card
            MyCommon.QueryStr = "select CustomerPK from CardIDs with (NoLock) where ExtCardID='" & MyCryptLib.SQL_StringEncrypt(CardID, True) & "' and CardTypeID=" & iCardTypeID & ";"
      dt = MyCommon.LXS_Select
      If dt.Rows.Count = 0 Then
        If RetCode = StatusCodes.SUCCESS AndAlso RetMsg = "" Then
          'RetCode = StatusCodes.INVALID_CARD
          'RetMsg = "Card " & CardID & " of type " & CardTypeID & " not found in Logix."
          RetCode = StatusCodes.INVALID_CARD
          RetMsg = "Card " & CardID & " of type " & CardTypeID & " not found in Logix."
        Else
          RetCode = RetCode
          RetMsg = RetMsg
        End If
      Else
        CustomerPK = MyCommon.NZ(dt.Rows(0).Item("CustomerPK"), 0)
      End If
      
      If RetCode = StatusCodes.SUCCESS Then
        
        'First build an XML list of customer groups of which the customer is a part; this list
        'will be passed into the stored procedure that returns the list of offers.
        MyCommon.QueryStr = "select CustomerGroupID from GroupMembership with (NoLock) where CustomerPK=" & CustomerPK & " and Deleted=0;"
        dt = MyCommon.LXS_Select()
        cgXml = "<customergroups>"
        'cgXml = "<id>1</id><id>2</id>"
        rowCount = dt.Rows.Count
        If rowCount > 0 Then
          For Each row In dt.Rows
            cgXml &= "<id>" & MyCommon.NZ(row.Item("CustomerGroupID"), "") & "</id>"
          Next
        End If
        cgXml &= "</customergroups>"
        
        Writer.Formatting = Formatting.Indented
        Writer.Indentation = 4
        
        Writer.WriteStartDocument()
        Writer.WriteStartElement("Offers")
        
        MyCommon.QueryStr = "dbo.pa_CPE_KioskWeb_OffersByLocationAndCustomer"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = LocationID
        MyCommon.LRTsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = CustomerPK
        MyCommon.LRTsp.Parameters.Add("@cgXml", SqlDbType.Xml).Value = cgXml
        dtOffers = MyCommon.LRTsp_select
        MyCommon.Close_LRTsp()
        
        For Each row In dtOffers.Rows
          WriteKioskWebOffer(Writer, row)
        Next
        
        Writer.WriteEndElement() ' end Offers
        Writer.WriteEndDocument()
        Writer.Flush()
        
        ms.Seek(0, SeekOrigin.Begin)
        RetXmlDoc.Load(ms)
        
        attrib = RetXmlDoc.CreateAttribute("returnCode")
        attrib.Value = "SUCCESS"
        attrib = RetXmlDoc.SelectSingleNode("//Offers").Attributes.Append(attrib)
      Else
        RetXmlDoc = New XmlDocument()
        RetXmlDoc.LoadXml(GetErrorXML(RetCode, RetMsg))
      End If

    Catch ex As Exception
      MyCommon.Write_Log("KioskWeb.txt", "Exception: " & ex.ToString, True)
      RetXmlDoc = New XmlDocument()
      RetXmlDoc.LoadXml(GetErrorXML(StatusCodes.APPLICATION_EXCEPTION, ex.ToString))
    Finally
      If Not Writer Is Nothing Then
        Writer.Close()
      End If
      If Not ms Is Nothing Then
        ms.Close()
        ms.Dispose()
        ms = Nothing
      End If
      
      If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
      If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
    End Try
    
        Return RetXmlDoc
    End Function
  
    <WebMethod()> _
    Public Function GetWebKioskOffersByLocationCode(ByVal GUID As String, ByVal LocationCode As String) As XmlDocument
        Dim RetXmlDoc As New XmlDocument
        Dim ms As New MemoryStream
        Dim Writer As New XmlTextWriter(ms, System.Text.Encoding.UTF8)
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = ""
        Dim XsdType As Integer = 0
        Dim BatchGUID As String = System.Guid.NewGuid.ToString
        Dim dtOffers As DataTable
        Dim row As DataRow
        Dim attrib As XmlAttribute = Nothing
        Dim LocationID As Long = 0
    
        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
    
            ' validate the request
            If Not IsValidGUID(GUID) Then
                RetCode = StatusCodes.INVALID_GUID
                RetMsg = "GUID " & GUID & " is not valid for the KioskWeb web service."
            End If
      
            ' validate location
            LocationID = GetLocationID(LocationCode)
            If LocationID <= 0 Then
                RetCode = StatusCodes.INVALID_LOCATION
                RetMsg = "Location Code " & LocationCode & " not found in Logix."
            End If
      
            If RetCode = StatusCodes.SUCCESS Then
                Writer.Formatting = Formatting.Indented
                Writer.Indentation = 4
      
                Writer.WriteStartDocument()
                Writer.WriteStartElement("Offers")
      
                MyCommon.QueryStr = "dbo.pa_CPE_KioskWeb_OffersByLocation"
                MyCommon.Open_LRTsp()
                MyCommon.LRTsp.Parameters.Add("@LocationID", SqlDbType.BigInt).Value = LocationID
                dtOffers = MyCommon.LRTsp_select
                MyCommon.Close_LRTsp()
    
                For Each row In dtOffers.Rows
                    WriteKioskWebOffer(Writer, row)
                Next
      
                Writer.WriteEndElement() ' end Offers
                Writer.WriteEndDocument()
                Writer.Flush()

                ms.Seek(0, SeekOrigin.Begin)
                RetXmlDoc.Load(ms)

                attrib = RetXmlDoc.CreateAttribute("returnCode")
                attrib.Value = "SUCCESS"
                attrib = RetXmlDoc.SelectSingleNode("//Offers").Attributes.Append(attrib)
            Else
                RetXmlDoc = New XmlDocument()
                RetXmlDoc.LoadXml(GetErrorXML(RetCode, RetMsg))
            End If

        Catch ex As Exception
            MyCommon.Write_Log("KioskWeb.txt", "Exception: " & ex.ToString, True)
            RetXmlDoc = New XmlDocument()
            RetXmlDoc.LoadXml(GetErrorXML(StatusCodes.APPLICATION_EXCEPTION, ex.ToString))
        Finally
            If Not Writer Is Nothing Then
                Writer.Close()
            End If
            If Not ms Is Nothing Then
                ms.Close()
                ms.Dispose()
                ms = Nothing
            End If
      
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try
    
        Return RetXmlDoc
    End Function
  
    <WebMethod()> _
    Public Function LookupByAlternateID(ByVal GUID As String, ByVal AlternateID As String, ByVal BannerID As Integer) As XmlDocument
        Dim RetXmlDoc As New XmlDocument
        Dim ms As New MemoryStream
        Dim Writer As New XmlTextWriter(ms, System.Text.Encoding.UTF8)
        Dim RetCode As StatusCodes = StatusCodes.SUCCESS
        Dim RetMsg As String = ""
        Dim attrib As XmlAttribute = Nothing
        Dim dt As DataTable
        Dim row As DataRow
        Dim MyAltID As New Copient.AlternateID
        Dim Results(-1) As Copient.AltIdResult
        Dim i As Integer = 0
    
        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            If MyCommon.LXSadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixXS()
    
            ' validate the request
            If Not IsValidGUID(GUID) Then
                RetCode = StatusCodes.INVALID_GUID
                RetMsg = "GUID " & GUID & " is not valid for the KioskWeb web service."
            End If
  
            ' validate the BannerID
            If RetCode = StatusCodes.SUCCESS And BannerID > 0 Then
                MyCommon.QueryStr = "select BannerID from Banners with (NoLock) where Deleted=0 and BannerID=" & BannerID
                dt = MyCommon.LRT_Select
                If dt.Rows.Count = 0 Then
                    RetCode = StatusCodes.BANNER_ID_NOT_FOUND
                    RetMsg = "BannerID " & BannerID & " not found"
                End If
            End If
      
            ' validate the AltID 
            If RetCode = StatusCodes.SUCCESS Then
                Results = MyAltID.FindCustomersByAltID(AlternateID, BannerID)
                If Results Is Nothing OrElse Results.Length = 0 OrElse Results(0).CustomerPK <= 0 Then
                    RetCode = StatusCodes.ALT_ID_NOT_FOUND
                    RetMsg = "Alternate ID " & AlternateID & " was not found"
                    If BannerID > 0 Then RetMsg &= " for banner " & BannerID
                End If
            End If
      
            If RetCode = StatusCodes.SUCCESS Then
                Writer.Formatting = Formatting.Indented
                Writer.Indentation = 4
      
                Writer.WriteStartDocument()
                Writer.WriteStartElement("Cards")
      
                ' iterate over each of the customers that have the specified Alternate ID
                If Results IsNot Nothing AndAlso Results.Length > 0 Then
                    For i = 0 To Results.GetUpperBound(0)
                        If Results(i).CustomerPK > 0 Then
                            MyCommon.QueryStr = "dbo.pa_CPE_KioskWeb_LookupCardData"
                            MyCommon.Open_LXSsp()
                            MyCommon.LXSsp.Parameters.Add("@CustomerPK", SqlDbType.BigInt).Value = Results(i).CustomerPK
                            dt = MyCommon.LXSsp_select
                            MyCommon.Close_LXSsp()
    
                            ' write each of the cards for this customer
                            For Each row In dt.Rows
                                WriteKioskWebCard(Writer, row)
                            Next
            
                        End If
                    Next
                End If

                Writer.WriteEndElement() ' end Cards
                Writer.WriteEndDocument()
                Writer.Flush()

                ms.Seek(0, SeekOrigin.Begin)
                RetXmlDoc.Load(ms)

                attrib = RetXmlDoc.CreateAttribute("returnCode")
                attrib.Value = "SUCCESS"
                attrib = RetXmlDoc.SelectSingleNode("//Cards").Attributes.Append(attrib)
            Else
                RetXmlDoc = New XmlDocument()
                RetXmlDoc.LoadXml(GetCardErrorXML(RetCode, RetMsg))
            End If
        Catch ex As Exception
            MyCommon.Write_Log("KioskWeb.txt", "Exception: " & ex.ToString, True)
            RetXmlDoc = New XmlDocument()
            RetXmlDoc.LoadXml(GetCardErrorXML(StatusCodes.APPLICATION_EXCEPTION, ex.ToString))
        Finally
            If Not Writer Is Nothing Then
                Writer.Close()
            End If
            If Not ms Is Nothing Then
                ms.Close()
                ms.Dispose()
                ms = Nothing
            End If
      
            If MyCommon.LRTadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixRT()
            If MyCommon.LXSadoConn.State <> ConnectionState.Closed Then MyCommon.Close_LogixXS()
        End Try
    
        Return RetXmlDoc
    End Function
  
    Private Sub WriteKioskWebOffer(ByRef Writer As XmlTextWriter, ByVal row As DataRow)
        Dim CouponReward As String = ""
    
        Writer.WriteStartElement("Offer")
        Writer.WriteElementString("OfferID", MyCommon.NZ(row.Item("OfferID"), "0"))
        Writer.WriteElementString("Name", MyCommon.NZ(row.Item("OfferName"), ""))
        Writer.WriteElementString("Description", MyCommon.NZ(row.Item("Description"), ""))
        Writer.WriteElementString("OfferType", MyCommon.NZ(row.Item("OfferType"), ""))
        Writer.WriteElementString("ProgramID", MyCommon.NZ(row.Item("ProgramID"), "0"))
        Writer.WriteElementString("TriggerAmount", MyCommon.NZ(row.Item("TriggerAmount"), "0"))
        Writer.WriteElementString("Category", MyCommon.NZ(row.Item("Category"), "Unknown"))
        Writer.WriteElementString("Action", MyCommon.NZ(row.Item("Action"), ""))

        ' write the pass-thru reward as character data because it may or may not be XML, also
        ' any XML would be varied by reward making a comprehensive universal XSD infeasible.
        Writer.WriteStartElement("CouponReward")
        Writer.WriteCData(MyCommon.NZ(row.Item("CouponReward"), ""))
        Writer.WriteEndElement() ' CouponReward

        Writer.WriteEndElement() ' Offer

    End Sub
  
    Private Sub WriteKioskWebCard(ByRef Writer As XmlTextWriter, ByVal row As DataRow)

    
        Writer.WriteStartElement("Card")
        Writer.WriteElementString("CardID", MyCommon.NZ(MyCryptLib.SQL_StringDecrypt(row.Item("ExtCardID")), "0"))
        Writer.WriteElementString("CardStatus", GetCardStatusText(MyCommon.NZ(row.Item("CardStatusID"), -1)))
        Writer.WriteElementString("CardType", GetCardTypeText(MyCommon.NZ(row.Item("CardTypeID"), -1)))
        Writer.WriteElementString("CustomerPK", MyCommon.NZ(row.Item("CustomerPK"), "0"))
        Writer.WriteEndElement() ' Card

    End Sub
  
    Private Sub WriteKioskWebAccums(ByRef Writer As XmlTextWriter, ByVal row As DataRow)
        Dim CouponReward As String = ""
    
        Writer.WriteStartElement("Offer")
        Writer.WriteElementString("OfferID", MyCommon.NZ(row.Item("OfferID"), "0"))
        Writer.WriteElementString("OfferName", MyCommon.NZ(row.Item("OfferName"), ""))
        Writer.WriteElementString("AccumulationBalance", MyCommon.NZ(row.Item("AccumulationBalance"), 0))
        Writer.WriteEndElement() ' Offer

    End Sub
  
    Private Function IsValidGUID(ByVal GUID As String) As Boolean
        Dim IsValid As Boolean = False
        Dim ConnInc As New Copient.ConnectorInc
    
        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            IsValid = ConnInc.IsValidConnectorGUID(MyCommon, 42, GUID)
        Catch ex As Exception
            IsValid = False
        End Try
    
        Return IsValid
    End Function
  
    'function to chk customer card validation
    Private Function IsValidCustomerCard(ByVal ExtCardID As String, ByVal CardTypeID As String, ByRef RetCode As StatusCodes, ByRef RetMsg As String) As Boolean
        Dim IsValid As Boolean = False
        Dim validationRespCode As CardValidationResponse
        RetCode = StatusCodes.SUCCESS
        RetMsg = ""
        Try
            If (MyCommon.AllowToProcessCustomerCard(ExtCardID, CardTypeID, validationRespCode) = False) Then   
                If validationRespCode <> CardValidationResponse.SUCCESS Then
                    If validationRespCode = CardValidationResponse.CARDIDNOTNUMERIC OrElse validationRespCode = CardValidationResponse.INVALIDCARDFORMAT Then
                        RetCode = StatusCodes.INVALID_CARD
                    ElseIf validationRespCode = CardValidationResponse.CARDTYPENOTFOUND OrElse validationRespCode = CardValidationResponse.INVALIDCARDTYPEFORMAT Then
                        RetCode = StatusCodes.INVALID_CARD
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
    
    Private Function GetLocationID(ByVal LocationCode As String) As Long
        Dim LocationID As Long = 0
        Dim dt As DataTable
        
        Try
            If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
            MyCommon.QueryStr = "select LocationID from Locations with (NoLock) where Deleted=0 and ExtLocationCode='" & LocationCode & "';"
            dt = MyCommon.LRT_Select
            If dt.Rows.Count > 0 Then
                LocationID = MyCommon.NZ(dt.Rows(0).Item("LocationID"), 0)
            End If
        Catch ex As Exception
            LocationID = 0
        End Try
    
        Return LocationID
    End Function
    
    Private Function GetErrorXML(ByVal RetCode As StatusCodes, ByVal ErrorMessage As String) As String
        Dim ErrorXml As New StringBuilder()
    
        ErrorXml.Append("<Offers returnCode=""" & GetReturnCodeText(RetCode) & """>")
        ErrorXml.Append("  <Error>" & ErrorMessage & "</Error>")
        ErrorXml.Append("</Offers>")
    
        Return ErrorXml.ToString
    End Function
  
    Private Function GetCardErrorXML(ByVal RetCode As StatusCodes, ByVal ErrorMessage As String) As String
        Dim ErrorXml As New StringBuilder()
    
        ErrorXml.Append("<Cards returnCode=""" & GetReturnCodeText(RetCode) & """>")
        ErrorXml.Append("  <Error>" & ErrorMessage & "</Error>")
        ErrorXml.Append("</Cards>")
    
        Return ErrorXml.ToString
    End Function

    Private Function GetReturnCodeText(ByVal RetCode As StatusCodes) As String
        Dim RetCodeText As String = "SUCCESS"
    
        Select Case RetCode
            Case StatusCodes.SUCCESS
                RetCodeText = "SUCCESS"
            Case StatusCodes.APPLICATION_EXCEPTION
                RetCodeText = "APPLICATION_EXCEPTION"
            Case StatusCodes.INVALID_LOCATION
                RetCodeText = "INVALID_LOCATION"
            Case StatusCodes.INVALID_GUID
                RetCodeText = "INVALID_GUID"
            Case StatusCodes.ALT_ID_NOT_FOUND
                RetCodeText = "ALT_ID_NOT_FOUND"
            Case StatusCodes.INVALID_CARD   ''added
                RetCodeText = "INVALID_CARD"
        End Select
        Return RetCodeText
    End Function
    
    Private Function GetCardStatusText(ByVal CardStatusID As Integer) As String
        Dim CardStatusText As String = "UNKNOWN"

        Select Case CardStatusID
            Case 1
                CardStatusText = "ACTIVE"
            Case 2
                CardStatusText = "INACTIVE"
            Case 3
                CardStatusText = "CANCELED"
            Case 4
                CardStatusText = "EXPIRED"
            Case 5
                CardStatusText = "LOST_STOLEN"
            Case 6
                CardStatusText = "DEFAULT_CARD"
            Case Else
                CardStatusText = "UNKNOWN"
        End Select
    
        Return CardStatusText
    End Function
  
    Private Function GetCardTypeText(ByVal CustomerTypeID As Integer) As String
        Dim CardTypeText As String = "UNKNOWN"
    
        Select Case CustomerTypeID
            Case 0
                CardTypeText = "CUSTOMER"
            Case 1
                CardTypeText = "HOUSEHOLD"
            Case 2
                CardTypeText = "CAM"
            Case 3
                CardTypeText = "ALTERNATE"
            Case 4
                CardTypeText = "USERNAME"
            Case 5
                CardTypeText = "ASSOCIATE"
            Case 6
                CardTypeText = "Email Address"
            Case 7
                CardTypeText = "SECONDARY MEMBER"
            Case Else
                CardTypeText = "UNKNOWN"
        End Select
        Return CardTypeText
    End Function
  
End Class