<%@ Page Language="vb" Debug="true" CodeFile="cwCB.vb" Inherits="cwCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Xml" %>
<%@ Import Namespace="System.Xml.Xsl" %>
<%@ Import Namespace="System.Xml.XPath" %>
<%@ Import Namespace="Copient.commonShared" %>
<%
  Dim CopientFileName As String = "cwfeed.aspx"
  Dim CopientFileVersion As String = "7.3.1.138972" 
  Dim CopientProject As String = "Copient Logix"
  Dim CopientNotes As String = ""
  
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim AdminUserID As Long
  Dim rst As DataTable
  Dim row As DataRow
  Dim rst2 As DataTable
  Dim MessageString As String = ""
  Dim Caller As String = ""
  Dim Transform As String = ""
  Dim CustomerPK As String = ""
  Dim CustomerID As String = ""
  Dim CustomerTypeID As Integer = 0
  Dim CGroupID As String = ""
  Dim OfferID As String = ""
  
  ' AdminUserID = Verify_AdminUser(Logix)
  MyCommon.AppName = "cwfeed.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixWH()
  MyCommon.Open_LogixXS()
  
  If (LanguageID = 0) Then
    LanguageID = MyCommon.Extract_Val(Request.QueryString("LanguageID"))
  End If
  
  Caller = Request.QueryString("caller")
  
  If Caller = "web" Then
    ' Requests are coming from outside, with details on the querystring.
    If (Request.QueryString("mode") <> "") Then
      If (Request.QueryString("customerpk") = "" And Request.QueryString("customerid") = "") Then
        MessageString = "<b>Customer undefined.</b>"
      Else
        ' At least one customer identifier was provided, so we proceed
        CustomerPK = Request.QueryString("customerpk")
        CustomerID = Request.QueryString("customerid")
        If Request.QueryString("customertypeid") <> "" Then
          CustomerTypeID = MyCommon.Extract_Val(Request.QueryString("customertypeid"))
        End If
        If CustomerID <> "" And CustomerPK = "" Then
          ' A customer ID was provided, but not a PK.  We need the PK, so let's find it.
          MyCommon.QueryStr = "select CustomerPK from Customers with (NoLock) where PrimaryExtID='" & CustomerID & "' and CustomerTypeID=" & CustomerTypeID & ";"
          rst = MyCommon.LXS_Select
          If rst.Rows.Count > 0 Then
            CustomerPK = MyCommon.NZ(rst.Rows(0).Item("CustomerPK"), 0)
          End If
        End If
        
        Transform = Request.QueryString("transform")
        If Request.QueryString("cgroupid") <> "" Then
          CGroupID = Request.QueryString("cgroupid").ToString
        End If
        If Request.QueryString("offerid") <> "" Then
          OfferID = Request.QueryString("offerid").ToString
        End If
        
        ' Handle the various modes:
        
        If (Request.QueryString("mode") = "cust" And CustomerPK <> "") Then
          ' CUST MODE, which gives details regarding the customer's offers
          Send("<?xml version=""1.0"" encoding=""utf-8"" ?>")
          custOffers(CustomerPK, Transform)
          
        ElseIf (Request.QueryString("mode") = "info" And CustomerPK <> "") Then
          ' INFO MODE, which gives details about the customer himself.
          Send("<?xml version=""1.0"" encoding=""utf-8"" ?>")
          custInfo(CustomerPK, Transform)
          
        ElseIf (Request.QueryString("mode") = "optout" And CustomerPK <> "") Then
          ' OPT-OUT MODE, which removes the customer from a customer group
          If CGroupID = "" And OfferID = "" Then
            ' No identifiers provided.
            MessageString = "<b>ID undefined.</b>"
          ElseIf CGroupID = "" And OfferID <> "" Then
            ' Got an offer ID but not a group, so look up the associated group ID(s)
            MyCommon.QueryStr = "select 0 as EngineID, OC.ConditionID, OC.LinkID as CustomerGroupID, CG.AnyCardholder, CG.AnyCustomer, CG.NewCardholders, CG.AnyCAMCardholder, -1 as ROID from CM_ST_OfferConditions as OC with (NoLock) " & _
                                "inner join CustomerGroups as CG with (NoLock) on CG.CustomerGroupID=OC.LinkID " & _
                                "where OC.Deleted=0 and OC.ConditionTypeID=1 and OC.ExcludedID=0 and OC.OfferID=" & OfferID & " " & _
                                " union " & _
                                "select 2 as EngineID, -1 as ConditionID, ICG.CustomerGroupID as CustomerGroupID, CG.AnyCardholder, CG.AnyCustomer, CG.NewCardholders, CG.AnyCAMCardholder, ICG.RewardOptionID as ROID from CPE_ST_IncentiveCustomerGroups as ICG with (NoLock) " & _
                                "inner join CPE_ST_RewardOptions as RO with (NoLock) on RO.RewardOptionID=ICG.RewardOptionID " & _
                                "inner join CustomerGroups as CG with (NoLock) on CG.CustomerGroupID=ICG.CustomerGroupID " & _
                                "where ICG.Deleted=0 and RO.Deleted=0 and ICG.ExcludedUsers=0 and RO.IncentiveID=" & OfferID & ";"
            rst = MyCommon.LRT_Select
            If rst.Rows.Count > 0 Then
              For Each row In rst.Rows
                CGroupID = MyCommon.NZ(row.Item("CustomerGroupID"), 0)
                If (row.Item("AnyCardholder") = 0) AndAlso (row.Item("AnyCustomer") = 0) AndAlso (row.Item("NewCardholders") = 0) AndAlso (row.Item("AnyCAMCardholder") = 0) Then
                  ' It's not a special group, so just remove the customer from it.
                  Send("<?xml version=""1.0"" encoding=""utf-8"" ?>")
                  custOpt(CustomerPK, CGroupID, "out", Transform)
                Else
                  ' It IS a special group, so check to see if there is an excluded group associated with it.
                  MyCommon.QueryStr = "select 0 as EngineID, OC.ConditionID, OC.LinkID as CustomerGroupID, CG.AnyCardholder, CG.AnyCustomer, CG.NewCardholders, CG.AnyCAMCardholder, -1 as ROID from CM_ST_OfferConditions as OC with (NoLock) " & _
                                      "inner join CustomerGroups as CG with (NoLock) on CG.CustomerGroupID=OC.LinkID " & _
                                      "where OC.Deleted=0 and OC.ConditionTypeID=1 and OC.ExcludedID>0 and OC.OfferID=" & OfferID & " " & _
                                      " union " & _
                                      "select 2 as EngineID, -1 as ConditionID, ICG.CustomerGroupID as CustomerGroupID, CG.AnyCardholder, CG.AnyCustomer, CG.NewCardholders, CG.AnyCAMCardholder, ICG.RewardOptionID as ROID from CPE_ST_IncentiveCustomerGroups as ICG with (NoLock) " & _
                                      "inner join CPE_ST_RewardOptions as RO with (NoLock) on RO.RewardOptionID=ICG.RewardOptionID " & _
                                      "inner join CustomerGroups as CG with (NoLock) on CG.CustomerGroupID=ICG.CustomerGroupID " & _
                                      "where ICG.Deleted=0 and RO.Deleted=0 and ICG.ExcludedUsers=1 and RO.IncentiveID=" & OfferID & ";"
                  rst2 = MyCommon.LRT_Select
                  If rst2.Rows.Count > 0 Then
                    ' Yes, there's an excluded group, so in order to opt the customer out of the offer, add the customer to the excluded group.
                    CGroupID = MyCommon.NZ(rst2.Rows(0).Item("CustomerGroupID"), 0)
                    Send("<?xml version=""1.0"" encoding=""utf-8"" ?>")
                    custOpt(CustomerPK, CGroupID, "in", Transform)
                  Else
                    ' No excluded groups, so the attempt to opt-out is invalid.
                    MessageString = "<b>Cannot opt out of Any Customer, Any Cardholder or New Cardholders groups.</b>"
                  End If
                End If
              Next
            Else
              MessageString = "<b>No associated customer groups found for that offer.</b>"
            End If
          Else
            ' Got a customer group ID, so just remove the customer.
            If CGroupID > 2 Then
              Send("<?xml version=""1.0"" encoding=""utf-8"" ?>")
              custOpt(CustomerPK, CGroupID, "out", Transform)
            Else
              MessageString = "<b>Cannot opt out of that group.</b>"
            End If
          End If
          
        ElseIf (Request.QueryString("mode") = "optin" And CustomerPK <> "") Then
          ' OPT-IN MODE, which adds the customer to a customer group
          If CGroupID = "" And OfferID = "" Then
            ' No identifiers provided.
            MessageString = "<b>ID undefined.</b>"
          ElseIf CGroupID = "" And OfferID <> "" Then
            ' Got an offer ID but not a group, so look up the associated group ID(s)
            MyCommon.QueryStr = "select 0 as EngineID, OC.ConditionID, OC.LinkID as CustomerGroupID, CG.AnyCardholder, CG.AnyCustomer, CG.NewCardholders, CG.AnyCAMCardholder, -1 as ROID from CM_ST_OfferConditions as OC with (NoLock) " & _
                                "inner join CustomerGroups as CG with (NoLock) on CG.CustomerGroupID=OC.LinkID " & _
                                "where OC.Deleted=0 and OC.ConditionTypeID=1 and OC.ExcludedID=0 and OC.OfferID=" & OfferID & " " & _
                                " union " & _
                                "select 2 as EngineID, -1 as ConditionID, ICG.CustomerGroupID as CustomerGroupID, CG.AnyCardholder, CG.AnyCustomer, CG.NewCardholders, CG.AnyCAMCardholder, ICG.RewardOptionID as ROID from CPE_ST_IncentiveCustomerGroups as ICG with (NoLock) " & _
                                "inner join CPE_ST_RewardOptions as RO with (NoLock) on RO.RewardOptionID=ICG.RewardOptionID " & _
                                "inner join CustomerGroups as CG with (NoLock) on CG.CustomerGroupID=ICG.CustomerGroupID " & _
                                "where ICG.Deleted=0 and RO.Deleted=0 and ICG.ExcludedUsers=0 and RO.IncentiveID=" & OfferID & ";"
            rst = MyCommon.LRT_Select
            If rst.Rows.Count > 0 Then
              For Each row In rst.Rows
                CGroupID = MyCommon.NZ(row.Item("CustomerGroupID"), 0)
                If (row.Item("AnyCardholder") = 0) AndAlso (row.Item("AnyCustomer") = 0) AndAlso (row.Item("NewCardholders") = 0) AndAlso (row.Item("AnyCAMCardholder") = 0) Then
                  ' It's not a special group, so just add the customer to it.
                  Send("<?xml version=""1.0"" encoding=""utf-8"" ?>")
                  custOpt(CustomerPK, CGroupID, "in", Transform)
                Else
                  ' It IS a special group, so check to see if there is an excluded group associated with it.
                  MyCommon.QueryStr = "select 0 as EngineID, OC.ConditionID, OC.LinkID as CustomerGroupID, CG.AnyCardholder, CG.AnyCustomer, CG.NewCardholders, CG.AnyCAMCardholder, -1 as ROID from CM_ST_OfferConditions as OC with (NoLock) " & _
                                      "inner join CustomerGroups as CG with (NoLock) on CG.CustomerGroupID=OC.LinkID " & _
                                      "where OC.Deleted=0 and OC.ConditionTypeID=1 and OC.ExcludedID>0 and OC.OfferID=" & OfferID & " " & _
                                      " union " & _
                                      "select 2 as EngineID, -1 as ConditionID, ICG.CustomerGroupID as CustomerGroupID, CG.AnyCardholder, CG.AnyCustomer, CG.NewCardholders, CG.AnyCAMCardholder, ICG.RewardOptionID as ROID from CPE_ST_IncentiveCustomerGroups as ICG with (NoLock) " & _
                                      "inner join CPE_ST_RewardOptions as RO with (NoLock) on RO.RewardOptionID=ICG.RewardOptionID " & _
                                      "inner join CustomerGroups as CG with (NoLock) on CG.CustomerGroupID=ICG.CustomerGroupID " & _
                                      "where ICG.Deleted=0 and RO.Deleted=0 and ICG.ExcludedUsers=1 and RO.IncentiveID=" & OfferID & ";"
                  rst2 = MyCommon.LRT_Select
                  If rst2.Rows.Count > 0 Then
                    ' Yes, there's an excluded group, so in order to opt the customer into the offer, drop the customer from the excluded group.
                    CGroupID = MyCommon.NZ(rst2.Rows(0).Item("CustomerGroupID"), 0)
                    Send("<?xml version=""1.0"" encoding=""utf-8"" ?>")
                    custOpt(CustomerPK, CGroupID, "out", Transform)
                  Else
                    ' No excluded groups, so the attempt to opt-in is invalid.
                    MessageString = "<b>Cannot opt customer into the Any Customer, Any Cardholder or New Cardholders groups.</b>"
                  End If
                End If
              Next
            Else
              MessageString = "<b>No associated customer groups found for that offer.</b>"
            End If
          Else
            ' Got a customer group ID, so just add the customer.
            Send("<?xml version=""1.0"" encoding=""utf-8"" ?>")
            custOpt(CustomerPK, CGroupID, "in", Transform)
          End If
          
        Else
          ' No mode provided
          MessageString = "<b>Invalid mode.</b>"
        End If
      End If
    Else
      MessageString = "<b>Mode undefined.</b>"
    End If
    
  Else
    ' Requests are coming from within the customer-facing website.
    If (Request.Form("mode") <> "") Then
      If (Session("customerpk") Is Nothing) Then
        MessageString = "<b>Session not valid</b>"
      Else
        Transform = Request.Form("transform")
        If (Request.Form("mode") = "cust" And Session("customerpk").ToString <> "") Then
          custOffers(Session("customerpk").ToString, Transform)
        ElseIf (Request.Form("mode") = "info" And Session("customerpk").ToString <> "") Then
          custInfo(Session("customerpk").ToString, Transform)
        Else
          MessageString = "<b>Invalid session mode.</b>"
        End If
      End If
    Else
      MessageString = "<b>Mode undefined.</b>"
    End If
  End If
  
  If (MessageString <> "") Then
    Send(MessageString)
  End If
  
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixWH()
  MyCommon.Close_LogixXS()
  Logix = Nothing
  MyCommon = Nothing
%>

<script runat="server">
  Public DefaultLanguageID
  Public MyCommon As New Copient.CommonInc
  
  '----------------------------------------------------------------------------------
  
  Sub custInfo(ByVal CustomerPK As String, ByVal XFORM As String)
    Dim outStr As String
    Dim outputStr As New StringWriter
    
    Response.Cache.SetCacheability(HttpCacheability.NoCache)
    
    If (XFORM = "XML") Then
      outputStr.WriteLine(xmlHeader(CustomerPK))
    End If
    
    outStr = sendCustInfo(CustomerPK, XFORM)
    
    If (outStr.Length > 0) Then
      outputStr.WriteLine(outStr)
    End If
    
    If (XFORM = "XML") Then
      outputStr.WriteLine(xmlFooter())
      Dim outString As String = outputStr.ToString
      If (XFORM = "XFORM") Then
        Dim EntireFile As String
        Dim oRead As System.IO.StreamReader
        Dim oFile As System.IO.File
        oRead = oFile.OpenText(Server.MapPath("cw.xsl"))
        EntireFile = oRead.ReadToEnd()
        oRead.Close()
        outString = transform(outString, EntireFile)
      End If
      Send(outString)
    Else
      Send(outStr)
    End If
  End Sub
  
  Sub custOffers(ByVal ID As String, ByVal XFORM As String)
    Dim rst As DataTable
    Dim row As DataRow
    Dim outSTr As String = ""
    Dim outSTr2 As String = ""
    Dim outputStr As New StringWriter
    Dim currentoffers As String = ""
    Dim outStrOffers As String = ""
    Dim outStrGroups As String = ""
    Dim outStrTransactions As String = ""
    
    Response.Cache.SetCacheability(HttpCacheability.NoCache)
    MyCommon.Open_LogixRT()
    
    If (XFORM = "XML") Then
      outputStr.WriteLine(xmlHeader(ID))
    End If
    
    ' figure out the availble offers
    outSTr = Send_XMLCurrentOffers(ID, currentoffers, XFORM)
    outStrOffers = outSTr
    
    If (XFORM = "XML") Then
      If (outSTr.Length > 0) Then
        outputStr.WriteLine("<Offers>")
        outputStr.WriteLine(outSTr)
        outputStr.WriteLine("</Offers>")
      End If
    End If
    
    ' figure out the website offers to opt in
    outSTr = Send_XMLGroupOffers(ID, currentoffers, XFORM)
    outStrGroups = outSTr
    
    ' figure out the recent tranasction data
    outSTr2 = Send_XMLTransactions(ID, currentoffers, XFORM)
    outStrTransactions = outSTr2
    
    If (XFORM = "XML") Then
      If (outSTr.Length > 0) Then
        outputStr.WriteLine("<Groups>")
        outputStr.WriteLine(outSTr)
        outputStr.WriteLine("</Groups>")
      End If
      
      outputStr.WriteLine(xmlFooter())
      
      Dim outString As String = outputStr.ToString
      
      If (XFORM = "XFORM") Then
        Dim EntireFile As String
        Dim oRead As System.IO.StreamReader
        Dim oFile As System.IO.File
        oRead = oFile.OpenText(Server.MapPath("cw.xsl"))
        EntireFile = oRead.ReadToEnd()
        oRead.Close()
        outString = transform(outString, EntireFile)
      End If
      
      Send(outString)
    Else
      Send("  <h1>• Your offers •</h1>")
      Send("  <hr />")
      Send("  <br />")
      Send("  <p>")
      Send("    Welcome to your personal offers page!  On the left are offer you can currently use, and on the right are those that you're eligible to join.  Just click the buttons to opt into or out of whichever you choose!")
      Send("  </p>")
      Send("")
      Send("  <div id=""offers"">")
      Send("    <h2>Active Offers</h2>")
      Send(outStrOffers)
      Send("  </div>")
      Send("")
      Send("  <div class=""gutter"">")
      Send("  </div>")
      Send("")
      Send("  <div id=""groups"">")
      Send("    <h2>Available Offers</h2>")
      Send(outStrGroups)
      Send("  </div>")
      Send("  <br clear=""left"" />")
      Send("")
      If MyCommon.Fetch_WebOption(7) > 0 Then
        Send("  <hr />")
        Send("  <h1>• Your recent transactions •</h1>")
        Send("  <div id=""transactions"">")
        Send(outStrTransactions)
        Send("  </div>")
      End If
      Send("")
    End If
    
    MyCommon.Close_LogixRT()
  End Sub
  
  Sub custOpt(ByVal CustomerPK As String, ByVal CustomerGroupID As String, ByVal OptType As String, ByVal XFORM As String)
    Dim outStr As String
    Dim outputStr As New StringWriter
    Dim rst As DataTable
    Dim PrimaryExtID As String = ""
    Dim CustomerTypeID As Integer = 0
    Dim Successful As Boolean = False
    Dim outString As String = ""
    
    MyCommon.Open_LogixRT()
    MyCommon.Open_LogixXS()

    Response.Cache.SetCacheability(HttpCacheability.NoCache)
    
    If (XFORM = "XML") Then
      MyCommon.QueryStr = "select PrimaryExtID, CustomerTypeID from Customers with (NoLock) where CustomerPK=" & CustomerPK & ";"
      rst = MyCommon.LXS_Select()
      If rst.Rows.Count > 0 Then
        PrimaryExtID = MyCommon.NZ(rst.Rows(0).Item("PrimaryExtID"), 0)
        CustomerTypeID = MyCommon.NZ(rst.Rows(0).Item("CustomerTypeID"), 0)
      End If
      If (OptType = "out") Then
        MyCommon.QueryStr = "dbo.pt_GroupMembership_Delete"
        MyCommon.Open_LXSsp()
        MyCommon.LXSsp.Parameters.Add("@ExtCustomerID", SqlDbType.NVarChar, 26).Value = PrimaryExtID
        MyCommon.LXSsp.Parameters.Add("@CustomerGroupID", SqlDbType.BigInt).Value = CustomerGroupID
        MyCommon.LXSsp.Parameters.Add("@CustomerTypeID", SqlDbType.Int).Value = CustomerTypeID
        MyCommon.LXSsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
        MyCommon.LXSsp.ExecuteNonQuery()
        If MyCommon.LXSsp.Parameters("@Status").Value = 0 Then
          Successful = True
        ElseIf MyCommon.LXSsp.Parameters("@Status").Value = -1 Then
          Successful = False
        End If
        If Successful Then
          MyCommon.Activity_Log(4, CustomerGroupID, 1, Copient.PhraseLib.Lookup("history.cgroup-optout", LanguageID) & " " & PrimaryExtID)
          MyCommon.QueryStr = "update CustomerGroups with (RowLock) set LastUpdate=getdate() where CustomerGroupID=" & CustomerGroupID & ";"
          MyCommon.LRT_Execute()
          outStr = sendOptAck(PrimaryExtID, CustomerGroupID, OptType, 1, XFORM)
        Else
          outStr = sendOptAck(PrimaryExtID, CustomerGroupID, OptType, 0, XFORM)
        End If
      ElseIf (OptType = "in") Then
        MyCommon.QueryStr = "dbo.pt_GroupMembership_Insert"
        MyCommon.Open_LXSsp()
        MyCommon.LXSsp.Parameters.Add("@ExtCustomerID", SqlDbType.NVarChar, 26).Value = PrimaryExtID
        MyCommon.LXSsp.Parameters.Add("@CustomerTypeID", SqlDbType.Int).Value = CustomerTypeID
        MyCommon.LXSsp.Parameters.Add("@CustomerGroupID", SqlDbType.BigInt).Value = CustomerGroupID
        MyCommon.LXSsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
        MyCommon.LXSsp.ExecuteNonQuery()
        If MyCommon.LXSsp.Parameters("@Status").Value = 0 Then
          Successful = True
        ElseIf MyCommon.LXSsp.Parameters("@Status").Value = -1 Then
          Successful = False
        End If
        If Successful Then
          MyCommon.Activity_Log(4, CustomerGroupID, 1, Copient.PhraseLib.Lookup("history.cgroup-optin", LanguageID) & " " & PrimaryExtID)
          MyCommon.QueryStr = "update CustomerGroups with (RowLock) set LastUpdate=getdate() where CustomerGroupID=" & CustomerGroupID & ";"
          MyCommon.LRT_Execute()
          outStr = sendOptAck(PrimaryExtID, CustomerGroupID, OptType, 1, XFORM)
        Else
          outStr = sendOptAck(PrimaryExtID, CustomerGroupID, OptType, 0, XFORM)
        End If
      End If
      
      outputStr.WriteLine(xmlHeader(CustomerPK))
      outputStr.WriteLine(outStr)
      outputStr.WriteLine(xmlFooter())
      
      outString = outputStr.ToString
      Send(outString)
      
    End If
    
    MyCommon.Close_LogixRT()
    MyCommon.Close_LogixXS()
    
  End Sub
  
  Function xmlHeader(ByVal CustomerPK As Long) As String
    Return "<CustWeb CustomerPK=""" & CustomerPK & """ Mode=""Home"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">"
  End Function
  
  Function xmlFooter() As String
    Return "</CustWeb>"
  End Function
  
  Function transform(ByVal xmlDoc As String, ByVal xslDoc As String) As String
    Dim newstream As New MemoryStream
    Dim myXPathDocument As XPathDocument = New XPathDocument(stringToStream(xmlDoc))
    Dim myXSLPathDocument As XPathDocument = New XPathDocument(stringToStream(xslDoc))
    Dim myXslTransform As XslCompiledTransform = New XslCompiledTransform()
    ' Dim writer As XmlTextWriter = New XmlTextWriter(Path & "transform.xml", Nothing)
    
    myXslTransform.Load(myXSLPathDocument)
    myXslTransform.Transform(myXPathDocument, Nothing, newstream)
    'myxsltransform.Transform(myxpathdocument,nothing,
    
    Dim encoder As New System.Text.UTF8Encoding()
    Dim buffer As Byte()
    buffer = newstream.GetBuffer()
    
    Return encoder.GetString(buffer, 0, buffer.GetLength(0))
    'Return newstream
  End Function
  
  Function stringToStream(ByVal str As String) As Stream
    Dim myEncoder As New System.Text.ASCIIEncoding
    Dim bytes As Byte() = myEncoder.GetBytes(str)
    Dim ms As MemoryStream = New MemoryStream(bytes)
    Return (CType(ms, Stream))
  End Function
  
  Function Send_XMLCurrentOffers(ByVal Identifier As String, ByRef currentoffers As String, ByVal Transformation As String) As String
    Dim rst As System.Data.DataTable
    Dim row As System.Data.DataRow
    Dim rst2 As System.Data.DataTable
    Dim row2 As System.Data.DataRow
    Dim rst3 As System.Data.DataTable
    Dim row3 As System.Data.DataRow
    Dim rst4 As System.Data.DataTable
    Dim row4 As System.Data.DataRow
    Dim rst5 As System.Data.DataTable
    Dim row5 As System.Data.DataRow
    Dim rstWeb As System.Data.DataTable
    Dim rstExcluded As System.Data.DataTable
    Dim rowCount As Integer
    Dim CustomerPK As Long
    Dim ExtID As String
    Dim Employee As Boolean
    Dim OfferGroupType As Integer
    Dim ExtCustomerID As String = ""
    Dim CustomerGroupID As Long
    Dim OfferID As Integer
    Dim OfferName As String
    Dim OfferDesc As String
    Dim OfferStart As Date
    Dim OfferEnd As Date
    Dim OfferOdds As Integer
    Dim InstantWin As Integer
    Dim OfferDaysLeft As Integer
    Dim GraphicsFileName As String = ""
    Dim PrintedMessage As String = ""
    Dim AccumulationBalance As String = ""
    Dim ProgramID As String = ""
    Dim ProgID As Integer
    Dim ProgramName As String = ""
    Dim Amount As Long
    Dim AllowOptOut As Integer
    Dim EmployeesOnly As Integer
    Dim EmployeesExcluded As Integer
    Dim dt As New DataTable
    Dim retString As New StringBuilder
    Dim CustomerGroups As String = "0"
    Dim CgBuf As New StringBuilder()
    
    MyCommon.Open_LogixRT()
    MyCommon.Open_LogixXS()
    
    'First off, we check to see if there's an identifier in the URL
    If (Identifier <> "") Then
      ' There is, so first off we grab the customer's information
      If (IsNumeric(Identifier)) Then
        'The identifier the customer supplied is all numbers, so assume it's a card number
        MyCommon.QueryStr = "select C.CustomerPK, C.CustomerTypeID, C.Employee, C.CustomerStatusID, CE.Email " & _
                            "from Customers as C " & _
                            "left join CustomerExt as CE on CE.CustomerPK=C.CustomerPK " & _
                            "where C.CustomerPK='" & Identifier & "';"
      End If
      rst = MyCommon.LXS_Select
      If (rst.Rows.Count > 0) Then
        'A customer was found, so let's assign values to variables
        CustomerPK = rst.Rows(0).Item("CustomerPK")
        ExtID = GetCardNumber(CustomerPK, MyCommon.NZ(rst.Rows(0).Item("CustomerTypeID"), 0))
        Employee = rst.Rows(0).Item("Employee")
        
        'Next, get the associated customer groups...
        MyCommon.QueryStr = "select distinct CustomerGroupID from GroupMembership where CustomerPK=" & CustomerPK & " and Deleted=0;"
        rst = MyCommon.LXS_Select()
        rst.Rows.Add(New String() {"1"})
        rst.Rows.Add(New String() {"2"})
        rowCount = rst.Rows.Count
        ' build up a customer group ID List for this customer
        For Each row In rst.Rows
          If (CgBuf.Length > 0) Then CgBuf.Append(",")
          CgBuf.Append(MyCommon.NZ(row.Item("CustomerGroupID"), 0))
        Next
        CustomerGroups = CgBuf.ToString
        
        'The customer's in at least one group, so for each one we'll grab the associated offer(s)
        For Each row In rst.Rows
          MyCommon.QueryStr = "select distinct O.OfferID, O.ExtOfferID, O.IsTemplate, O.CMOADeployStatus, O.StatusFlag, O.OddsOfWinning, O.InstantWin, " & _
                              "O.Name, O.Description, O.ProdStartDate, O.ProdEndDate, 0 as AllowOptOut, O.EmployeeFiltering as EmployeesOnly, O.NonEmployeesOnly as EmployeesExcluded, LinkID, OID.EngineID " & _
                              "from CM_ST_Offers as O with (NoLock) " & _
                              "left join CM_ST_OfferConditions as OC with (NoLock) on OC.OfferID=O.OfferID " & _
                              "inner join OfferIDs as OID with (NoLock) on OID.OfferID=O.OfferID " & _
                              "where O.IsTemplate=0 and O.Deleted=0 and O.CMOADeployStatus=1 and OC.Deleted=0 and OC.ConditionTypeID=1 " & _
                              "and O.DisabledOnCFW=0 and ProdEndDate>'" & Today.AddDays(-1).ToString & "' and LinkID=" & row.Item("CustomerGroupID") & _
                              "union all " & _
                              "select distinct I.IncentiveID, I.ClientOfferID, I.IsTemplate, I.CPEOADeployStatus, I.StatusFlag, 0 as OddsOfWinning, 0 as InstantWin, " & _
                              "I.IncentiveName, Convert(nvarchar(2000),I.Description) as Description, I.StartDate, I.EndDate, I.AllowOptOut, I.EmployeesOnly, I.EmployeesExcluded, ICG.CustomerGroupID, OID.EngineID " & _
                              "from CPE_ST_Incentives as I with (NoLock) " & _
                              "left join CPE_ST_RewardOptions as RO with (NoLock) on I.IncentiveID=RO.IncentiveID and RO.TouchResponse=0 and RO.Deleted=0 " & _
                              "left join CPE_ST_IncentiveCustomerGroups as ICG with (NoLock) on RO.RewardOptionID=ICG.RewardOptionID and ICG.ExcludedUsers=0 " & _
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
            OfferStart = row2.Item("ProdStartDate")
            OfferEnd = row2.Item("ProdEndDate")
            OfferOdds = row2.Item("OddsOfWinning")
            InstantWin = MyCommon.NZ(row2.Item("InstantWin"), 0)
            OfferDaysLeft = DateDiff("d", Today, OfferEnd)
            AllowOptOut = IIf(MyCommon.NZ(row2.Item("AllowOptOut"), False), 1, 0)
            EmployeesOnly = IIf(MyCommon.NZ(row2.Item("EmployeesOnly"), False), 1, 0)
            EmployeesExcluded = IIf(MyCommon.NZ(row2.Item("EmployeesExcluded"), False), 1, 0)
            CustomerGroupID = row2.Item("LinkID")
            
            ' Filter out the website offers
            MyCommon.QueryStr = "select OfferID from OfferIDs where OfferID=" & OfferID & " and EngineID=3;"
            rstWeb = MyCommon.LRT_Select
            
            ' Filter out the offers where the customer is in the excluded customer group
            MyCommon.QueryStr = "select ExcludedID from CM_ST_OfferConditions with (NoLock) " & _
                                "where OfferID=" & OfferID & " and ExcludedID in (" & CustomerGroups & ") " & _
                                "union " & _
                                "select CustomerGroupID as ExcludedID from CPE_ST_IncentiveCustomerGroups ICG with (NoLock) " & _
                                "inner join CPE_ST_RewardOptions RO with (NoLock) on RO.RewardOptionID=ICG.RewardOptionID " & _
                                "where ICG.Deleted=0 and RO.Deleted=0 and RO.IncentiveID=" & OfferID & " and ExcludedUsers=1 " & _
                                "and CustomerGroupID in (" & CustomerGroups & ");"
            rstExcluded = MyCommon.LRT_Select
            
            If (rstWeb.Rows.Count = 0 AndAlso rstExcluded.Rows.Count = 0) Then
              
              'Find the name of the associated (and non-excluding) location group
              MyCommon.QueryStr = "select OL.OfferID, OL.LocationGroupID, OL.Excluded, LG.Name from OfferLocations as OL " & _
                                  "left join LocationGroups as LG on LG.LocationGroupID=OL.LocationGroupID " & _
                                  "where OL.OfferID=" & OfferID & " and OL.Excluded=0;"
              rst3 = MyCommon.LRT_Select
              currentoffers += OfferID & ","
              
              'Find any associated points programs
              MyCommon.QueryStr = "select O.Offerid, LinkID, ProgramName, PP.ProgramID, PromoVarID from CM_ST_OfferRewards as OFR with (NoLock) " & _
                                  "left join CM_ST_RewardPoints as RP with (NoLock) on RP.RewardPointsID=OFR.LinkID " & _
                                  "left join PointsPrograms as PP with (NoLock) on RP.ProgramID=PP.ProgramID " & _
                                  "left join CM_ST_Offers as O with (NoLock) on O.OfferID=OFR.OfferID " & _
                                  "where (RewardTypeID=2 and O.Deleted=0 and OFR.Deleted=0) " & _
                                  "and RP.ProgramID is not null " & _
                                  "and O.OfferID=" & OfferID & _
                                  " union " & _
                                  "select " & OfferID & " as OfferID, D.OutputID, PP.ProgramName, PP.ProgramID, PP.PromoVarID " & _
                                  "from CPE_ST_Deliverables D with (NoLock) inner join CPE_ST_DeliverablePoints DP with (NoLock) on D.OutputID=DP.PKID " & _
                                  "inner join PointsPrograms PP with (NoLock) on DP.ProgramID=PP.ProgramID " & _
                                  "where D.RewardOptionID in (select RO.RewardOptionID from CPE_ST_RewardOptions RO with (NoLock) where IncentiveID=" & OfferID & ") " & _
                                  "and D.Deleted=0 and DP.Deleted=0 and PP.Deleted=0 and D.DeliverableTypeID=8;"
              rst4 = MyCommon.LRT_Select
              For Each row4 In rst4.Rows
                ProgramID = row4.Item("ProgramID")
                ProgramName = MyCommon.NZ(row4.Item("ProgramName"), "unknown").ToString.Replace(",", " ")
              Next
              
              If (ProgramName <> "" And ProgramID <> "") Then
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
              
              'So finally we have all the info we need. Now display it.
              If (Transformation = "XML") Then
                retString.Append("  <Offer>" & vbCrLf)
                retString.Append("    <OfferID>" & OfferID & "</OfferID>" & vbCrLf)
                retString.Append("    <Name>" & OfferName & "</Name>" & vbCrLf)
                retString.Append("    <Description>" & OfferDesc & "</Description>" & vbCrLf)
                retString.Append("    <StartDate>" & OfferStart & "</StartDate>" & vbCrLf)
                retString.Append("    <EndDate>" & OfferEnd & "</EndDate>" & vbCrLf)
                retString.Append("    <AllowOptOut>" & AllowOptOut & "</AllowOptOut>" & vbCrLf)
                retString.Append("    <EmployeesOnly>" & EmployeesOnly & "</EmployeesOnly>" & vbCrLf)
                retString.Append("    <EmployeesExcluded>" & EmployeesExcluded & "</EmployeesExcluded>" & vbCrLf)
                retString.Append("    <Points>" & Amount & "</Points>" & vbCrLf)
                retString.Append("    <Accumulation>" & AccumulationBalance & "</Accumulation>" & vbCrLf)
                retString.Append("    <BodyText>" & PrintedMessage & "</BodyText>" & vbCrLf)
                retString.Append("    <Graphic>" & GraphicsFileName & "</Graphic>" & vbCrLf)
                retString.Append("  </Offer>" & vbCrLf)
              ElseIf (Transformation = "HTML") Then
                retString.Append("<div class=""offer"" id=""offer" & OfferID & """>" & vbCrLf)
                retString.Append("  <h3 class=""name"" alt=""Offer# " & OfferID & ", CGroup# " & CustomerGroupID & """ title=""Offer " & OfferID & ", CGroup " & CustomerGroupID & """>" & OfferName & "</h3>" & vbCrLf)
                If (GraphicsFileName <> "") Then
                  retString.Append("    <img src=""images\" & GraphicsFileName & """ align=""right"" />")
                End If
                retString.Append("  <div class=""description"">" & vbCrLf)
                retString.Append("    " & OfferDesc & vbCrLf)
                If (PrintedMessage <> "") Then
                  retString.Append(" (<a href=""javascript:openNamedPopup('prntmsg.aspx?identifier=" & Identifier & "&OfferID=" & OfferID & "', 'PrntMsg');"">Details</a>)" & vbCrLf)
                End If
                retString.Append("  </div>" & vbCrLf)
                retString.Append("  <div class=""valid"">" & vbCrLf)
                retString.Append("    Valid <span class=""startdate"">" & OfferStart & "</span> - <span class=""enddate"">" & OfferEnd & "</span><br />" & vbCrLf)
                If (OfferDaysLeft > 1) Then
                  retString.Append("    It will expire in " & OfferDaysLeft & " days.<br />")
                ElseIf (OfferDaysLeft = 1) Then
                  retString.Append("    It will expire tomorrow.<br />")
                ElseIf (OfferDaysLeft = 0) Then
                  retString.Append("    It expires today.<br />")
                ElseIf (OfferDaysLeft = -1) Then
                  retString.Append("    It expired yesterday.<br />")
                ElseIf (OfferDaysLeft < -1) Then
                  retString.Append("    It expired " & Math.Abs(OfferDaysLeft) & " days ago.<br />")
                End If
                retString.Append("  </div>")
                If (ProgramName <> "") Then
                  retString.Append("    <br class=""half"" />")
                  retString.Append("    Your ""<span alt=""Program " & ProgramID & """ title=""Program " & ProgramID & """>" & ProgramName & "</span>"" balance: " & Amount)
                End If
                retString.Append("</div>")
                retString.Append("<br />")
                retString.Append("<hr />")
                retString.Append("")
              ElseIf (Transformation = "nothing") Then
              Else
              End If
              ProgramID = ""
              ProgramName = ""
            End If
          Next
        Next
      End If
    End If
    Return retString.ToString
  End Function
  
  Public Function Send_XMLGroupOffers(ByRef Identifier As String, ByVal CurrentOffers As String, ByVal Transformation As String)
    Dim rst As System.Data.DataTable
    Dim rst2 As System.Data.DataTable
    Dim row2 As System.Data.DataRow
    Dim rst3 As System.Data.DataTable
    Dim rst4 As System.Data.DataTable
    Dim row4 As System.Data.DataRow
    Dim rstWeb As System.Data.DataTable
    Dim rstCG As System.Data.DataTable
    Dim CurrentOffersClause As String = ""
    Dim CPECurrentOffers As String = ""
    Dim CustomerGroups As New StringBuilder()
    Dim CustomerPK As Long
    Dim CustomerGroupID As Long
    Dim RewardGroupID As Long
    Dim ROID As Long
    Dim ExtID As String
    Dim Employee As Boolean
    Dim ExtCustomerID As String = ""
    Dim OfferID As Integer
    Dim OfferName As String
    Dim OfferDesc As String
    Dim OfferStart As Date
    Dim OfferEnd As Date
    Dim OfferOdds As Integer
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
    
    'First off, we check to see if there's an identifier in the URL
    If (Identifier <> "") Then
      
      'There is, so first off we grab the customer's information
      'The identifier the customer supplied is all numbers, so assume it's a card number
      MyCommon.QueryStr = "select C.CustomerPK, C.CustomerStatusID, C.Employee, CE.Email, CustomerTypeID " & _
                          "from Customers as C with (NoLock) " & _
                          "left join CustomerExt as CE with (NoLock) on CE.CustomerPK=C.CustomerPK " & _
                          "where C.CustomerPK='" & Identifier & "';"
      rst = MyCommon.LXS_Select
      If (rst.Rows.Count > 0) Then
        'A customer was found, so let's assign values to variables
        CustomerPK = rst.Rows(0).Item("CustomerPK")
        CustomerTypeID = MyCommon.NZ(rst.Rows(0).Item("CustomerTypeID"), 0)
        ExtID = GetCardNumber(CustomerPK, CustomerTypeID)
        Employee = rst.Rows(0).Item("Employee")
        
        If (CurrentOffers.Length > 0) Then
          CurrentOffers = CurrentOffers.Substring(0, CurrentOffers.Length - 1)
          CPECurrentOffers = " and I.IncentiveID not in (" & CurrentOffers & ") "
        End If
        
        'Next, get the associated customer groups...
        MyCommon.QueryStr = "select CustomerGroupID from GroupMembership where CustomerPK=" & CustomerPK & " and Deleted=0;"
        rstCG = MyCommon.LXS_Select()
        rstCG.Rows.Add(New String() {"1"})
        rstCG.Rows.Add(New String() {"2"})
        rowCount = rstCG.Rows.Count
        For i = 0 To rowCount - 1
          CustomerGroups.Append(MyCommon.NZ(rstCG.Rows(i).Item("CustomerGroupID"), -1))
          If (i < rowCount - 1) Then CustomerGroups.Append(",")
        Next
        MyCommon.QueryStr = "select I.IncentiveID, I.IncentiveName, I.Description, I.StartDate, I.EndDate, ICG.CustomerGroupID, " & _
                            "ICG.ExcludedUsers, D.OutputID as RewardGroup, I.AllowOptOut, I.EmployeesOnly, I.EmployeesExcluded, RO.RewardOptionID " & _
                            "from CPE_ST_Incentives as I with (NoLock) " & _
                            "inner join OfferIDs as OID with (NoLock) on I.IncentiveID=OID.OfferID " & _
                            "inner join CPE_ST_RewardOptions as RO with (NoLock) on RO.IncentiveID=I.IncentiveID and RO.TouchResponse=0 and RO.Deleted=0 " & _
                            "inner join CPE_ST_Deliverables as D with (NoLock) on D.RewardOptionID=RO.RewardOptionID and D.Deleted=0 and DeliverableTypeID=5 and RewardOptionPhase=3 " & _
                            "inner join CPE_ST_IncentiveCustomerGroups as ICG with (NoLock) on ICG.RewardOptionID=RO.RewardOptionID " & _
                            " and ICG.Deleted=0 and ICG.CustomerGroupID in (" & CustomerGroups.ToString & ") and ICG.ExcludedUsers=0 " & _
                            "where I.Deleted=0 And I.EndDate>='" & Today.AddDays(-1).ToString & "' and OID.EngineID=3 " & CPECurrentOffers & ";"
        rst2 = MyCommon.LRT_Select
        If (rst2.Rows.Count > 0) Then
          For Each row2 In rst2.Rows
            OfferID = row2.Item("IncentiveID")
            CurrentOffers += OfferID & ","
            OfferName = row2.Item("IncentiveName")
            OfferDesc = row2.Item("Description")
            If (OfferDesc.Trim() = "") Then OfferDesc = OfferName
            OfferStart = row2.Item("StartDate")
            OfferEnd = row2.Item("EndDate")
            'OfferOdds = row2.Item("OddsOfWinning")
            OfferDaysLeft = DateDiff("d", Today, OfferEnd)
            CustomerGroupID = row2.Item("CustomerGroupID")
            RewardGroupID = MyCommon.NZ(row2.Item("RewardGroup"), -1)
            AllowOptOut = MyCommon.NZ(row2.Item("AllowOptOut"), False)
            EmployeesOnly = IIf(MyCommon.NZ(row2.Item("EmployeesOnly"), False), 1, 0)
            EmployeesExcluded = IIf(MyCommon.NZ(row2.Item("EmployeesExcluded"), False), 1, 0)
            
            ROID = MyCommon.NZ(row2.Item("RewardOptionId"), -1)
            
            ' Filter out Offers where the customer is already a member of the reward group
            MyCommon.QueryStr = "select MembershipID from GroupMembership where CustomerPK=" & CustomerPK & " " & _
                                "and Deleted=0 and CustomerGroupID=" & RewardGroupID & ";"
            rstWeb = MyCommon.LXS_Select
            OptOutOffer = (rstWeb.Rows.Count > 0)
            
            ' Check if the customer is in a group that is excluded from this offer
            MyCommon.QueryStr = "select CustomerGroupID from CPE_ST_IncentiveCustomerGroups ICG " & _
                                "where RewardOptionID=" & ROID & " and CustomerGroupID in (" & CustomerGroups.ToString & ") and ExcludedUsers=1 and Deleted=0;"
            rstWeb = MyCommon.LRT_Select
            ExcludedFromOffer = (rstWeb.Rows.Count > 0)
            
            ' Check if the customer meets any points condition, if applicable, that exist for this offer
            MyCommon.QueryStr = "select QtyForIncentive, ProgramID from CPE_ST_IncentivePointsGroups where RewardOptionID=" & ROID & " and Deleted=0;"
            rstWeb = MyCommon.LRT_Select
            If (rstWeb.Rows.Count > 0) Then
              ' check if the customer has enough points in the program
              QtyRequired = MyCommon.NZ(rstWeb.Rows(0).Item("QtyForIncentive"), 0)
              MyCommon.QueryStr = "select Amount from Points where ProgramID=" & MyCommon.NZ(rstWeb.Rows(0).Item("ProgramID"), -1) & " and CustomerPK=" & CustomerPK & ";"
              rstWeb = MyCommon.LXS_Select
              If (rstWeb.Rows.Count > 0) Then
                PointsConditionOK = (MyCommon.NZ(rstWeb.Rows(0).Item("Amount"), 0) >= QtyRequired)
              Else
                ' no points entry for this customer
                PointsConditionOK = False
              End If
            Else
              ' no points condition exist for this web offer
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
              
              'Finally we have all the info we need, so display it.
              If (Transformation = "XML") Then
                retstring.Append("  <Offer>" & vbCrLf)
                retstring.Append("    <OfferID>" & OfferID & "</OfferID>" & vbCrLf)
                retstring.Append("    <Name>" & OfferName & "</Name>" & vbCrLf)
                retstring.Append("    <Description>" & OfferDesc & "</Description>" & vbCrLf)
                retstring.Append("    <StartDate>" & OfferStart & "</StartDate>" & vbCrLf)
                retstring.Append("    <EndDate>" & OfferEnd & "</EndDate>" & vbCrLf)
                retstring.Append("    <AllowOptOut>" & IIf(AllowOptOut, 1, 0) & "</AllowOptOut>" & vbCrLf)
                retstring.Append("    <EmployeesOnly>" & EmployeesOnly & "</EmployeesOnly>" & vbCrLf)
                retstring.Append("    <EmployeesExcluded>" & EmployeesExcluded & "</EmployeesExcluded>" & vbCrLf)
                retstring.Append("    <BodyText>" & PrintedMessage & "</BodyText>" & vbCrLf)
                retstring.Append("    <Graphic>" & GraphicsFileName & "</Graphic>" & vbCrLf)
                retstring.Append("  </Offer>" & vbCrLf)
              ElseIf (Transformation = "HTML") Then
                retstring.Append("<div class=""group"" id=""group" & OfferID & """>" & vbCrLf)
                retstring.Append("  <h3 class=""name"" alt=""Offer# " & OfferID & ", CGroup# " & CustomerGroupID & """ title=""Offer " & OfferID & ", CGroup " & CustomerGroupID & """>" & OfferName & "</h3>" & vbCrLf)
                GraphicsFileName = RetrieveGraphicPath(OfferID)
                If (GraphicsFileName <> "") Then
                  retstring.Append("    <img src=""images\" & GraphicsFileName & """ align=""right"" />")
                End If
                If (OptOutOffer AndAlso AllowOptOut) Then
                  retstring.Append("  <form action=""home.aspx"" method=""post"" name=""OptOut"">" & vbCrLf)
                  retstring.Append("    <input type=""hidden"" id=""opt-type"" name=""opt-type"" value=""out"" />")
                  retstring.Append("    <input type=""hidden"" id=""opt-extid"" name=""opt-extid"" value=""" & ExtID & """ />" & vbCrLf)
                  retstring.Append("    <input type=""hidden"" id=""opt-cardtypeid"" name=""opt-cardtypeid"" value=""" & CustomerTypeID & """ />" & vbCrLf)
                  retstring.Append("    <input type=""hidden"" id=""opt-cgroupid"" name=""opt-cgroupid"" value=""" & RewardGroupID & """ />" & vbCrLf)
                                    retstring.Append("    <input type=""hidden"" id=""opt-rgroupid"" name=""opt-rgroupid"" value=""" & RewardGroupID & """ />" & vbCrLf)
                  retstring.Append("    <input type=""hidden"" id=""opt-offerid"" name=""opt-offerid"" value=""" & OfferID & """ />" & vbCrLf)
                  retstring.Append("    <input class=""optout"" type=""submit"" value=""X OPT OUT"" />" & vbCrLf)
                  retstring.Append("  </form>" & vbCrLf)
                Else
                  retstring.Append("  <form action=""home.aspx"" method=""post"" name=""OptIn"">" & vbCrLf)
                  retstring.Append("    <input type=""hidden"" id=""opt-type"" name=""opt-type"" value=""in"" />")
                  retstring.Append("    <input type=""hidden"" id=""opt-extid"" name=""opt-extid"" value=""" & ExtID & """ />" & vbCrLf)
                  retstring.Append("    <input type=""hidden"" id=""opt-cardtypeid"" name=""opt-cardtypeid"" value=""" & CustomerTypeID & """ />" & vbCrLf)
                  retstring.Append("    <input type=""hidden"" id=""opt-cgroupid"" name=""opt-cgroupid"" value=""" & CustomerGroupID & """ />" & vbCrLf)
                  retstring.Append("    <input type=""hidden"" id=""opt-rgroupid"" name=""opt-rgroupid"" value=""" & RewardGroupID & """ />" & vbCrLf)
                  retstring.Append("    <input type=""hidden"" id=""opt-allowoptout"" name=""opt-allowoptout"" value=""" & AllowOptOut & """ />" & vbCrLf)
                  retstring.Append("    <input class=""optin"" type=""submit"" value=""+ JOIN"" />" & vbCrLf)
                  retstring.Append("  </form>" & vbCrLf)
                End If
                retstring.Append("  <div class=""description"">" & vbCrLf)
                retstring.Append("    " & OfferDesc & vbCrLf)
                retstring.Append(" (<a href=""javascript:openNamedPopup('prntmsg.aspx?identifier=" & Identifier & "&OfferID=" & OfferID & "&cg=" & RewardGroupID & "', 'PrntMsg');"">Details</a>)" & vbCrLf)
                retstring.Append("  </div>" & vbCrLf)
                retstring.Append("  <div class=""valid"">" & vbCrLf)
                retstring.Append("    Valid <span class=""startdate"">" & OfferStart & "</span> - <span class=""enddate"">" & OfferEnd & "</span><br />" & vbCrLf)
                If (OfferDaysLeft > 1) Then
                  retstring.Append("    It will expire in " & OfferDaysLeft & " days.<br />")
                ElseIf (OfferDaysLeft = 1) Then
                  retstring.Append("    It will expire tomorrow.<br />")
                ElseIf (OfferDaysLeft = 0) Then
                  retstring.Append("    It expires today.<br />")
                ElseIf (OfferDaysLeft = -1) Then
                  retstring.Append("    It expired yesterday.<br />")
                ElseIf (OfferDaysLeft < -1) Then
                  retstring.Append("    It expired " & Math.Abs(OfferDaysLeft) & " days ago.<br />")
                End If
                retstring.Append("  </div>")
                Send("    <br class=""half"" />")
                If (OfferOdds > 0) Then
                  Send("    <b>Odds of winning:</b> 1:" & OfferOdds & "<br />")
                End If
                retstring.Append("</div>")
                retstring.Append("<br />")
                retstring.Append("<hr />")
                retstring.Append("")
              ElseIf (Transformation = "nothing") Then
              Else
              End If
            End If
          Next
        End If
      End If
    End If
    
    Return retstring.ToString
  End Function
  
  Function sendCustInfo(ByVal CustomerPK As String, ByVal Transformation As String)
    Dim retString As New StringWriter
    Dim rst As DataTable
    Dim PrimaryExtID As String

    MyCommon.Open_LogixXS()
    MyCommon.QueryStr = "select C.CustomerPK, C.FirstName, C.LastName, C.Employee, C.CustomerStatusID, C.CurrYearSTD, C.LastYearSTD, " & _
                        "C.CustomerTypeID, C.HHPK, E.Address, E.City, E.State, E.Zip, E.Country, E.PhoneAsEntered as Phone, E.Email, E.DOB " & _
                        "from Customers as C with (NoLock) " & _
                        "left join CustomerEXT as E on C.CustomerPK=E.CustomerPK " & _
                        "where C.CustomerPK=" & CustomerPK & ";"
    rst = MyCommon.LXS_Select()
    If rst.Rows.Count > 0 Then
      PrimaryExtID = GetCardNumber(Long.Parse(CustomerPK), MyCommon.NZ(rst.Rows(0).Item("CustomerTypeID"), 0))
      If (Transformation = "XML") Then
        retString.WriteLine("<Customer Editable=""true"" >")
        retString.WriteLine("  <PrimaryExtID>" & PrimaryExtID & "</PrimaryExtID>")
        retString.WriteLine("  <FirstName>" & MyCommon.NZ(rst.Rows(0).Item("FirstName"), "") & "</FirstName>")
        retString.WriteLine("  <LastName>" & MyCommon.NZ(rst.Rows(0).Item("LastName"), "") & "</LastName>")
        retString.WriteLine("  <Employee>" & MyCommon.NZ(rst.Rows(0).Item("Employee"), "") & "</Employee>")
        retString.WriteLine("  <CardStatusID>" & MyCommon.NZ(rst.Rows(0).Item("CustomerStatusID"), "") & "</CardStatusID>")
        retString.WriteLine("  <CurrYearSTD>" & MyCommon.NZ(rst.Rows(0).Item("CurrYearSTD"), "") & "</CurrYearSTD>")
        retString.WriteLine("  <LastYearSTD>" & MyCommon.NZ(rst.Rows(0).Item("LastYearSTD"), "") & "</LastYearSTD>")
        retString.WriteLine("  <CustomerTypeID>" & MyCommon.NZ(rst.Rows(0).Item("CustomerTypeID"), "") & "</CustomerTypeID>")
        retString.WriteLine("  <HHPK>" & MyCommon.NZ(rst.Rows(0).Item("HHPK"), "") & "</HHPK>")
        retString.WriteLine("  <Address>" & MyCommon.NZ(rst.Rows(0).Item("Address"), "") & "</Address>")
        retString.WriteLine("  <City>" & MyCommon.NZ(rst.Rows(0).Item("City"), "") & "</City>")
        retString.WriteLine("  <State>" & MyCommon.NZ(rst.Rows(0).Item("State"), "") & "</State>")
        retString.WriteLine("  <Zip>" & MyCommon.NZ(rst.Rows(0).Item("Zip"), "") & "</Zip>")
        retString.WriteLine("  <Country>" & MyCommon.NZ(rst.Rows(0).Item("Country"), "") & "</Country>")
        retString.WriteLine("  <Phone>" & MyCommon.NZ(rst.Rows(0).Item("Phone"), "") & "</Phone>")
        retString.WriteLine("  <Email>" & MyCommon.NZ(rst.Rows(0).Item("Email"), "") & "</Email>")
        retString.WriteLine("</Customer>")
      ElseIf (Transformation = "HTML") Then
        retString.WriteLine("<h2>Your details</h2>")
        retString.WriteLine("<div class=""box"" id=""detailsbox"">")
        retString.WriteLine("  <div id=""name"">")
        retString.WriteLine("      " & MyCommon.NZ(rst.Rows(0).Item("FirstName"), "") & " " & MyCommon.NZ(rst.Rows(0).Item("LastName"), ""))
        retString.WriteLine("  </div>")
        retString.WriteLine("  <div id=""contact"">")
        retString.WriteLine("    <div id=""address"">")
        retString.WriteLine("      " & MyCommon.NZ(rst.Rows(0).Item("Address"), ""))
        retString.WriteLine("    </div>")
        retString.WriteLine("    <div id=""city"">")
        retString.WriteLine("      " & MyCommon.NZ(rst.Rows(0).Item("City"), ""))
        retString.WriteLine("    </div>")
        retString.WriteLine("    <div id=""state"">")
        retString.WriteLine("      " & MyCommon.NZ(rst.Rows(0).Item("State"), ""))
        retString.WriteLine("    </div>")
        retString.WriteLine("    <div id=""zip"">")
        retString.WriteLine("      " & MyCommon.NZ(rst.Rows(0).Item("Zip"), ""))
        retString.WriteLine("    </div>")
        retString.WriteLine("    <div id=""country"">")
        retString.WriteLine("      " & MyCommon.NZ(rst.Rows(0).Item("Country"), ""))
        retString.WriteLine("    </div>")
        retString.WriteLine("    <div id=""email"">")
        If MyCommon.NZ(rst.Rows(0).Item("Email"), "") <> "" Then
          retString.WriteLine("      <a href=""mailto:" & MyCommon.NZ(rst.Rows(0).Item("Email"), "") & """>" & MyCommon.NZ(rst.Rows(0).Item("Email"), "") & "</a>")
        End If
        retString.WriteLine("    </div>")
        Dim PhoneNumber As String = MyCommon.NZ(rst.Rows(0).Item("Phone"), "")
        retString.WriteLine("    <div id=""phone"">")
        If (PhoneNumber = "") Then
        Else
          retString.WriteLine("      " & PhoneNumber(0))
        End If
        retString.WriteLine("    </div>")
        Dim DateOfBirth As String = MyCommon.NZ(rst.Rows(0).Item("DOB"), "")
        Dim DOBParts() As String = {"", "", ""}
        If (DateOfBirth IsNot Nothing) Then
          Select Case DateOfBirth.Length
            Case 4
              DOBParts(0) = ""
              DOBParts(1) = ""
              DOBParts(2) = DateOfBirth
            Case 8
              DOBParts(0) = DateOfBirth.Substring(0, 2)
              DOBParts(1) = DateOfBirth.Substring(2, 2)
              DOBParts(2) = DateOfBirth.Substring(4)
          End Select
        End If
        retString.WriteLine("    <div id=""dob"">")
        If (DOBParts(0) = "") AndAlso (DOBParts(1) = "") AndAlso (DOBParts(2) = "") Then
        Else
          retString.WriteLine("      " & DOBParts(0) & "/" & DOBParts(1) & "/" & DOBParts(2))
        End If
        retString.WriteLine("    </div>")
        retString.WriteLine("  </div>")
        retString.WriteLine("  <div id=""identifiers"">")
        retString.WriteLine("    <div id=""primaryextid"">")
        retString.WriteLine("      Your NCR Program Card is")
        retString.WriteLine("      <span class=""cardnumber"">" & PrimaryExtID & "</span>")
        retString.WriteLine("    </div>")
        retString.WriteLine("    <div id=""employeestatus"">")
        retString.WriteLine("      You are an employee.")
        retString.WriteLine("    </div>")
        retString.WriteLine("    <div style=""display:none;"">")
        retString.WriteLine("      " & CustomerPK)
        retString.WriteLine("    </div>")
        retString.WriteLine("  </div>")
        retString.WriteLine("  <div id=""savings"">")
        retString.WriteLine("    <div id=""curryearstd"">")
        retString.WriteLine("      This year you've saved")
        retString.WriteLine("      <span class=""stdvalue"">$" & MyCommon.NZ(rst.Rows(0).Item("CurrYearSTD"), "") & "</span>.")
        retString.WriteLine("    </div>")
        retString.WriteLine("    <div id=""lastyearstd"">")
        retString.WriteLine("      Last year you saved")
        retString.WriteLine("      <span class=""stdvalue"">$" & MyCommon.NZ(rst.Rows(0).Item("LastYearSTD"), "") & "</span>.")
        retString.WriteLine("    </div>")
        retString.WriteLine("  </div>")
        retString.WriteLine("  <br />")
        retString.WriteLine("  <form method=""post"" action=""#"" id=""editform"" name=""editform"">")
        retString.WriteLine("    <input type=""button"" class=""medium"" onclick=""javascript:openMiniPopup('edit.aspx?PrimaryExtID=" & PrimaryExtID & "&CustPK=" & MyCommon.NZ(rst.Rows(0).Item("CustomerPK"), "") & "');"" value=""Edit your details"" id=""edit"" name=""edit""/>")
        retString.WriteLine("  </form>")
		retString.WriteLine("  <form method=""post"" action=""#"" id=""editlistform"" name=""editlistform"">")
        retString.WriteLine("    <input type=""button"" class=""medium"" onclick=""javascript:openMiniPopup('list.aspx?CustPK=" & MyCommon.NZ(rst.Rows(0).Item("CustomerPK"), "") & "');"" value=""Edit your shopping list"" id=""edit"" name=""edit""/>")
        retString.WriteLine("  </form>")
        retString.WriteLine("</div>")
        retString.WriteLine("<div class=""box"" id=""logoutbox"">")
        retString.WriteLine("  Protect your privacy by logging out when finished.")
        retString.WriteLine("  <form action=""index.html"" id=""logoutform"" name=""logoutform"" target=""_top"">")
        retString.WriteLine("    <input id=""logout"" name=""logout"" class=""medium"" type=""submit"" value=""Logout"" />")
        retString.WriteLine("  </form>")
        retString.WriteLine("</div>")
      ElseIf (Transformation = "nothing") Then
      Else
      End If
    End If
    
    Return retString.ToString
  End Function
  
  Function Send_XMLTransactions(ByRef Identifier As String, ByVal CurrentOffers As String, ByVal Transformation As String)
    Dim rst As DataTable
    Dim row As DataRow
    Dim rst2 As DataTable
    Dim row2 As DataRow
    Dim rst3 As DataTable
    Dim transCt As Integer = 0
    Dim transOffers As StringBuilder
    Dim transRdmptAmt As StringBuilder
    Dim transRdmptCt As StringBuilder
    Dim startPosition As Integer
    Dim endPosition As Integer
    Dim UnknownPhrase As String = ""
    Dim PrimaryExtID As String = ""
    Dim OfferName As String = ""
    Dim XID As String = ""
    Dim TotalRedeemCt As Integer = 0
    Dim TotalRedeemAmt As Double = 0.0
    Dim retstring As New StringBuilder
    
    MyCommon.Open_LogixRT()
    MyCommon.Open_LogixWH()
    MyCommon.Open_LogixXS()
    
    MyCommon.QueryStr = "select top 1 ExtCardID from CardIDs with (NoLock) where CustomerPK='" & Identifier & "';"
    rst = MyCommon.LXS_Select
    If (rst.Rows.Count > 0) Then
      PrimaryExtID = MyCommon.NZ(rst.Rows(0).Item("ExtCardID"), "")
    End If
    
    retstring.Append("<table summary=""Transactions"" style=""width:535px;"">")
    retstring.Append("  <thead>")
    retstring.Append("    <tr>")
    retstring.Append("      <th align=""left"" class=""th-datetime"" scope=""col"">Date</a></th>")
    retstring.Append("      <th align=""left"" class=""th-id"" scope=""col"">Store</a>")
    retstring.Append("      <th align=""center"" class=""th-transaction"" scope=""col"" style=""text-align: center;"">Transaction#</a></th>")
    retstring.Append("      <th align=""right"" class=""th-amount"" scope=""col"" style=""text-align: right;"">Amount</a></th>")
    retstring.Append("    </tr>")
    retstring.Append("  </thead>")
    retstring.Append("  <tbody>")
    
    MyCommon.QueryStr = "select Top " & MyCommon.Fetch_WebOption(7) & " CustomerPrimaryExtID, Max(TransDate) as TransactionDate, ExtLocationCode, sum(RedemptionAmount) as RedemptionAmount, " & _
                        "sum(RedemptionCount) as RedemptionCount, TerminalNum, TransNum, count(*) as DetailRecords " & _
                        "from TransRedemptionView with (NoLock) where CustomerPrimaryExtID in ('" & PrimaryExtID & "') " & _
                        "group by CustomerPrimaryExtID, TransNum, TerminalNum, ExtLocationCode " & _
                        "order by TransactionDate DESC;"
    rst = MyCommon.LWH_Select
        
    If rst.Rows.Count > 0 Then
      For Each row In rst.Rows
        transCt += 1
        retstring.Append("<tr>")
        retstring.Append("  <td>")
        If MyCommon.NZ(row.Item("TransactionDate"), "1/1/1980") = "1/1/1980" Then
          retstring.Append("?")
        Else
          retstring.Append(Format(row.Item("TransactionDate"), "dd MMM yyyy, HH:mm:ss"))
        End If
        retstring.Append("</td>")
        retstring.Append("  <td>" & MyCommon.NZ(row.Item("ExtLocationCode"), UnknownPhrase) & "</td>")
        retstring.Append("  <td style=""text-align:center;word-break:break-all"">" & MyCommon.NZ(row.Item("TransNum"), UnknownPhrase) & "</td>")
        retstring.Append("  <td style=""text-align:right;"">" & MyCommon.NZ(row.Item("RedemptionAmount"), UnknownPhrase) & "</td>")
        retstring.Append("</tr>")
        ' Write detail rows
        retstring.Append("<tr id=""trTrans" & transCt & """ style=""color:#888888;"">")
        retstring.Append("  <td></td>")
        MyCommon.QueryStr = "select OfferID, RedemptionAmount, RedemptionCount from TransRedemptionView with (NoLock) " & _
                            "where CustomerPrimaryExtID in ('" & PrimaryExtID & "') " & " and TransNum='" & MyCommon.NZ(row.Item("TransNum"), UnknownPhrase) & "';"
        rst2 = MyCommon.LWH_Select
        If (rst2.Rows.Count > 0) Then
          transOffers = New StringBuilder(500)
          transRdmptAmt = New StringBuilder(100)
          transRdmptCt = New StringBuilder(100)
          For Each row2 In rst2.Rows
            MyCommon.QueryStr = "select Name as OfferName, ExtOfferID as XID from CM_ST_Offers with (NoLock) where OfferID=" & MyCommon.NZ(row2.Item("OfferID"), -1) & _
                                " union all " & _
                                "select IncentiveName as OfferName, ClientOfferID as XID from CPE_ST_Incentives with (NoLock) where IncentiveID=" & MyCommon.NZ(row2.Item("OfferID"), -1) & ";"
            rst3 = MyCommon.LRT_Select
            If (rst3.Rows.Count > 0) Then
              OfferName = MyCommon.NZ(rst3.Rows(0).Item("OfferName"), "")
              XID = MyCommon.NZ(rst3.Rows(0).Item("XID"), "[None]")
            End If
            transOffers.Append("Offer #" & MyCommon.NZ(row2.Item("OfferID"), "") & ": """ & OfferName & """<br />")
            transRdmptAmt.Append(MyCommon.NZ(row2.Item("RedemptionAmount"), "") & "<br />")
            transRdmptCt.Append(MyCommon.NZ(row2.Item("RedemptionCount"), "") & "<br />")
          Next
          retstring.Append("  <td colspan=""2"">" & transOffers.ToString & "</td>")
          retstring.Append("  <td style=""text-align:right;"">" & transRdmptAmt.ToString & "</td>")
        End If
        retstring.Append("</tr>")
        TotalRedeemAmt += MyCommon.NZ(row.Item("RedemptionAmount"), 0.0)
        TotalRedeemCt += MyCommon.NZ(row.Item("RedemptionCount"), 0)
      Next
    Else
      retstring.Append("<tr>")
      retstring.Append("  <td colspan=""7"" style=""text-align:center""><i>No transaction history</i></td>")
      retstring.Append("</tr>")
    End If
    retstring.Append("  </tbody>")
    retstring.Append("</table>")
    
    Return retstring.ToString
    
    MyCommon.Close_LogixRT()
    MyCommon.Close_LogixWH()
    MyCommon.Close_LogixXS()
    
  End Function
  
  Function sendOptAck(ByVal PrimaryExtID As String, ByVal CustomerGroupID As String, ByVal OptType As String, ByVal Successful As Boolean, ByVal Transformation As String)
    Dim retString As New StringWriter
    
    If (Transformation = "XML") Then
      retString.WriteLine("<Acknowledgment>")
      retString.WriteLine("  <PrimaryExtID>" & PrimaryExtID & "</PrimaryExtID>")
      retString.WriteLine("  <CustomerGroupID>" & CustomerGroupID & "</CustomerGroupID>")
      If OptType = "out" Then
        If Successful Then
          retString.WriteLine("  <OptOut>1</OptOut>")
        Else
          retString.WriteLine("  <OptOut>0</OptOut>")
        End If
      ElseIf OptType = "in" Then
        If Successful Then
          retString.WriteLine("  <OptIn>1</OptIn>")
        Else
          retString.WriteLine("  <OptIn>0</OptIn>")
        End If
      End If
      retString.WriteLine("</Acknowledgment>")
    End If
    
    Return retString.ToString
    
  End Function
  
  Private Function RetrieveAccumulationBalance(ByVal OfferID As Long, ByVal CustomerPK As Long) As String
    Dim rst As System.Data.DataTable
    Dim row As System.Data.DataRow
    Dim EngineID As Integer = -1
    Dim AccumProgram As Boolean = False
    Dim RewardOptionID As Long = -1
    Dim HHEnable As Boolean = False
    Dim UnitType As Integer = 0
    Dim HouseholdPK As Integer = 0
    Dim TotalAccum As Double
    Dim AccumulationBalance As String = ""
    
    MyCommon.QueryStr = "select EngineID from OfferIDs where OfferID = " & OfferID
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      EngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), -1)
    End If
    
    ' ~~~~~~~~~~~~~~~~~
    If EngineID = 2 Then
      ' First find the accumulation info
      MyCommon.QueryStr = "select IPG.AccumMin, RO.RewardOptionID, RO.HHEnable, IPG.QtyUnitType " & _
                          "from CPE_ST_IncentiveProductGroups as IPG " & _
                          "inner join CPE_ST_RewardOptions as RO on IPG.RewardOptionID=RO.RewardOptionID and IPG.Deleted=0 and IPG.ExcludedProducts=0 and RO.Deleted=0 " & _
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
          MyCommon.QueryStr = "select HHPK from Customers where CustomerPK='" & CustomerPK & "';"
          rst = MyCommon.LXS_Select
          If (rst.Rows.Count > 0) Then
            HouseholdPK = MyCommon.NZ(rst.Rows(0).Item("HHPK"), 0)
          End If
        End If
      End If
      
      If AccumProgram Then
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
          If UnitType = 1 Then
            AccumulationBalance = Format(TotalAccum, "###,##0")
          ElseIf UnitType = 2 Then
            AccumulationBalance = "$" & Format(TotalAccum, "###,##0.00")
          ElseIf UnitType = 3 Then
            AccumulationBalance = Format(TotalAccum, "###,##0.000")
          End If
        End If
      End If
      
    End If
    ' ~~~~~~~~~~~~~~~~~
    
    Return AccumulationBalance
  End Function
  
  Private Function RetrieveGraphicPath(ByVal OfferID As Long) As String
    Dim GraphicsFileName As String = ""
    Dim GraphicsFilePath As String = ""
    Dim GraphicsNewFilePath As String = ""
    Dim rst As System.Data.DataTable
    
    Try
      ' Find if a graphic is assigned to this offer
      MyCommon.QueryStr = "select OnScreenAdID, Name, ImageType, Width, Height from OnScreenAds where Deleted=0 and OnScreenAdID in " & _
                          " (select OutputID from CPE_ST_Deliverables where Deleted=0 and DeliverableTypeID=1 and RewardOptionPhase=1 and RewardOptionID in " & _
                          "  (select RewardOptionID from CPE_ST_RewardOptions where IncentiveID=" & OfferID & " and TouchResponse=0 and Deleted=0));"
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
          GraphicsNewFilePath = Server.MapPath("home.aspx")
          GraphicsNewFilePath = GraphicsNewFilePath.Substring(0, GraphicsNewFilePath.LastIndexOf("\"))
          GraphicsNewFilePath += "\images\" & GraphicsFileName
          If Not (File.Exists(GraphicsNewFilePath)) Then
            File.Copy(GraphicsFilePath, GraphicsNewFilePath)
            If Not (File.Exists(GraphicsNewFilePath)) Then
              GraphicsFileName = ""
            End If
          End If
        End If
      End If
    Catch ex As Exception
      GraphicsFileName = ""
    End Try
    
    Return GraphicsFileName
  End Function
  
  Private Function RetrievePrintedMessage(ByVal OfferID As Long) As String
    Dim PMsgBuf As New StringBuilder()
    Dim rst As System.Data.DataTable
    
    MyCommon.Open_LogixRT()
    MyCommon.QueryStr = "select PMTypes.Description, PMTiers.TierLevel, PMTiers.BodyText " & _
                        "from CPE_ST_PrintedMessages as PM with (NoLock) " & _
                        "inner join CPE_ST_PrintedMessageTiers as PMTiers with (NoLock) on PM.MessageID=PMTiers.MessageID " & _
                        "inner join PrintedMessageTypes as PMTypes with (NoLock) on PM.MessageTypeID=PMTypes.TypeID " & _
                        "inner join CPE_ST_Deliverables as D with (NoLock) on D.OutputID=PM.MessageID " & _
                        "inner join CPE_ST_RewardOptions as RO with (NoLock) on RO.RewardOptionID=D.RewardOptionID " & _
                        "where RO.IncentiveID=" & OfferID & " and RO.Deleted=0 and D.Deleted=0 and D.RewardOptionPhase=1 and DeliverableTypeID=4;"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      PMsgBuf.Append(MyCommon.NZ(rst.Rows(0).Item("BodyText"), ""))
    End If
    MyCommon.Close_LogixRT()
    
    Return PMsgBuf.ToString()
  End Function
  
  Private Function GenerateGroupOffers(ByVal GroupID As Integer) As String
    Dim OfferString As String = ""
    
    Return OfferString
  End Function
  
  Private Function GenerateMessagePreview(ByVal newText As String) As String
    Dim PrintLine As String
    Dim Found As Boolean = False
    Dim ExtFound As Boolean = False
    Dim NewLine As String
    Dim NewBRLine As String
    Dim y As Integer = 0
    Dim FormatLine As String
    Dim WrappedPrintLine As String
    Dim RawTextLine As String
    Dim FormatChar As String
    Dim ExtFormatChar As String
    Dim ENDMESSAGE As String
    Dim RawMessage As String
    Dim MESSAGE As String
    Dim x As Integer
    Dim rst As System.Data.DataTable
    Dim row As System.Data.DataRow
    Dim PrinterTag As String
    Dim PreviewText As String
    
    ' This 2D array will hold the replacement for real-time engine tags; just add or modify
    Dim SubTags(,) As String = { _
                                {"|CUSTOMERID|", "|TSD|", "|LYTS|", "|CURRDATE|", "|OFFERSTART|", "|OFFEREND|", "|TOTALPOIINTS|", "|ACCUMANT|", "|REMAINAMT|"}, _
                                {"###################", "000.00", "000.00", "xx/xx/xxxx", "xx/xx/xxxx", "xx/xx/xxxx", "xx", "000.00", "000.00"}}
    
    RawMessage = newText
    ' Here we will format the message (which is currently in RawMessage).
    ' First, look for and replace CUSTOMERID TSD LYTS CURRDATE OFFERSTART OFFEREND TOTALPOIINTS ACCUMANT REMAINAMT 
    MESSAGE = RawMessage
    For x = 0 To SubTags.GetUpperBound(1)
      MESSAGE = Replace(MESSAGE, SubTags(0, x), SubTags(1, x))
    Next
    
    Dim strArray() As String = MESSAGE.Split(vbLf)
    
    ' Blank out MESSAGE so we can refill it
    MESSAGE = ""
    
    ' Get the tags
    MyCommon.QueryStr = "select PrinterTypeID, '|' + tag + '|' as tag,isnull(pt.previewchars,'') as previewchars from markuptags as MT left join printertranslation as pt on MT.markupid=pt.markupid"
    rst = MyCommon.LRT_Select
    
    For Each PrintLine In strArray
      ' Search on each printer tag to replace
      ' Pull out details (name, page width, etc.) for that printer
      RawTextLine = PrintLine
      
      For Each row In rst.Rows
        PrinterTag = row.Item("tag")
        PreviewText = row.Item("previewchars")
        If (InStr(RawTextLine, PrinterTag) And row.Item("PrinterTypeID") = 4) Then
          RawTextLine = Replace(RawTextLine, PrinterTag, PreviewText)
          ENDMESSAGE = "</span>"
          'debugging = debugging & "ext replacing " & PrinterTag & "<br />"
        End If
      Next
      
      ' Get any remaining tags
      For Each row In rst.Rows
        PrinterTag = row.Item("tag")
        If (InStr(RawTextLine, PrinterTag)) Then
          RawTextLine = Replace(RawTextLine, PrinterTag, "")
        End If
      Next
      
      MESSAGE = MESSAGE & RawTextLine & ENDMESSAGE & "<br />"
      ENDMESSAGE = ""
    Next
    
    Return MESSAGE
  End Function
      
  Private Function GetCardNumber(ByVal CustomerPK As Long, ByVal CardTypeID As Integer) As String
    Dim CardNum As String = ""
    Dim rst As DataTable
    
    'The identifier the customer supplied is all numbers, so assume it's a card number
    MyCommon.QueryStr = "select top 1 ExtCardID from CardIDs with (NoLock) where CustomerPK=" & CustomerPK & _
                        "  and CardTypeID=" & CardTypeID
    rst = MyCommon.LXS_Select
    If rst.Rows.Count > 0 Then
      CardNum = MyCommon.NZ(rst.Rows(0).Item("ExtCardID"), "")
    End If

    Return CardNum
  End Function
</script>
