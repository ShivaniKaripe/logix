<%@ Page Language="vb" Debug="true" CodeFile="ncr-cwCB.vb" Inherits="cwCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Xml.Xsl" %>
<%@ Import Namespace="System.Xml.XPath" %>
<%
  Dim CopientFileName As String = "ncr-cwfeed.aspx"
  Dim CopientFileVersion As String = "7.3.1.138972"
  Dim CopientProject As String = "Copient Logix"
  Dim CopientNotes As String = ""

  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim ErrorString As String = ""
  Dim Transform As String = ""

  ' AdminUserID = Verify_AdminUser(Logix)
  MyCommon.AppName = "ncr-cwfeed.aspx"
  MyCommon.Open_LogixRT()

  If (LanguageID = 0) Then
    LanguageID = MyCommon.Extract_Val(Request.QueryString("LanguageID"))
  End If

  If (Request.Form("mode") <> "") Then
    If (Session("customerpk") Is Nothing) Then
      ErrorString = "<b>Session not valid</b>"
    Else
      Transform = Request.Form("Transform")
      If (Request.Form("mode") = "cust" And Session("customerpk").ToString <> "") Then
        custOffers(Session("customerpk").ToString, Transform)
      ElseIf (Request.Form("mode") = "info" And Session("customerpk").ToString <> "") Then
        custInfo(Session("customerpk").ToString, Transform)
      Else
        ErrorString = "<b>Invalid session mode</b>"
      End If
    End If
  Else
    ErrorString = "<b>" & Copient.PhraseLib.Lookup("term.noarguments", LanguageID) & "</b>"
  End If

  If (ErrorString <> "") Then
    Send(ErrorString)
  End If

  MyCommon.Close_LogixRT()
  Logix = Nothing
  MyCommon = Nothing
%>
<script runat="server">
  Public DefaultLanguageID
  Public MyCommon As New Copient.CommonInc

  '----------------------------------------------------------------------------------

  Sub custInfo(ByVal identifier As String, ByVal XFORM As String)
    Dim outStr As String
    Dim outputStr As New StringWriter

    Response.Cache.SetCacheability(HttpCacheability.NoCache)

    If (XFORM = "XML") Then
      outputStr.WriteLine(xmlHeader())
    End If

    outStr = sendCustInfo(identifier, XFORM)

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

  Sub custOffers(ByVal id As String, ByVal XFORM As String)
    Dim outSTr As String
    Dim outputStr As New StringWriter
    Dim currentoffers As String = ""
    Dim outStrOffers As String = ""
    Dim outStrGroups As String = ""

    Response.Cache.SetCacheability(HttpCacheability.NoCache)
    MyCommon.Open_LogixRT()

    If (XFORM = "XML") Then
      outputStr.WriteLine(xmlHeader())
    End If

    ' figure out the available offers
    outSTr = Send_XMLCurrentOffers(id, currentoffers, XFORM)
    outStrOffers = outSTr

    If (XFORM = "XML") Then
      If (outSTr.Length > 0) Then
        outputStr.WriteLine("<Offers>")
        outputStr.WriteLine(outSTr)
        outputStr.WriteLine("</Offers>")
      End If
    End If

    ' figure out the website offers to opt in
    outSTr = Send_XMLGroupOffers(id, currentoffers, XFORM)
    outStrGroups = outSTr

    If (XFORM = "XML") Then
      If (outSTr.Length > 0) Then
        outputStr.WriteLine("<Groups>")
        outputStr.WriteLine(outSTr)
        outputStr.WriteLine("</Groups>")
      End If

      outputStr.WriteLine(xmlFooter())

      Dim outString As String = outputStr.ToString

      If (XFORM) Then
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
      Send("    <h2>Offers</h2>")
      Send(outStrOffers)
      Send("  </div>")
      Send("")
      Send("  <div class=""gutter"">")
      Send("  </div>")
      Send("")
      Send("  <div id=""groups"">")
      Send("    <h2>Groups</h2>")
      Send(outStrGroups)
      Send("  </div>")
    End If

    MyCommon.Close_LogixRT()
  End Sub

  Function xmlHeader() As String
    ' Return "<?xml-stylesheet type=""text/xsl"" ?>" & _
    Return "<CustWeb CustomerPK=""123456789"" Mode=""Home"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">"
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

  Function Send_XMLCurrentOffers(ByVal identifier As String, ByRef currentoffers As String, ByVal transformation As String) As String
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
    Dim CustomerPK As Long
    Dim ExtID As String
    Dim Employee As Boolean
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
    Dim ProgramID As String = ""
    Dim ProgID As Integer
    Dim ProgramName As String = ""
    Dim Amount As Long
    Dim retString As New StringBuilder
    Dim CustomerGroups As String = "0"
    Dim CgBuf As New StringBuilder()

    MyCommon.Open_LogixRT()
    MyCommon.Open_LogixXS()

    'First off, we check to see if there's an identifier in the URL
    If (identifier <> "") Then
      ' There is, so first off we grab the customer's information
      If (IsNumeric(identifier)) Then
        'The identifier the customer supplied is all numbers, so assume it's a card number
        MyCommon.QueryStr = "SELECT C.CustomerPK,C.PrimaryExtID,C.Employee,C.CardStatusID,CE.Email " & _
                            "FROM Customers AS C " & _
                            "LEFT JOIN CustomerExt AS CE on CE.CustomerPK=C.CustomerPK " & _
                            "WHERE C.CustomerPK='" & identifier & "'"
      End If
      rst = MyCommon.LXS_Select
      If (rst.Rows.Count > 0) Then
        'A customer was found, so let's assign values to variables
        CustomerPK = rst.Rows(0).Item("CustomerPK")
        ExtID = rst.Rows(0).Item("PrimaryExtID")
        Employee = rst.Rows(0).Item("Employee")

        'Next, get the associated customer groups...
        MyCommon.QueryStr = "SELECT distinct CustomerGroupID FROM GroupMembership WHERE CustomerPK=" & CustomerPK & " and Deleted=0"
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
          MyCommon.QueryStr = "select distinct O.OfferID,O.ExtOfferID,O.IsTemplate,O.CMOADeployStatus,O.StatusFlag,O.OddsOfWinning,O.InstantWin, " & _
                              "O.Name,O.Description,O.ProdStartDate,O.ProdEndDate,LinkID,OID.EngineID from Offers as O " & _
                              "LEFT JOIN OfferConditions as OC on OC.OfferID=O.OfferID " & _
                              "INNER JOIN OfferIDs as OID on OID.OfferID=O.OfferID " & _
                              "where O.IsTemplate=0 and O.Deleted=0 and O.CMOADeployStatus=1 and OC.Deleted=0 and OC.ConditionTypeID=1 " & _
                              "and O.DisabledOnCFW=0 and ProdEndDate>'" & Today.AddDays(-1).ToString & "' and LinkID=" & row.Item("CustomerGroupID") & _
                              "union all " & _
                              "select distinct I.IncentiveID,I.ClientOfferID,I.IsTemplate,I.CPEOADeployStatus,I.StatusFlag,0 as OddsOfWinning,0 as InstantWin, " & _
                              "I.IncentiveName, Convert(nvarchar(2000),I.Description) as Description,I.StartDate,I.EndDate,ICG.CustomerGroupID,OID.EngineID from CPE_Incentives I " & _
                              "LEFT JOIN CPE_RewardOptions RO on I.IncentiveID=RO.IncentiveID and RO.TouchResponse=0 and RO.Deleted=0 " & _
                              "LEFT JOIN CPE_IncentiveCustomerGroups ICG on RO.RewardOptionID=ICG.RewardOptionID and ICG.ExcludedUsers = 0 " & _
                              "INNER JOIN OfferIDs as OID on OID.OfferID=I.IncentiveID " & _
                              "where (I.IsTemplate=0 and I.Deleted=0 and ICG.Deleted=0) " & _
                              "and I.DisabledOnCFW=0 and I.EndDate>'" & Today.AddDays(-1).ToString & "' and CustomerGroupID=" & row.Item("CustomerGroupID")
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
            CustomerGroupID = row2.Item("LinkID")

            ' Filter out the website offers
            MyCommon.QueryStr = "select OfferID from OfferIDs where OfferID=" & OfferID & " and EngineID=3;"
            rstWeb = MyCommon.LRT_Select

            ' Filter out the offers where the customer is in the excluded customer group
            MyCommon.QueryStr = "select ExcludedID from OfferConditions with (NoLock) " & _
                                "where OfferID = " & OfferID & " and ExcludedID in (" & CustomerGroups & ") " & _
                                "union " & _
                                "select CustomerGroupID as ExcludedID from CPE_IncentiveCustomerGroups ICG with (NoLock) " & _
                                "inner join CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID = ICG.RewardOptionID " & _
                                "where ICG.Deleted=0 and RO.Deleted=0 and RO.IncentiveID = " & OfferID & " and ExcludedUsers=1 " & _
                                "and CustomerGroupID in (" & CustomerGroups & ");"
            rstExcluded = MyCommon.LRT_Select

            If (rstWeb.Rows.Count = 0 AndAlso rstExcluded.Rows.Count = 0) Then

              'Find the name of the associated (and non-excluding) location group
              MyCommon.QueryStr = "SELECT OL.OfferID,OL.LocationGroupID,OL.Excluded,LG.Name FROM OfferLocations AS OL " & _
                                  "LEFT JOIN LocationGroups AS LG on LG.LocationGroupID=OL.LocationGroupID " & _
                                  "WHERE OL.OfferID=" & OfferID & " and OL.Excluded=0"
              rst3 = MyCommon.LRT_Select
              currentoffers += OfferID & ","

              'Find any associated points programs
              MyCommon.QueryStr = "select O.Offerid,Linkid,ProgramName,PP.ProgramID,PromoVarID from offerrewards as OFR with (NoLock) " & _
                                  "left join RewardPoints as RP with (NoLock) on RP.RewardPointsID=OFR.LinkID " & _
                                  "left join PointsPrograms as PP with (NoLock) on RP.ProgramID=PP.ProgramID " & _
                                  "left join Offers as O with (NoLock) on O.OfferID=OFR.OfferID " & _
                                  "where(rewardtypeid = 2 And O.deleted = 0 And OFR.deleted = 0) " & _
                                  "and RP.ProgramID is not null " & _
                                  "and O.Offerid=" & OfferID & _
                                  " union " & _
                                  "select " & OfferID & " as OfferID, D.OutputID, PP.ProgramName, PP.ProgramID, PP.PromoVarID " & _
                                  "from CPE_Deliverables D with (NoLock) inner join CPE_DeliverablePoints DP with (NoLock) on D.OutputID = DP.PKID " & _
                                  "inner join PointsPrograms PP with (NoLock) on DP.ProgramID = PP.ProgramID " & _
                                  "where D.RewardOptionID in (select RO.RewardOptionID from CPE_RewardOptions RO with (NoLock) where incentiveId=" & OfferID & ") " & _
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

              'So finally we have all the info we need. Now display it.
              If (transformation = "XML") Then
                retString.Append("<Offer>" & vbCrLf)
                retString.Append("  <OfferID>" & OfferID & "</OfferID>" & vbCrLf)
                retString.Append("  <Name>" & OfferName & "</Name>" & vbCrLf)
                retString.Append("  <Description>" & OfferDesc & "</Description>" & vbCrLf)
                retString.Append("  <StartDate>" & OfferStart & "</StartDate>" & vbCrLf)
                retString.Append("  <EndDate>" & OfferEnd & "</EndDate>" & vbCrLf)
                retString.Append("</Offer>" & vbCrLf)
              ElseIf (transformation = "HTML") Then
                retString.Append("<div class=""offer"" id=""offer" & OfferID & """>" & vbCrLf)
                retString.Append("  <h3 class=""name"" alt=""Offer# " & OfferID & ", CGroup# " & CustomerGroupID & """ title=""Offer " & OfferID & ", CGroup " & CustomerGroupID & """>" & OfferName & "</h3>" & vbCrLf)
                GraphicsFileName = RetrieveGraphicPath(OfferID)
                If (GraphicsFileName <> "") Then
                  retString.Append("    <img src=""images\" & GraphicsFileName & """ align=""right"" />")
                End If
                retString.Append("  <div class=""description"">" & vbCrLf)
                retString.Append("    " & OfferDesc & vbCrLf)
                PrintedMessage = RetrievePrintedMessage(OfferID)
                If (PrintedMessage <> "") Then
                  retString.Append(" (<a href=""javascript:openNamedPopup('prntmsg.aspx?identifier=" & identifier & "&OfferID=" & OfferID & "', 'PrntMsg');"">Details</a>)" & vbCrLf)
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
                  retString.Append("    Your <span alt=""Program " & ProgramID & """ title=""Program " & ProgramID & """>" & ProgramName & "</span> balance: " & Amount)
                End If
                retString.Append("</div>")
                retString.Append("<br />")
                retString.Append("<hr />")
                retString.Append("")
              ElseIf (transformation = "nothing") Then
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

  Public Function Send_XMLGroupOffers(ByRef identifier As String, ByVal CurrentOffers As String, ByVal transformation As String)
    Dim rst As System.Data.DataTable
    Dim rst2 As System.Data.DataTable
    Dim row2 As System.Data.DataRow
    Dim rst3 As System.Data.DataTable
    Dim rstWeb As System.Data.DataTable
    Dim rstCG As System.Data.DataTable
    Dim CPECurrentOffers As String = ""
    Dim CustomerGroups As New StringBuilder()
    Dim CustomerPK As Long
    Dim CustomerGroupID As Long
    Dim RewardGroupID As Long
    Dim ROID As Long
    Dim ExtID As String
    Dim Employee As Boolean
    Dim OfferID As Integer
    Dim OfferName As String
    Dim OfferDesc As String
    Dim OfferStart As Date
    Dim OfferEnd As Date
    Dim OfferOdds As Integer
    Dim OfferDaysLeft As Integer
    Dim GraphicsFileName As String = ""
    Dim AllowOptOut As Boolean = False
    Dim OptOutOffer As Boolean = False
    Dim ExcludedFromOffer As Boolean = False
    Dim PointsConditionOK As Boolean = True
    Dim QtyRequired As Integer = 0
    Dim rowCount As Integer = 0
    Dim i As Integer
    Dim retstring As New StringBuilder
    Dim CustomerTypeID As Integer = 0

    'First off, we check to see if there's an identifier in the URL
    If (identifier <> "") Then

      'There is, so first off we grab the customer's information
      'The identifier the customer supplied is all numbers, so assume it's a card number
      MyCommon.QueryStr = "SELECT C.CustomerPK,C.PrimaryExtID,C.Employee,C.CardStatusID,CE.Email, CustomerTypeID " & _
                          "FROM Customers AS C " & _
                          "LEFT JOIN CustomerExt AS CE on CE.CustomerPK=C.CustomerPK " & _
                          "WHERE C.CustomerPK='" & identifier & "'"
      rst = MyCommon.LXS_Select
      If (rst.Rows.Count > 0) Then
        'A customer was found, so let's assign values to variables
        CustomerPK = rst.Rows(0).Item("CustomerPK")
        CustomerTypeID = MyCommon.NZ(rst.Rows(0).Item("CustomerTypeID"), 0)
        ExtID = rst.Rows(0).Item("PrimaryExtID")
        Employee = rst.Rows(0).Item("Employee")

        If (CurrentOffers.Length > 0) Then
          CurrentOffers = CurrentOffers.Substring(0, CurrentOffers.Length - 1)
          CPECurrentOffers = " and I.IncentiveID not in (" & CurrentOffers & ") "
        End If

        'Next, get the associated customer groups...
        MyCommon.QueryStr = "SELECT CustomerGroupID FROM GroupMembership WHERE CustomerPK=" & CustomerPK & " and Deleted=0"
        rstCG = MyCommon.LXS_Select()
        rstCG.Rows.Add(New String() {"1"})
        rstCG.Rows.Add(New String() {"2"})
        rowCount = rstCG.Rows.Count
        For i = 0 To rowCount - 1
          CustomerGroups.Append(MyCommon.NZ(rstCG.Rows(i).Item("CustomerGroupID"), -1))
          If (i < rowCount - 1) Then CustomerGroups.Append(",")
        Next
        MyCommon.QueryStr = "select I.IncentiveID,I.IncentiveName,I.Description,I.StartDate,I.EndDate,ICG.CustomerGroupID, " & _
                            "ICG.ExcludedUsers,D.OutputID as RewardGroup,I.AllowOptOut,RO.RewardOptionID " & _
                            "from CPE_Incentives I " & _
                            "inner join OfferIDs OID on I.IncentiveID=OID.OfferID " & _
                            "inner join CPE_RewardOptions RO on RO.IncentiveID=I.IncentiveID and RO.TouchResponse=0 and RO.Deleted=0 " & _
                            "inner join CPE_Deliverables D on D.RewardOptionID=RO.RewardOptionID and D.Deleted=0 and DeliverableTypeID=5 and RewardOptionPhase=3 " & _
                            "inner join CPE_IncentiveCustomerGroups ICG on ICG.RewardOptionID=RO.RewardOptionID " & _
                            " and ICG.Deleted=0 and ICG.CustomerGroupID in (" & CustomerGroups.ToString & ") and ICG.ExcludedUsers=0 " & _
                            "where I.Deleted=0 And I.StatusFlag=0 and I.EndDate>='" & Today.AddDays(-1).ToString & "' and OID.EngineID=3 " & CPECurrentOffers & ";"

        'MyCommon.QueryStr = "SELECT distinct O.OfferID,O.ExtOfferID,O.IsTemplate,O.CMOADeployStatus,O.StatusFlag,O.OddsOfWinning, O.InstantWin, " & _
        '                    "O.Name,O.Description,O.ProdStartDate,O.ProdEndDate, LinkID from Offers as O " & _
        '                    "LEFT JOIN OfferConditions as OC on OC.OfferID=O.OfferID " & _
        '                    "WHERE O.IsTemplate=0 and O.Deleted=0 and O.CMOADeployStatus=1 and OC.Deleted=0 and OC.ConditionTypeID=1 " & _
        '                    "and O.DisabledOnCFW = 0 and ProdEndDate > '" & Today.AddDays(-1).ToString & "' and LinkID in (" & CustomerGroups.ToString & ")" & _
        '                    " union all " & _
        '                    "select I.IncentiveID, I.ClientOfferID, I.IsTemplate, I.CPEOADeployStatus, I.StatusFlag, 0 as OddsOfWinning, 0 as InstantWin, " & _
        '                    "I.IncentiveName, I.Description, I.StartDate, I.EndDate, ICG.CustomerGroupID " & _
        '                    "from CPE_Incentives I LEFT JOIN CPE_RewardOptions RO on I.IncentiveID = RO.IncentiveID and RO.TouchResponse = 0 and RO.Deleted =0 " & _
        '                    "LEFT JOIN CPE_IncentiveCustomerGroups  ICG on RO.RewardOptionID = ICG.RewardOptionID " & _
        '                    "WHERE(I.IsTemplate = 0 And I.Deleted = 0 And ICG.Deleted = 0) " & _
        '                    "and I.DisabledOnCFW = 0 and I.EndDate > '" & Today.AddDays(-1).ToString & "' and CustomerGroupID in (" & CustomerGroups.ToString & ")"

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
            ROID = MyCommon.NZ(row2.Item("RewardOptionId"), -1)

            ' Filter out Offers where the customer is already a member of the reward group
            MyCommon.QueryStr = "select MembershipID from GroupMembership where CustomerPK=" & CustomerPK & _
                                    " and deleted=0 and CustomerGroupID=" & RewardGroupID & ";"
            rstWeb = MyCommon.LXS_Select
            OptOutOffer = (rstWeb.Rows.Count > 0)

            ' Check if the customer is in a group that is excluded from this offer
            MyCommon.QueryStr = "select CustomerGroupID from CPE_IncentiveCustomerGroups ICG " & _
                                "where RewardOptionID =" & ROID & " and CustomerGroupID in (" & CustomerGroups.ToString & ") and ExcludedUsers = 1 and Deleted = 0;"
            rstWeb = MyCommon.LRT_Select
            ExcludedFromOffer = (rstWeb.Rows.Count > 0)

            ' Check if the customer meets any points condition, if applicable, that exist for this offer
            MyCommon.QueryStr = "select QtyForIncentive, ProgramID from CPE_IncentivePointsGroups where RewardOptionID=" & ROID & " and deleted = 0;"
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
              MyCommon.QueryStr = "SELECT OL.OfferID,OL.LocationGroupID,OL.Excluded,LG.Name FROM OfferLocations AS OL " & _
                                  "LEFT JOIN LocationGroups AS LG on LG.LocationGroupID=OL.LocationGroupID " & _
                                  "WHERE OL.OfferID=" & OfferID & " and OL.Excluded=0"
              rst3 = MyCommon.LRT_Select

              'Find the name of the associated (and non-excluding) location group
              MyCommon.QueryStr = "SELECT OL.OfferID,OL.LocationGroupID,OL.Excluded,LG.Name FROM OfferLocations AS OL " & _
                                  "LEFT JOIN LocationGroups AS LG on LG.LocationGroupID=OL.LocationGroupID " & _
                                  "WHERE OL.OfferID=" & OfferID & " and OL.Excluded=0"
              rst3 = MyCommon.LRT_Select

              'So finally we have all the info we need. Now display it.
              If (transformation = "XML") Then
                retstring.Append("<Offer>" & vbCrLf)
                retstring.Append("  <OfferID>" & OfferID & "</OfferID>" & vbCrLf)
                retstring.Append("  <Name>" & OfferName & "</Name>" & vbCrLf)
                retstring.Append("  <Description>" & OfferDesc & "</Description>" & vbCrLf)
                retstring.Append("  <StartDate>" & OfferStart & "</StartDate>" & vbCrLf)
                retstring.Append("  <EndDate>" & OfferEnd & "</EndDate>" & vbCrLf)
                retstring.Append("</Offer>" & vbCrLf)
              ElseIf (transformation = "HTML") Then
                retstring.Append("<div class=""group"" id=""group" & OfferID & """>" & vbCrLf)
                retstring.Append("  <h3 class=""name"" alt=""Offer# " & OfferID & ", CGroup# " & CustomerGroupID & """ title=""Offer " & OfferID & ", CGroup " & CustomerGroupID & """>" & OfferName & "</h3>" & vbCrLf)
                GraphicsFileName = RetrieveGraphicPath(OfferID)
                If (GraphicsFileName <> "") Then
                  retstring.Append("    <img src=""images\" & GraphicsFileName & """ align=""right"" />")
                End If
                If (OptOutOffer AndAlso AllowOptOut) Then
                  retstring.Append("  <form action=""ncr-home.aspx"" method=""post"" name=""OptOut"">" & vbCrLf)
                  retstring.Append("    <input type=""hidden"" id=""opt-type"" name=""opt-type"" value=""out"" />")
                  retstring.Append("    <input type=""hidden"" id=""opt-extid"" name=""opt-extid"" value=""" & ExtID & """ />" & vbCrLf)
                  retstring.Append("    <input type=""hidden"" id=""opt-cardtypeid"" name=""opt-cardtypeid"" value=""" & CustomerTypeID & """ />" & vbCrLf)
                  retstring.Append("    <input type=""hidden"" id=""opt-cgroupid"" name=""opt-cgroupid"" value=""" & RewardGroupID & """ />" & vbCrLf)
                  retstring.Append("    <input type=""hidden"" id=""opt-offerid"" name=""opt-offerid"" value=""" & OfferID & """ />" & vbCrLf)
                  retstring.Append("    <input class=""optout"" type=""submit"" value=""X OPT OUT"" />" & vbCrLf)
                  retstring.Append("  </form>" & vbCrLf)
                Else
                  retstring.Append("  <form action=""ncr-home.aspx"" method=""post"" name=""OptIn"">" & vbCrLf)
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
                'PrintedMessage = RetrievePrintedMessage(OfferID)
                'If (PrintedMessage <> "") Then
                retstring.Append(" (<a href=""javascript:openNamedPopup('prntmsg.aspx?identifier=" & identifier & "&OfferID=" & OfferID & "&cg=" & RewardGroupID & "', 'PrntMsg');"">Details</a>)" & vbCrLf)
                'End If
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
              ElseIf (transformation = "nothing") Then
              Else
              End If
            End If
          Next
        End If
      End If
    End If

    Return retstring.ToString
  End Function

  Function sendCustInfo(ByVal identifier As String, ByVal transformation As String) As String
    Dim retString As New StringWriter
    Dim rst As DataTable

    MyCommon.Open_LogixXS()
    MyCommon.QueryStr = "select C.CustomerPK,C.PrimaryExtID,C.FirstName,C.LastName,C.Employee,C.CardStatusID,C.CurrYearSTD,C.LastYearSTD," & _
                        " C.CustomerTypeID,C.HHPK,E.Address,E.City,E.State,E.Zip,E.Country,E.PhoneAsEntered as Phone,E.Email,E.DOB from customers as C " & _
                        " left join CustomerEXT as E on C.customerpk=E.customerpk where C.customerpk=" & identifier
    rst = MyCommon.LXS_Select()
    If rst.Rows.Count > 0 Then
      If (transformation = "XML") Then
        retString.WriteLine("<Customer Editable=""true"" >")
        retString.WriteLine("  <PrimaryExtID>" & MyCommon.NZ(rst.Rows(0).Item("PrimaryExtID"), "") & "</PrimaryExtID>")
        retString.WriteLine("  <FirstName>" & MyCommon.NZ(rst.Rows(0).Item("FirstName"), "") & "</FirstName>")
        retString.WriteLine("  <LastName>" & MyCommon.NZ(rst.Rows(0).Item("LastName"), "") & "</LastName>")
        retString.WriteLine("  <Employee>" & MyCommon.NZ(rst.Rows(0).Item("Employee"), "") & "</Employee>")
        retString.WriteLine("  <CardStatusID>" & MyCommon.NZ(rst.Rows(0).Item("CardStatusID"), "") & "</CardStatusID>")
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
      ElseIf (transformation = "HTML") Then
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
        retString.WriteLine("      <a href=""mailto:" & MyCommon.NZ(rst.Rows(0).Item("Email"), "") & """>" & MyCommon.NZ(rst.Rows(0).Item("Email"), "") & "</a>")
        retString.WriteLine("    </div>")
        retString.WriteLine("    <div id=""phone"">")
        Dim PhoneNumber As String = MyCommon.NZ(rst.Rows(0).Item("Phone"), "")
        retString.WriteLine("      " & PhoneNumber)
        retString.WriteLine("    </div>")
        retString.WriteLine("    <div id=""dob"">")
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
        retString.WriteLine("      " & DOBParts(0) & "/" & DOBParts(1) & "/" & DOBParts(2))
        retString.WriteLine("    </div>")
        retString.WriteLine("  </div>")
        retString.WriteLine("  <div id=""identifiers"">")
        retString.WriteLine("    <div id=""primaryextid"">")
        retString.WriteLine("      Your NCR Program Card is")
        retString.WriteLine("      <span class=""cardnumber"">" & MyCommon.NZ(rst.Rows(0).Item("PrimaryExtID"), "") & "</span>")
        retString.WriteLine("    </div>")
        retString.WriteLine("    <div id=""employeestatus"">")
        retString.WriteLine("      You are an employee.")
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
        retString.WriteLine("    <input type=""button"" class=""medium"" onclick=""javascript:openMiniPopup('ncr-edit.aspx?PrimaryExtID=" & MyCommon.NZ(rst.Rows(0).Item("PrimaryExtID"), "") & "&CustPK=" & MyCommon.NZ(rst.Rows(0).Item("CustomerPK"), "") & "');"" value=""Edit your details"" id=""edit"" name=""edit"" />")
        retString.WriteLine("  </form>")
        retString.WriteLine("</div>")
        retString.WriteLine("<div class=""box"" id=""logoutbox"">")
        retString.WriteLine("  Protect your privacy by logging out when finished.")
        retString.WriteLine("  <form action=""index.html"" id=""logoutform"" name=""logoutform"" target=""_top"">")
        retString.WriteLine("    <input id=""logout"" name=""logout"" class=""medium"" type=""submit"" value=""Logout"" />")
        retString.WriteLine("  </form>")
        retString.WriteLine("</div>")
      ElseIf (transformation = "nothing") Then
      Else
      End If
    End If

    Return retString.ToString
  End Function

  Private Function RetrieveGraphicPath(ByVal OfferID As Long) As String
    Dim GraphicsFileName As String = ""
    Dim GraphicsFilePath As String = ""
    Dim GraphicsNewFilePath As String = ""
    Dim rst As System.Data.DataTable

    Try
      ' Find if a graphic is assigned to this offer
      MyCommon.QueryStr = "select OnScreenAdID, Name, ImageType, Width, Height from OnScreenAds where Deleted = 0 and OnScreenAdID in " & _
                          "(select OutputID from CPE_Deliverables where Deleted = 0 and DeliverableTypeID = 1 and RewardOptionPhase = 1 " & _
                          "and RewardOptionID in (select RewardOptionID from CPE_RewardOptions where IncentiveID = " & OfferID & " and TouchResponse = 0 and Deleted = 0))"
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
          GraphicsNewFilePath = Server.MapPath("ncr-home.aspx")
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
                        "from PrintedMessages PM inner join PrintedMessageTiers PMTiers on PM.MessageID = PMTiers.MessageID " & _
                        "inner join PrintedMessageTypes PMTypes on PM.MessageTypeID = PMTypes.TypeID " & _
                        "inner join CPE_Deliverables D on D.OutputID = PM.MessageID " & _
                        "inner join CPE_RewardOptions RO on RO.RewardOptionID = D.RewardOptionID " & _
                        "where RO.IncentiveId = " & OfferID & " and RO.Deleted = 0 and D.Deleted = 0 and D.RewardOptionPhase = 1 and DeliverableTypeID = 4;"
    rst = MyCommon.LRT_Select

    If (rst.Rows.Count > 0) Then
      PMsgBuf.Append(MyCommon.NZ(rst.Rows(0).Item("BodyText"), ""))
    End If

    MyCommon.Close_LRTsp()

    Return PMsgBuf.ToString()
  End Function

  Private Function GenerateMessagePreview(ByVal newText As String) As String
    Dim PrintLine As String
    Dim RawTextLine As String
    Dim ENDMESSAGE As String = ""
    Dim RawMessage As String
    Dim MESSAGE As String
    Dim x As Integer
    Dim rst As System.Data.DataTable
    Dim row As System.Data.DataRow
    Dim PrinterTag As String
    Dim PreviewText As String

    ' this 2d array will hold the replace and replacement for real time engine tags just add or modify
    ' here
    Dim SubTags(,) As String = { _
                                {"|CUSTOMERID|", "|TSD|", "|LYTS|", "|CURRDATE|", "|OFFERSTART|", "|OFFEREND|", "|TOTALPOIINTS|", "|ACCUMANT|", "|REMAINAMT|"}, _
                                {"###################", "000.00", "000.00", "xx/xx/xxxx", "xx/xx/xxxx", "xx/xx/xxxx", "xx", "000.00", "000.00"}}

    RawMessage = newText

    ' here we will format the message, the message is currently in RawMessage
    ' so lets get it formatted correctly same same
    ' as the local server

    ' first look for and replace  CUSTOMERID  TSD  LYTS  CURRDATE  OFFERSTART  OFFEREND  TOTALPOIINTS  ACCUMANT  REMAINAMT
    MESSAGE = RawMessage
    For x = 0 To SubTags.GetUpperBound(1)
      MESSAGE = Replace(MESSAGE, SubTags(0, x), SubTags(1, x))
    Next

    Dim strArray() As String = MESSAGE.Split(vbLf)

    ' ok we split lets blank out MESSAGE so we can refill it
    MESSAGE = ""

    ' get the tags
    MyCommon.QueryStr = "select PrinterTypeID, '|' + tag + '|' as tag,isnull(pt.previewchars,'') as previewchars from markuptags as MT left join printertranslation as pt on MT.markupid=pt.markupid"
    rst = MyCommon.LRT_Select

    For Each PrintLine In strArray
      ' now lets search on each printer tag to replace
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

      ' now get any remaining tags
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
</script>