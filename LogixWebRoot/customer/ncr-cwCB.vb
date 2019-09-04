Imports System.IO
' version:7.3.1.138972.Official Build (SUSDAY10202)

Public Class cwCB
  Inherits System.Web.UI.Page
  Public UserName As String
  Public LanguageID As Integer
  Dim MyCommon As New Copient.CommonInc

  Public Sub Send(ByVal WebText As String)
    Response.Write(WebText & vbCrLf)
  End Sub

  Public Sub Sendb(ByVal WebText As String)
    Response.Write(WebText)
  End Sub

  Public Sub Send_HeadBegin(ByVal Handheld As Boolean, ByVal PageTitle As String)
    Send("<!-- IE6 quirks mode -->")
    Send("<!DOCTYPE html ")
    Send("     PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN""")
    Send("     ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">")
    Send("<html xmlns=""http://www.w3.org/1999/xhtml"" lang=""en"" xml:lang=""en"">")
    Send("<head>")
    Send("<title>" & PageTitle & "</title>")
  End Sub

  Public Sub Send_Metas(ByVal Handheld As Boolean)
    Send("<meta name=""copyright"" content=""&copy; Copyright 2011, NCR Supermarkets"" />")
    Send("<meta name=""description"" content=""NCR Supermarkets customer-facing website"" />")
    Send("<meta name=""content-type"" content=""text/html; charset=utf-8"" />")
    Send("<meta name=""robots"" content=""noindex, nofollow"" />")
    Send("<meta http-equiv=""cache-control"" content=""no-cache"" />")
    Send("<meta http-equiv=""pragma"" content=""no-cache"" />")
  End Sub

  Public Sub Send_Links(ByVal Handheld As Boolean)
    Send("<link rel=""icon"" href=""images/ncr-favicon.ico"" type=""image/x-icon"" />")
    Send("<link rel=""shortcut icon"" href=""images/ncr-favicon.ico"" type=""image/x-icon"" />")
    If (Handheld) Then
      Send("<link rel=""stylesheet"" href=""css/ncr-cw-handheld.css"" type=""text/css"" media=""handheld, screen"" />")
    Else
      Send("<link rel=""stylesheet"" href=""css/ncr-cw-screen.css"" type=""text/css"" media=""screen"" />")
    End If
  End Sub

  Public Sub Send_HeadEnd(ByVal Handheld As Boolean)
    If (Handheld) Then
    Else
      Send("<script src=""javascript/cw.js"" type=""text/javascript""></script>")
      Send("<script src=""javascript/AC_RunActiveContent.js"" type=""text/javascript""></script>")
    End If
    Send("</head>")
  End Sub

  Public Sub Send_BodyBegin(ByVal Handheld As Boolean, ByVal Popup As Boolean)
    If (Handheld) Then
      Send("<body>")
      Send("<table id=""wrap"" cellpadding=""0"" cellspacing=""0"" summary=""Site"">")
    Else
      If (Popup = True) Then
        Send("<body class=""popup"">")
      Else
        Send("<body>")
      End If
      Send("<div id=""wrap"">")
      Send("<a id=""top"" name=""top""></a>")
    End If
    Send("")
  End Sub

  Public Sub Send_InnerWrapBegin(ByVal Handheld As Boolean)
    Send("<div id=""innerwrap"">")
    Send("")
  End Sub

  Public Sub Send_Logo(ByVal Handheld As Boolean)
    If (Handheld) Then
      Send("<tr>")
      Send("  <td id=""logos"" colspan=""2"">")
      Send("    <a href=""ncr-home.aspx""><img src=""images/logo-mini.jpg"" alt="""" title="""" /></a>")
      Send("  </td>")
      Send("</tr>")
    Else
      Send("<div id=""logos"">")
      Send("  <img src=""images/ncr-logo.png"" id=""logo"" alt="""" title="""" />")
      Send("</div>")
    End If
    Send("")
  End Sub

  Public Sub Send_Menu(ByVal Handheld As Boolean)
    If (Handheld) Then
      Send("<tr>")
      Send("  <td id=""menu"" valign=""bottom"" colspan=""2"">")
      Send("    <a href=""http://www.ncr.com/products_and_services/point_of_sale/index.jsp""><span class=""textnav"">Savings</span></a>")
      Send("    <a href=""http://www.ncr.com/solutions/printer_consumables_solutions/2st_two_sided_thermal_printing/index.jsp?lang=EN""><span class=""textnav"">Recipes</span></a>")
      Send("    <a href=""http://www.ncr.com/solutions/pos_solutions/index.jsp""><span class=""textnav"">Shopping</span></a>")
      Send("    <a href=""http://www.ncr.com/industry/retail/index.jsp""><span class=""textnav"">Food</span></a>")
      Send("    <a href=""http://www.ncr.com/about_ncr/corporate_citizenship/workplace_health_safety.jsp?lang=EN""><span class=""textnav"">Health</span></a>")
      Send("    <a href=""http://www.ncr.com/about_ncr/corporate_citizenship/community.jsp?lang=EN""><span class=""textnav"">Community</span></a>")
      Send("    <a href=""http://www.ncr.com/products_and_services/index.jsp?lang=EN""><span class=""textnav"">Other</span></a>")
      Send("    <a href=""http://www.ncr.com/about_ncr/index.jsp?lang=EN""><span class=""textnav"">About</span></a>")
      Send("  </td>")
      Send("</tr>")
    Else
      Send("<div id=""menu"">")
      Send("  <a href=""http://www.ncr.com/products_and_services/point_of_sale/index.jsp"" id=""menu1"">Savings</a>")
      Send("  <div class=""spacer""></div>")
      Send("  <a href=""http://www.ncr.com/solutions/printer_consumables_solutions/2st_two_sided_thermal_printing/index.jsp?lang=EN"" id=""menu2"">Recipes</a>")
      Send("  <div class=""spacer""></div>")
      Send("  <a href=""http://www.ncr.com/solutions/pos_solutions/index.jsp"" id=""menu3"">Shopping</a>")
      Send("  <div class=""spacer""></div>")
      Send("  <a href=""http://www.ncr.com/industry/retail/index.jsp"" id=""menu4"">Food &amp; Entertaining</a>")
      Send("  <div class=""spacer""></div>")
      Send("  <a href=""http://www.ncr.com/about_ncr/corporate_citizenship/workplace_health_safety.jsp?lang=EN"" id=""menu5"">Healthy Living</a>")
      Send("  <div class=""spacer""></div>")
      Send("  <a href=""http://www.ncr.com/about_ncr/corporate_citizenship/community.jsp?lang=EN"" id=""menu6"">Our Community</a>")
      Send("  <div class=""spacer""></div>")
      Send("  <a href=""http://www.ncr.com/products_and_services/index.jsp?lang=EN"" id=""menu7"">Other Services</a>")
      Send("  <div class=""spacer""></div>")
      Send(" <a href=""http://www.ncr.com/about_ncr/index.jsp?lang=EN"" id=""menu8"">Our Company</a>")
      Send("</div>")
      Send("<hr class=""hidden"" />")
    End If
    Send("")
  End Sub

  Public Sub Send_Submenu(ByVal Handheld As Boolean)
    If (Handheld) Then
      Send("<tr>")
      Send(" <td id=""submenu"" valign=""bottom"" colspan=""2"">")
      Send("  <a href=""http://www.ncr.com/solutions/index.jsp?lang=EN""><span class=""textnav"">WeeklyAd</span></a>")
      Send("  <a href=""http://www.ncr.com/about_ncr/company_overview/index.jsp?lang=EN""><span class=""textnav"">Locator</span></a>")
      Send("  <a href=""http://www.ncr.com/about_ncr/careers/index.jsp""><span class=""textnav"">Employment</span></a>")
      Send("  <a href=""http://www.ncr.com/about_ncr/index.jsp?lang=EN""><span class=""textnav"">SiteMap</span></a>")
      Send("  <a href=""http://www.ncr.com/utility/contact_us/index.jsp?lang=EN""><span class=""textnav"">Contact</span></a>")
      Send(" </td>")
      Send("</tr>")
    Else
      Send("<div id=""submenu"">")
      Send("  <a href=""http://www.ncr.com/utility/contact_us/index.jsp?lang=EN"" id=""submenu5"">Contact</a>")
      Send("  <div class=""spacer"">|</div>")
      Send("  <a href=""http://www.ncr.com/about_ncr/index.jsp?lang=EN"" id=""submenu4"">Site Map</a>")
      Send("  <div class=""spacer"">|</div>")
      Send("  <a href=""http://www.ncr.com/about_ncr/careers/index.jsp"" id=""submenu3"">Employment</a>")
      Send("  <div class=""spacer"">|</div>")
      Send("  <a href=""http://www.ncr.com/about_ncr/company_overview/index.jsp?lang=EN"" id=""submenu2"">Store Locator</a>")
      Send("  <div class=""spacer"">|</div>")
      Send("  <a href=""http://www.ncr.com/solutions/index.jsp?lang=EN"" id=""submenu1"">Weekly Ad</a>")
      Send("</div>")
      Send("<hr class=""hidden"" />")
    End If
    Send("")
  End Sub

  Public Sub Send_SidebarBegin(ByVal Handheld As Boolean)
    If (Handheld) Then
      Send("<tr>")
      Send("  <td id=""sidebar"" valign=""top"">")
    Else
      Send("<div id=""sidebar"">")
    End If
  End Sub

  Public Sub Send_Login(ByVal Handheld As Boolean)
    If (Handheld) Then
      Send("<h2>LOGIN</h2>")
      Send("  <p>Using your NCR Program Card number or your email address, log in to see your NCR offers!</p>")
      Send("  <form action=""ncr-home.aspx"" method=""post"" id=""loginform"" name=""loginform"" target=""_top"">")
      Send("    <label for=""identifier"">Email or card#:</label><br />")
      Send("    <input type=""text"" id=""identifier"" name=""identifier"" /><br />")
      Send("    <label for=""password"">Password:</label><br />")
      Send("    <input type=""password"" id=""password"" name=""password"" /><br />")
      Send("    <br class=""half"" />")
      Send("    <input type=""submit"" id=""login"" name=""submit"" value=""Login"" />")
      Send("  </form>")
      Send("  <br />")
      Send("  <p>First time here? Call customer service at 888-555-5555 to get a password.</p>")
    Else
      Send("<h2>LOGIN</h2>")
      Send("<div class=""box"">")
      Send("  <p>Using your NCR Program Card number or your email address, log in to see your NCR offers!</p>")
      Send("  <form action=""ncr-home.aspx"" method=""post"" id=""loginform"" name=""loginform"" target=""_top"">")
      Send("    <label for=""identifier"">Email or card#:</label><br />")
      Send("    <input type=""text"" id=""identifier"" name=""identifier"" /><br />")
      Send("    <label for=""password"">Password:</label><br />")
      Send("    <input type=""password"" id=""password"" name=""password"" /><br />")
      Send("    <br class=""half"" />")
      Send("    <input type=""submit"" id=""login"" name=""submit"" value=""Login"" />")
      Send("  </form>")
      Send("  <br />")
      Send("  <p>First time here? Call customer service at 888-555-5555 to get a password.</p>")
      Send("</div>")
    End If
  End Sub

  Public Sub Send_Details(ByVal Handheld As Boolean)
    Dim rst As System.Data.DataTable
    Dim MyCryptlib As New Copient.CryptLib
    MyCommon.Open_LogixRT()
    MyCommon.Open_LogixXS()
    Dim AlertMessage As String = ""
    Dim CustomerPK As Long
    Dim ExtID As String = ""
    Dim FirstName As String = ""
    Dim LastName As String = ""
    Dim Employee As Boolean
    Dim CardStatus As String = ""
    Dim Address As String = ""
    Dim City As String = ""
    Dim State As String = ""
    Dim Zip As String = ""
    Dim Phone As String = ""
    Dim Email As String = ""
    Dim Identifier As String = ""

    Identifier = If(Request.Form("identifier") <> "", Request.Form("identifier"), "")
    If (Identifier = "") Then
      Identifier = If(Request.QueryString("identifier") <> "", Request.QueryString("identifier"), "")
    End If
    If (Identifier <> "") Then
      If (IsNumeric(Identifier)) Then
        'The identifier's all numbers, so assume it's a cardnumber
        MyCommon.QueryStr = "SELECT C.CustomerPK,C.PrimaryExtID,C.FirstName,C.LastName,C.Employee,C.CardStatusID,C.CurrYearSTD,C.LastYearSTD," & _
                            "CS.Description,CE.Address,CE.City,CE.State,CE.Zip,CE.PhoneAsEntered as Phone,CE.Email " & _
                            "FROM Customers AS C " & _
                            "INNER JOIN CardStatus AS CS on C.CardStatusID=CS.CardStatusID " & _
                            "LEFT JOIN CustomerExt AS CE on CE.CustomerPK=C.CustomerPK " & _
                            "WHERE PrimaryExtID='" & Identifier & "'"
      Else
        'The identifier's not all numbers, so assume an email address
        MyCommon.QueryStr = "SELECT C.CustomerPK,C.PrimaryExtID,C.FirstName,C.LastName,C.Employee,C.CardStatusID,C.CurrYearSTD,C.LastYearSTD," & _
                            "CS.Description,CE.Address,CE.City,CE.State,CE.Zip,CE.PhoneAsEntered as Phone,CE.Email " & _
                            "FROM Customers AS C " & _
                            "INNER JOIN CardStatus AS CS on C.CardStatusID=CS.CardStatusID " & _
                            "LEFT JOIN CustomerExt AS CE on CE.CustomerPK=C.CustomerPK " & _
                            "WHERE CE.Email='" & Identifier & "'"
      End If
      rst = MyCommon.LXS_Select
      If (rst.Rows.Count > 0) Then
        CustomerPK = rst.Rows(0).Item("CustomerPK")
        ExtID = rst.Rows(0).Item("PrimaryExtID")
        FirstName = MyCommon.NZ(rst.Rows(0).Item("FirstName"), "")
        LastName = MyCommon.NZ(rst.Rows(0).Item("LastName"), "")
        Employee = MyCommon.NZ(rst.Rows(0).Item("Employee"), "")
        CardStatus = MyCommon.NZ(rst.Rows(0).Item("Description"), "")
        CardStatus = CardStatus.Replace("_", "/")
        CardStatus = (StrConv(CardStatus, VbStrConv.Lowercase))
        Address = MyCommon.NZ(rst.Rows(0).Item("Address"), "")
        City = MyCommon.NZ(rst.Rows(0).Item("City"), "")
        State = MyCommon.NZ(rst.Rows(0).Item("State"), "")
        Zip = MyCommon.NZ(rst.Rows(0).Item("Zip"), "")
        Phone = MyCommon.NZ(rst.Rows(0).Item("Phone"), "")
        Email = MyCommon.NZ(rst.Rows(0).Item("Email"), "")
      End If
    Else
    End If

    If (Request.Form("save") <> "") Then
      If (Request.Form("newpass1") <> "") Then
        If (Request.Form("newpass1") = Request.Form("newpass2")) Then
          MyCommon.QueryStr = "update Customers set FirstName=N'" & Request.Form("firstname") & "'," & _
          "LastName=N'" & Request.Form("lastname") & "'," & _
          "Password=N'" & MyCryptlib.SQL_StringEncrypt(Request.Form("newpass1")) & " where CustomerPK=" & CustomerPK
          MyCommon.LXS_Execute()
        Else
          AlertMessage = "The new passwords you entered don't match."
        End If
      Else
        MyCommon.QueryStr = "update Customers set FirstName=N'" & Request.Form("firstname") & "'," & _
        "LastName=N'" & Request.Form("lastname") & "' where CustomerPK=" & CustomerPK
        MyCommon.LXS_Execute()
      End If

      MyCommon.QueryStr = "update CustomerExt set Address=N'" & Request.Form("address1") & "'," & _
      "City=N'" & Request.Form("city") & "'," & _
      "State=N'" & Request.Form("state") & "'," & _
      "Zip=N'" & Request.Form("zip") & "'," & _
      "PhoneAsEntered=N'" & Request.Form("phone") & "'," & _
      "PhoneDigitsOnly=N'" & MyCommon.DigitsOnly(Request.Form("phone")) & "'," & _
      "Email=N'" & Request.Form("emailaddress") & "' where CustomerPK=" & CustomerPK
      MyCommon.LXS_Execute()

      FirstName = Request.Form("firstname")
      LastName = Request.Form("lastname")
      Address = Request.Form("address1")
      City = Request.Form("city")
      State = Request.Form("state")
      Zip = Request.Form("zip")
      Phone = Request.Form("phone")
      Email = Request.Form("emailaddress")
    End If

    If (Handheld) Then
    Else
      Send("<h2>Your details</h2>")
    End If
    Send("<div class=""box"" id=""detailsbox"">")
    Send("  <div id=""name"">")
    If (FirstName = "") And (LastName = "") Then
      Send("    Anonymous")
    Else
      Send("    " & FirstName & " " & LastName)
    End If
    Send("  </div>")
    Send("  <div id=""contact"">")
    Send("    <div id=""address"">")
    If (Address = "") Then
      Send("      No address provided<br />")
    Else
      Send("      " & Address)
    End If
    Send("    </div>")
    Send("    <div id=""city"">")
    If (City = "") Then
      Send("      No City")
    Else
      Send("      " & City)
    End If
    Send("    </div>")
    Send("    <div id=""state"">")
    If (State = "") Then
      Send("      No State")
    Else
      Send("      " & State)
    End If
    Send("    </div>")
    Send("    <div id=""zip"">")
    If (Zip = "") Then
      Send("      No ZIP")
    Else
      Send("      " & Zip)
    End If
    Send("    </div>")
    Send("    <div id=""country"">")
    Send("    </div>")
    Send("    <div id=""phone"">")
    If (Phone = "") Then
      Send("      No phone provided")
    Else
      Send("      " & Phone)
    End If
    Send("    </div>")
    Send("    <div id=""email"">")
    If (Email = "") Then
      Send("      No email provided<br />")
    Else
      Send("      <a href=""mailto:" & Email & """>" & Email & "</a>")
    End If
    Send("    </div>")
    Send("  </div>")
    Send("  <div id=""identifiers"">")
    Send("    <div id=""primaryextid"">")
    If (ExtID = "") Then
      Send("      No card number")
    Else
      Send("      Your NCR Program Card is")
      Send("      <span class=""cardnumber"">" & ExtID & "</span>")
    End If
    Send("    </div>")
    Send("    <div id=""cardstatus"">")
    Send("      Your card is " & CardStatus & ".")
    Send("    </div>")
    Send("    <div id=""employeestatus"">")
    If (Employee = True) Then
      Send("      You are an employee.<br />")
    End If
    Send("    </div>")
    Send("    <form action=""#"" id=""editform"" name=""editform"">")
    Send("      <button class=""medium"" id=""edit"" name=""edit"" type=""button"" onclick=""javascript:handleEditButton();"">Edit your details</button><br />")
    Send("    </form>")
    Send("  </div>")
    Send("</div>")
    Send("<div id=""editdetails"" style=""display:none;"">")
    Send("  <form id=""editform"" method=""post"" action=""nhome.aspx"">")
    Send("    <input type=""hidden"" name=""identifier"" id=""identifier"" value=""" & Identifier & """ />")
    Send("    <label for=""firstname"">First name</label>, <label for=""lastname"">Last name</label><br />")
    Send("    <input id=""firstname"" name=""firstname"" value=""" & FirstName & """ /><input id=""lastname"" name=""lastname"" value=""" & LastName & """ /><br />")
    Send("    <label for=""address1"">Address</label><br />")
    Send("    <input id=""address1"" name=""address1"" value=""" & Address & """ /><br />")
    Send("    <label for=""city"">City</label>, <label for=""state"">State</label><br />")
    Send("    <input id=""city"" name=""city"" value=""" & City & """ /><input id=""state"" name=""state"" value=""" & State & """ /><br />")
    Send("    <label for=""zip"">ZIP Code</label><br />")
    Send("    <input id=""zip"" name=""zip"" value=""" & Zip & """ /><br />")
    Send("    <label for=""phone"">Phone</label><br />")
    Send("    <input id=""phone"" name=""phone"" value=""" & Phone & """ /><br />")
    Send("    <label for=""emailaddress"">Email</label><br />")
    Send("    <input id=""emailaddress"" name=""emailaddress"" value=""" & Email & """ /><br />")
    Send("    <label for=""newpass1"">New password<br />(enter twice)</label><br />")
    Send("    <input id=""newpass1"" name=""newpass1"" type=""password"" value="""" /><br />")
    Send("    <input id=""newpass2"" name=""newpass2"" type=""password"" value="""" /><br />")
    Send("    <input id=""save"" name=""save"" type=""submit"" onclick=""javascript:{document.getElementById('showdetails').style.display='block';document.getElementById('editdetails').style.display='none'}"" value=""Save"" /><br />")
    Send("  </form>")
    Send("</div>")
    Send("")
  End Sub

  Public Sub Send_Balances(ByVal Handheld As Boolean)
    Dim rst As System.Data.DataTable
    Dim row As System.Data.DataRow
    Dim rst2 As System.Data.DataTable
    Dim row2 As System.Data.DataRow
    Dim rst3 As System.Data.DataTable
    Dim rst5 As System.Data.DataTable
    Dim rst6 As System.Data.DataTable
    Dim row6 As System.Data.DataRow
    Dim rowCount As Integer
    Dim CustomerPK As Long
    Dim ExtID As String = ""
    Dim OfferGroupType As Integer
    Dim ProgramID As String = ""
    Dim ProgramName As String = ""
    Dim ProgramIDArray() As String
    Dim ProgramNameArray() As String
    Dim OfferID As Integer
    Dim Identifier As String = ""

    MyCommon.Open_LogixRT()
    MyCommon.Open_LogixXS()

    Send("")
    Send("<div class=""box"" id=""balances"">")

    Identifier = If(Request.Form("identifier") <> "", Request.Form("identifier"), "")
    If (Identifier = "") Then
      Identifier = If(Request.QueryString("identifier") <> "", Request.QueryString("identifier"), "")
    End If

    'First off, we check to see if there's an identifier in the URL
    If (Identifier <> "") Then
      ' There is, so first off we grab the customer's information
      If (IsNumeric(Identifier)) Then
        'The identifier the customer supplied is all numbers, so assume it's a card number
        MyCommon.QueryStr = "SELECT C.CustomerPK,C.PrimaryExtID,C.CardStatusID,C.CurrYearSTD,C.LastYearSTD,CE.Email " & _
                            "FROM Customers AS C " & _
                            "LEFT JOIN CustomerExt AS CE on CE.CustomerPK=C.CustomerPK " & _
                            "WHERE PrimaryExtID='" & Identifier & "'"
      Else
        'The identifier the customer supplied isn't just numbers, so assume an email address
        MyCommon.QueryStr = "SELECT C.CustomerPK,C.PrimaryExtID,C.CardStatusID,C.CurrYearSTD,C.LastYearSTD,CE.Email " & _
                            "FROM Customers AS C " & _
                            "LEFT JOIN CustomerExt AS CE on CE.CustomerPK=C.CustomerPK " & _
                            "WHERE Email='" & Identifier & "'"
      End If
      rst = MyCommon.LXS_Select
      If (rst.Rows.Count > 0) Then
        'A customer was found, so let's assign values to variables
        CustomerPK = rst.Rows(0).Item("CustomerPK")
        ExtID = rst.Rows(0).Item("PrimaryExtID")

        Send("  <h3>Savings</h3>")
        Send("  <p>This year you've saved $" & MyCommon.NZ(rst.Rows(0).Item("CurrYearSTD"), "0") & ".</p>")

        'Next, get the associated customer groups...
        MyCommon.QueryStr = "SELECT CustomerGroupID FROM GroupMembership WHERE CustomerPK=" & CustomerPK & " and Deleted=0"
        rst = MyCommon.LXS_Select()
        rst.Rows.Add(New String() {"1"})
        rst.Rows.Add(New String() {"2"})
        rowCount = rst.Rows.Count

        If (rowCount > 0) Then
          'The customer's in at least one group, so for each one we'll grab the associated offer(s)
          For Each row In rst.Rows
            MyCommon.QueryStr = "SELECT O.OfferID,O.ExtOfferID,O.IsTemplate,O.CMOADeployStatus,O.StatusFlag,O.OddsOfWinning, O.InstantWin, " & _
                                "O.Name,O.Description,O.ProdStartDate,O.ProdEndDate from Offers as O " & _
                                "LEFT JOIN OfferConditions as OC on OC.OfferID=O.OfferID " & _
                                "WHERE O.IsTemplate=0 and O.Deleted=0 and O.CMOADeployStatus=1 and OC.Deleted=0 and O.DisabledOnCFW = 0 and OC.ConditionTypeID=1 and LinkID=" & row.Item("CustomerGroupID") & _
                                " union all " & _
                                "select I.IncentiveID, I.ClientOfferID, I.IsTemplate, I.CPEOADeployStatus, I.StatusFlag, 0 as OddsOfWinning, 0 as InstantWin, " & _
                                "I.IncentiveName, I.Description, I.StartDate, I.EndDate " & _
                                "from CPE_Incentives I LEFT JOIN CPE_RewardOptions RO on I.IncentiveID = RO.IncentiveID and RO.TouchResponse = 0 and RO.Deleted =0 " & _
                                "LEFT JOIN CPE_IncentiveCustomerGroups  ICG on RO.RewardOptionID = ICG.RewardOptionID " & _
                                "WHERE(I.IsTemplate = 0 And I.Deleted = 0 And ICG.Deleted = 0) " & _
                                "and I.DisabledOnCFW = 0 and CustomerGroupID=" & row.Item("CustomerGroupID")

            rst2 = MyCommon.LRT_Select

            'Set the general info for each offer found
            For Each row2 In rst2.Rows
              OfferID = row2.Item("OfferID")

              'Find the name of the associated (and non-excluding) location group
              MyCommon.QueryStr = "SELECT OL.OfferID,OL.LocationGroupID,OL.Excluded,LG.Name FROM OfferLocations AS OL " & _
                                  "LEFT JOIN LocationGroups AS LG on LG.LocationGroupID=OL.LocationGroupID " & _
                                  "WHERE OL.OfferID=" & OfferID & " and OL.Excluded=0"
              rst3 = MyCommon.LRT_Select

              'Find the name of the customer group
              'MyCommon.QueryStr = "SELECT Name,AllowOptIn,AllowOptOut FROM CustomerGroups WHERE CustomerGroupID=" & row.Item("CustomerGroupID") & " AND AllowOptIn=0"
              'rst4 = MyCommon.LRT_Select

              'Get LinkID
              MyCommon.QueryStr = "SELECT LinkID FROM OfferConditions WHERE ConditionTypeID=1 AND Deleted=0 AND OfferID=" & row2.Item("OfferID")
              rst5 = MyCommon.LRT_Select
              If (rst5.Rows.Count > 0) Then
                OfferGroupType = rst5.Rows(0).Item("LinkID")
              End If

              'For the current offer, fill the array with values of any points programs we run across
              MyCommon.QueryStr = "SELECT 1 as EngineID, O.OfferID,LinkID,ProgramName,PP.ProgramID,PromoVarID " & _
                                   "FROM OfferRewards OFR with (NoLock) " & _
                                   "LEFT JOIN RewardPoints  RP with (NoLock) ON RP.RewardPointsID=OFR.LinkID " & _
                                   "LEFT JOIN PointsPrograms  PP with (NoLock) ON RP.ProgramID=PP.ProgramID " & _
                                   "LEFT JOIN Offers O with (NoLock) ON O.OfferID=OFR.OfferID " & _
                                   "WHERE(rewardtypeid = 2 And O.deleted = 0 And OFR.deleted = 0) " & _
                                   "AND RP.ProgramID IS NOT null " & _
                                   "AND O.Offerid=" & row2.Item("OfferID") & _
                                   " UNION " & _
                                   "SELECT 2 as EngineID, RO.IncentiveID, D.DeliverableID, PP.ProgramName, PP.ProgramID, PromoVarID " & _
                                   "FROM PointsPrograms PP with (NoLock) " & _
                                   "LEFT JOIN CPE_DeliverablePoints DP with (NoLock) on PP.ProgramID=DP.ProgramID " & _
                                   "INNER JOIN CPE_Deliverables D with (NoLock) on D.DeliverableID=DP.DeliverableID " & _
                                   "INNER JOIN CPE_RewardOptions RO with (NoLock) on RO.RewardOptionId = D.RewardOptionID " & _
                                   "WHERE(RO.IncentiveID = " & row2.Item("OfferID") & " And PP.Deleted = 0 And DP.Deleted = 0 And D.Deleted = 0) " & _
                                   "AND RO.Deleted=0 and D.RewardOptionPhase=3 " & _
                                   "ORDER by PP.ProgramName;"

              'MyCommon.QueryStr = "SELECT O.OfferID,LinkID,ProgramName,PP.ProgramID,PromoVarID from OfferRewards as OFR " & _
              '                    "LEFT JOIN RewardPoints AS RP ON RP.RewardPointsID=OFR.LinkID " & _
              '                    "LEFT JOIN PointsPrograms AS PP ON RP.ProgramID=PP.ProgramID " & _
              '                    "LEFT JOIN Offers as O ON O.OfferID=OFR.OfferID " & _
              '                    "WHERE (rewardtypeid = 2 AND O.deleted = 0 AND OFR.deleted = 0) " & _
              '                    "AND RP.ProgramID IS NOT null " & _
              '                    "AND O.Offerid=" & row2.Item("OfferID")
              rst6 = MyCommon.LRT_Select
              For Each row6 In rst6.Rows
                If (ProgramID = "") Then
                  ProgramID = row6.Item("ProgramID")
                  ProgramName = MyCommon.NZ(row6.Item("ProgramName"), "Unknown!").ToString.Replace(",", " ")
                Else
                  'Check for uniqueness on building up the string
                  Dim tmpArray() As String
                  Dim w As Integer
                  Dim Found As Boolean = False
                  tmpArray = ProgramID.Split(",")
                  For w = 0 To tmpArray.GetUpperBound(0)
                    If (tmpArray(w) = row6.Item("ProgramID")) Then
                      Found = True
                    End If
                  Next
                  If (Not Found) Then
                    ProgramID = ProgramID & "," & row6.Item("ProgramID")
                    ProgramName = ProgramName & "," & MyCommon.NZ(row6.Item("ProgramName"), "Unknown!").ToString.Replace(",", " ")
                  End If
                End If
              Next

            Next
          Next
        Else
          'The customer's isn't in any groups
        End If
      Else
        'No customer was found
      End If
    Else
      'No identifier was found in the URL
    End If

    If (Identifier <> "") Then
      'Select from Points in XS on the CustomerPK, then take the returned
      'PromoVarIDs and select the matching points program names from RT
      If (CustomerPK <> 0 And ProgramName <> "" And ProgramID <> "") Then
        ProgramNameArray = ProgramName.Split(",")
        ProgramIDArray = ProgramID.Split(",")
        Dim z As Integer
        Dim Amount As Long
        Send("  <h3>Balances</h3>")
        For z = 0 To ProgramNameArray.GetUpperBound(0)
          'Grab two things: the PromoVarID and the balance (if any)
          Dim promoVarLocal As Long
          promoVarLocal = 0

          MyCommon.QueryStr = "select PromoVarID from PointsPrograms where ProgramID=" & ProgramIDArray(z)
          rst3 = MyCommon.LRT_Select
          If (rst3.Rows.Count > 0) Then
            promoVarLocal = rst3.Rows(0).Item("PromoVarID")
          End If

          MyCommon.QueryStr = "select Amount from Points where CustomerPK=" & CustomerPK & " and PromoVarID=" & promoVarLocal
          rst3 = MyCommon.LXS_Select
          If (rst3.Rows.Count > 0) Then
            Amount = rst3.Rows(0).Item("Amount")
          Else
            Amount = 0
          End If

          ' Demo fix to Limit the Programs displayed to no more than 3 so page does not scroll
          If (z < 3) Then
            If (Amount = 1) Then
              Send("  <b>" & Amount & "</b> point in " & ProgramNameArray(z) & "<br />")
            Else
              Send("  <b>" & Amount & "</b> points in " & ProgramNameArray(z) & "<br />")
            End If
          End If
        Next
      Else
        Send("  You have no points balances associated with any active offers.<br />")
      End If
    End If
    Send("</div>")
  End Sub

  Public Sub Send_Accumulations(ByVal Handheld As Boolean, ByVal OfferID As Long)
    Dim EngineID As Integer = -1
    Dim rst As System.Data.DataTable
    Dim row As System.Data.DataRow
    Dim AccumProgram As Boolean = False
    Dim RewardOptionID As Long = -1
    Dim HHEnable As Boolean = False
    Dim UnitType As Integer = 0
    Dim HHPrimaryID As Integer = 0
    Dim Identifier As String = ""
    Dim CustomerPK As Long
    Dim TotalAccum As Double
    Dim CurrentAccum As Double

    MyCommon.QueryStr = "select EngineID from OfferIDs where OfferID = " & OfferID
    rst = MyCommon.LRT_Select

    If (rst.Rows.Count > 0) Then
      EngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), -1)
    End If

    Identifier = If(Request.Form("identifier") <> "", Request.Form("identifier"), "")
    If (Identifier = "") Then
      Identifier = If(Request.QueryString("identifier") <> "", Request.QueryString("identifier"), "")
    End If

    ' Currently, only CPE supports accumulation balance
    If (EngineID = 2) Then
      MyCommon.QueryStr = "select IPG.AccumMin, RO.RewardOptionID, RO.HHEnable, IPG.QtyUnitType " & _
                 "from CPE_IncentiveProductGroups as IPG Inner Join CPE_RewardOptions as RO on IPG.RewardOptionID=RO.RewardOptionID and IPG.Deleted=0 and IPG.ExcludedProducts=0 and RO.Deleted=0 " & _
                 "where RO.IncentiveID=" & OfferID & ";"
      rst = MyCommon.LRT_Select
      If (rst.Rows.Count > 0) Then
        If MyCommon.NZ(rst.Rows(0).Item("AccumMin"), 0) > 0 Then
          AccumProgram = True
        End If
        RewardOptionID = rst.Rows(0).Item("RewardOptionID")
        HHEnable = MyCommon.NZ(rst.Rows(0).Item("HHEnable"), False)
        UnitType = MyCommon.NZ(rst.Rows(0).Item("QtyUnitType"), 2)
      End If

      If HHEnable Then
        MyCommon.QueryStr = "select HHPK from Customers where PrimaryExtID='" & Identifier & "';"
        rst = MyCommon.LXS_Select
        If (rst.Rows.Count > 0) Then
          HHPrimaryID = MyCommon.NZ(rst.Rows(0).Item("HHPK"), 0)
        End If
      End If

      If AccumProgram Then
        'Query for the CustomerPK from the External ID
        MyCommon.QueryStr = "select CustomerPK from Customers with (NoLock) where PrimaryExtId='" & Identifier & "';"
        rst = MyCommon.LXS_Select()
        If (rst.Rows.Count > 0) Then
          CustomerPK = rst.Rows(0).Item("CustomerPK")
        End If

        If HHEnable Then
          MyCommon.QueryStr = "select RA.AccumulationDate, RA.LastServerID, RA.ServerSerial, RA.CustomerPK, RA.PurchCustomerPK, RA.TotalPrice, RA.QtyPurchased, RA.Deleted " & _
                              "from CPE_RewardAccumulation as RA with (NOLOCK) where (RA.CustomerPK=" & CustomerPK & " or RA.CustomerPK=" & HHPrimaryID & ") and RA.RewardOptionID=" & RewardOptionID & " order by AccumulationDate;"
        Else
          MyCommon.QueryStr = "select RA.AccumulationDate, RA.LastServerID, RA.ServerSerial, RA.CustomerPK, RA.PurchCustomerPK, RA.TotalPrice, RA.QtyPurchased, RA.Deleted, RA.LocationID " & _
                              "from CPE_RewardAccumulation as RA with (NOLOCK) where RA.CustomerPK=" & CustomerPK & " and RA.RewardOptionID=" & RewardOptionID & " order by AccumulationDate;"
        End If

        rst = MyCommon.LXS_Select
        If (rst.Rows.Count > 0) Then
          TotalAccum = 0
          CurrentAccum = 0
          For Each row In rst.Rows
            If HHEnable Then
              'DistCardNum = "0"
              'MyCommon.QueryStr = "select ClientUserID1 from Users where UserID=" & MyCommon.NZ(rst.Rows(0).Item("PurchUserID"), 0) & ";"
              'rst2 = MyCommon.LXS_Select
              'If Not (rst2.Rows.Count > 0) Then
              '    DistCardNum = MyCommon.NZ(rst2.Rows(0).Item("ClientUserID1"), "0")
              'End If
            End If

            If UnitType = 1 Then
              TotalAccum = TotalAccum + Math.Round(MyCommon.NZ(row.Item("QtyPurchased"), 0), 0)
              If row.Item("Deleted") = False Then CurrentAccum = CurrentAccum + Math.Round(MyCommon.NZ(row.Item("QtyPurchased"), 0), 0)
            ElseIf UnitType = 2 Then
              TotalAccum = TotalAccum + Math.Round(MyCommon.NZ(row.Item("TotalPrice"), 0), 2)
              If row.Item("Deleted") = False Then CurrentAccum = CurrentAccum + Math.Round(MyCommon.NZ(row.Item("TotalPrice"), 0), 2)
            ElseIf UnitType = 3 Then
              TotalAccum = TotalAccum + Math.Round(MyCommon.NZ(row.Item("QtyPurchased"), 0), 3)
              If row.Item("Deleted") = False Then CurrentAccum = CurrentAccum + Math.Round(MyCommon.NZ(row.Item("QtyPurchased"), 0), 3)
            End If
          Next
          Sendb("    Total Amt Accumulated:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")
          If UnitType = 1 Then
            Sendb(Format(TotalAccum, "###,##0"))
          ElseIf UnitType = 2 Then
            Sendb("$" & Format(TotalAccum, "###,##0.00"))
          ElseIf UnitType = 3 Then
            Sendb(Format(TotalAccum, "###,##0.000"))
          End If
          Send("    <br />")
          Sendb("    Current Amt Accumulated:&nbsp;&nbsp;")
          If UnitType = 1 Then
            Sendb(Format(CurrentAccum, "###,##0"))
          ElseIf UnitType = 2 Then
            Sendb("$" & Format(CurrentAccum, "###,##0.00"))
          ElseIf UnitType = 3 Then
            Sendb(Format(CurrentAccum, "###,##0.000"))
          End If
          Send("    <br /><br class=""half"">")
        End If
      End If
    End If
  End Sub

  Public Sub Send_Logout(ByVal Handheld As Boolean)
    Send("<div class=""box"" id=""logoutbox"">")
    Send("  Protect your privacy by logging out when finished.<br />")
    Send("  <form action=""index.aspx"" id=""logoutform"" name=""logoutform"" target=""_top"">")
    Send("    <input id=""logout"" name=""logout"" class=""medium"" type=""submit"" value=""Logout"" />")
    Send("  </form>")
    Send("</div>")
  End Sub

  Public Sub Send_SidebarEnd(ByVal Handheld As Boolean)
    If (Handheld) Then
      Send("  </td>")
    Else
      Send("  <hr class=""hidden"" />")
      Send("</div>")
    End If
    Send("")
  End Sub

  Public Sub Send_Gutter(ByVal Handheld As Boolean)
    Send("<div class=""gutter"">")
    Send("</div>")
    Send("")
  End Sub

  Public Sub Send_MainBegin(ByVal Handheld As Boolean)
    If (Handheld) Then
      Send("<td id=""main"" valign=""top"">")
    Else
      Send("<div id=""main"">")
      Send("")
    End If
  End Sub

  Public Sub Send_Ads(ByVal Handheld As Boolean)
    If (Handheld) Then
      Send("  <img src=""images/general-mini.jpg"" alt=""Ads"" title=""Ads"" align=""right"" /><br />")
    Else
      Send("")
      Send("  <div id=""ads"">")
      Send("    <img src=""images/general.png"" alt=""Ads"" title=""Ads"" width=""160"" /><br />")
      Send("  </div>")
      Send("")
    End If
  End Sub

  Public Sub Send_Flash()
    Send("<div id=""flash"">")
    Send("  <script type=""text/javascript"">")
    Send("  AC_FL_RunContent(")
    Send("    'classid','clsid:d27cdb6e-ae6d-11cf-96b8-444553540000',")
    Send("    'codebase','http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0',")
    Send("    'width','360',")
    Send("    'height','240',")
    Send("    'id','flash',")
    Send("    'align','left',")
    Send("    'standby','...',")
    Send("    'src','images/flash',")
    Send("    'pluginspage','http://www.macromedia.com/go/getflashplayer',")
    Send("    'allowScriptAccess','sameDomain',")
    Send("    'bgcolor','#e7f7e9',")
    Send("    'loop','false',")
    Send("    'menu','false',")
    Send("    'movie','images/flash',")
    Send("    'quality','high',")
    Send("    'wmode','transparent' );")
    Send("  </script>")
    Send("  <noscript>")
    Send("    <object classid=""clsid:d27cdb6e-ae6d-11cf-96b8-444553540000"" codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0"" width=""400"" height=""240"" id=""flash"" align=""right"" standby=""..."">")
    Send("      <param name=""allowScriptAccess"" value=""sameDomain"" />")
    Send("      <param name=""bgcolor"" value=""#e7f7e9"" />")
    Send("      <param name=""loop"" value=""false"" />")
    Send("      <param name=""menu"" value=""false"" />")
    Send("      <param name=""movie"" value=""images/flash.swf"" />")
    Send("      <param name=""quality"" value=""high"" />")
    Send("      <param name=""wmode"" value=""transparent"" />")
    Send("    <!--[if !IE]> <-->")
    Send("    <object data=""images/flash.swf"" width=""360"" height=""240"" type=""application/x-shockwave-flash"">")
    Send("      <param name=""bgcolor"" value=""#e7f7e9"" />")
    Send("      <param name=""loop"" value=""false"" />")
    Send("      <param name=""pluginurl"" value=""http://www.macromedia.com/go/getflashplayer"" />")
    Send("      <param name=""quality"" value=""high"" />")
    Send("      <param name=""wmode"" value=""transparent"" />")
    Send("    </object>")
    Send("    <!--> <![endif]-->")
    Send("    </object>")
    Send("  </noscript>")
    Send("</div>")
  End Sub

  Public Sub Send_CurrentOffers(ByRef CurrentOffers As String, ByVal Handheld As Boolean)
    Dim rst As System.Data.DataTable
    Dim row As System.Data.DataRow
    Dim rst2 As System.Data.DataTable
    Dim row2 As System.Data.DataRow
    Dim rst3 As System.Data.DataTable
    Dim rstWeb As System.Data.DataTable
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
    Dim Identifier As String
    Dim PrintedMessage As String = ""

    MyCommon.Open_LogixRT()
    MyCommon.Open_LogixXS()

    Send("")
    Send("  <div id=""offers"">")

    Identifier = If(Request.Form("identifier") <> "", Request.Form("identifier"), "")
    If (Identifier = "") Then
      Identifier = If(Request.QueryString("identifier") <> "", Request.QueryString("identifier"), "")
    End If

    'First off, we check to see if there's an identifier in the URL
    If (Identifier <> "") Then

      ' There is, so first off we grab the customer's information
      If (IsNumeric(Identifier)) Then
        'The identifier the customer supplied is all numbers, so assume it's a card number
        MyCommon.QueryStr = "SELECT C.CustomerPK,C.PrimaryExtID,C.Employee,C.CardStatusID,CE.Email " & _
                            "FROM Customers AS C " & _
                            "LEFT JOIN CustomerExt AS CE on CE.CustomerPK=C.CustomerPK " & _
                            "WHERE PrimaryExtID='" & Identifier & "'"
      Else
        'The identifier the customer supplied isn't just numbers, so assume an email address
        MyCommon.QueryStr = "SELECT C.CustomerPK,C.PrimaryExtID,C.Employee,C.CardStatusID,CE.Email " & _
                            "FROM Customers AS C " & _
                            "LEFT JOIN CustomerExt AS CE on CE.CustomerPK=C.CustomerPK " & _
                            "WHERE Email='" & Identifier & "'"
      End If
      rst = MyCommon.LXS_Select
      If (rst.Rows.Count > 0) Then
        'A customer was found, so let's assign values to variables
        CustomerPK = rst.Rows(0).Item("CustomerPK")
        ExtID = rst.Rows(0).Item("PrimaryExtID")
        Employee = rst.Rows(0).Item("Employee")

        Send("    <h2>Active Offers</h2>")

        'Next, get the associated customer groups...
        MyCommon.QueryStr = "SELECT CustomerGroupID FROM GroupMembership WHERE CustomerPK=" & CustomerPK & " and Deleted=0"
        rst = MyCommon.LXS_Select()
        rst.Rows.Add(New String() {"1"})
        rst.Rows.Add(New String() {"2"})
        rowCount = rst.Rows.Count
        'The customer's in at least one group, so for each one we'll grab the associated offer(s)
        For Each row In rst.Rows
          MyCommon.QueryStr = "SELECT distinct O.OfferID,O.ExtOfferID,O.IsTemplate,O.CMOADeployStatus,O.StatusFlag,O.OddsOfWinning, O.InstantWin, " & _
                              "O.Name,O.Description,O.ProdStartDate,O.ProdEndDate, LinkID from Offers as O " & _
                              "LEFT JOIN OfferConditions as OC on OC.OfferID=O.OfferID " & _
                              "WHERE O.IsTemplate=0 and O.Deleted=0 and O.CMOADeployStatus=1 and OC.Deleted=0 and OC.ConditionTypeID=1 " & _
                              "and O.DisabledOnCFW = 0 and ProdEndDate > '" & Today.AddDays(-1).ToString & "' and LinkID=" & row.Item("CustomerGroupID") & _
                              "union all " & _
                              "select I.IncentiveID, I.ClientOfferID, I.IsTemplate, I.CPEOADeployStatus, I.StatusFlag, 0 as OddsOfWinning, 0 as InstantWin, " & _
                              "I.IncentiveName, I.Description, I.StartDate, I.EndDate, ICG.CustomerGroupID " & _
                              "from CPE_Incentives I LEFT JOIN CPE_RewardOptions RO on I.IncentiveID = RO.IncentiveID and RO.TouchResponse = 0 and RO.Deleted =0 " & _
                              "LEFT JOIN CPE_IncentiveCustomerGroups  ICG on RO.RewardOptionID = ICG.RewardOptionID " & _
                              "WHERE(I.IsTemplate = 0 And I.Deleted = 0 And ICG.Deleted = 0) " & _
                              "and I.DisabledOnCFW = 0 and I.EndDate > '" & Today.AddDays(-1).ToString & "' and CustomerGroupID=" & row.Item("CustomerGroupID")
          '"WHERE(I.IsTemplate = 0 And I.Deleted = 0 And ICG.Deleted = 0) " & _
          'Send(MyCommon.QueryStr)
          rst2 = MyCommon.LRT_Select

          'Set the general info for each offer found
          For Each row2 In rst2.Rows

            ID = row2.Item("OfferID")
            OfferName = row2.Item("Name")
            OfferDesc = row2.Item("Description")
            If (OfferDesc.Trim() = "") Then OfferDesc = OfferName
            OfferStart = row2.Item("ProdStartDate")
            OfferEnd = row2.Item("ProdEndDate")
            OfferOdds = row2.Item("OddsOfWinning")
            InstantWin = MyCommon.NZ(row2.Item("InstantWin"), 0)
            OfferDaysLeft = DateDiff("d", Today, OfferEnd)
            CustomerGroupID = row2.Item("LinkID")

            ' Filter out the Website Offers
            MyCommon.QueryStr = "select OfferID from OfferIDs where OfferID=" & OfferID & " and EngineID=3;"
            rstWeb = MyCommon.LRT_Select

            If (rstWeb.Rows.Count = 0) Then
              CurrentOffers += OfferID & ","

              'Find the name of the associated (and non-excluding) location group
              MyCommon.QueryStr = "SELECT OL.OfferID,OL.LocationGroupID,OL.Excluded,LG.Name FROM OfferLocations AS OL " & _
                                  "LEFT JOIN LocationGroups AS LG on LG.LocationGroupID=OL.LocationGroupID " & _
                                  "WHERE OL.OfferID=" & OfferID & " and OL.Excluded=0"
              rst3 = MyCommon.LRT_Select

              'So finally we have all the info we need. Now display it.
              Send("    <div class=""currentoffer"">")
              Send("    <h3 class=""name"" title=""Offer ID:" & OfferID & """ alt=""OfferID: " & OfferID & """ >" & OfferDesc & "</h3>")
              If (rst3.Rows.Count = 0 OrElse rst3.Rows(0).Item("Name") = "All Locations") Then
                Sendb("    Available at all stores.<br />")
              Else
                Send_SelectStoresLink(Handheld, Identifier, OfferID)
              End If
              Send("    Begins " & FormatDateTime(OfferStart, DateFormat.LongDate) & "<br />")
              Send("    Ends " & FormatDateTime(OfferEnd, DateFormat.LongDate) & "<br />")

              If (OfferDaysLeft > 1) Then
                Send("    It will expire in " & OfferDaysLeft & " days.<br />")
              ElseIf (OfferDaysLeft = 1) Then
                Send("    It will expire tomorrow.<br />")
              ElseIf (OfferDaysLeft = 0) Then
                Send("    It expires today.<br />")
              ElseIf (OfferDaysLeft = -1) Then
                Send("    It expired yesterday.<br />")
              ElseIf (OfferDaysLeft < -1) Then
                Send("    It expired " & Math.Abs(OfferDaysLeft) & " days ago.<br />")
              End If
              Send("    <br class=""half"" />")
              If (InstantWin > 0 AndAlso OfferOdds > 0) Then
                Send("    <b>Odds of winning:</b> 1:" & OfferOdds & "<br />")
              End If

              Send_Accumulations(Handheld, OfferID)

              GraphicsFileName = RetrieveGraphicPath(OfferID)
              If (GraphicsFileName <> "") Then
                Send("    <center><img src=""images\" & GraphicsFileName & """ /></center>")
              End If

              PrintedMessage = RetrievePrintedMessage(OfferID)
              If (PrintedMessage <> "") Then
                Send_DetailsLink(Handheld, Identifier, OfferID)
              End If
              Send("    </div>")
              Send("    <hr />")
              Send("    <br class=""half"">")
              Send("")
            End If
          Next
        Next
      Else
        'No customer was found, so...
        Send("    <p>We can't find card number " & Request.Form("identifier") & ".</p>")
      End If
    Else
      'No identifier was found, so...
      Sendb("    <p>You must be logged in to see your offers.</p>")
    End If
    Send("    <br />")
    Send("   </div>")
  End Sub

  Public Sub Send_OptInOffers(ByRef CurrentOffers As String, ByVal Handheld As Boolean)
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
    Dim Identifier As String
    Dim PrintedMessage As String = ""
    Dim AllowOptOut As Boolean = False
    Dim OptOutOffer As Boolean = False
    Dim ExcludedFromOffer As Boolean = False
    Dim rowCount As Integer = 0
    Dim i As Integer

    Send("")
    Send("   <div id=""groups"">")

    Identifier = IIf(Request.Form("identifier") <> "", Request.Form("identifier"), "")
    If (Identifier = "") Then
      Identifier = IIf(Request.QueryString("identifier") <> "", Request.QueryString("identifier"), "")
    End If

    'First off, we check to see if there's an identifier in the URL
    If (Identifier <> "") Then

      ' There is, so first off we grab the customer's information
      If (IsNumeric(Identifier)) Then
        'The identifier the customer supplied is all numbers, so assume it's a card number
        MyCommon.QueryStr = "SELECT C.CustomerPK,C.PrimaryExtID,C.Employee,C.CardStatusID,CE.Email " & _
                            "FROM Customers AS C " & _
                            "LEFT JOIN CustomerExt AS CE on CE.CustomerPK=C.CustomerPK " & _
                            "WHERE PrimaryExtID='" & Identifier & "'"
      Else
        'The identifier the customer supplied isn't just numbers, so assume an email address
        MyCommon.QueryStr = "SELECT C.CustomerPK,C.PrimaryExtID,C.Employee,C.CardStatusID,CE.Email " & _
                            "FROM Customers AS C " & _
                            "LEFT JOIN CustomerExt AS CE on CE.CustomerPK=C.CustomerPK " & _
                            "WHERE Email='" & Identifier & "'"
      End If
      rst = MyCommon.LXS_Select

      If (rst.Rows.Count > 0) Then
        'A customer was found, so let's assign values to variables
        CustomerPK = rst.Rows(0).Item("CustomerPK")
        ExtID = rst.Rows(0).Item("PrimaryExtID")
        Employee = rst.Rows(0).Item("Employee")

        Send("    <h2>Active Groups</h2>")

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

        MyCommon.QueryStr = "select I.IncentiveID, I.IncentiveName, I.Description, I.StartDate, I.EndDate, ICG.CustomerGroupID, ICG.ExcludedUsers, D.OutputID as RewardGroup, I.AllowOptOut, RO.RewardOptionID " & _
                            "from CPE_Incentives I inner join OfferIDs OID on I.IncentiveID = OID.OfferID " & _
                            "inner join CPE_RewardOptions RO on RO.IncentiveID = I.IncentiveID and RO.TouchResponse=0 and RO.Deleted=0 " & _
                            "inner join CPE_Deliverables D on D.RewardOptionID = RO.RewardOptionID and D.Deleted = 0 and DeliverableTypeID = 5 and RewardOptionPhase=3 " & _
                            "inner join CPE_IncentiveCustomerGroups ICG on ICG.RewardOptionID=RO.RewardOptionID and ICG.Deleted = 0 " & _
                            "   and ICG.CustomerGroupID in (" & CustomerGroups.ToString & ")and ICG.ExcludedUsers = 0 " & _
                            "where I.Deleted=0 And I.StatusFlag = 0 and I.EndDate >= '" & Today.ToString & "' and OID.EngineID=3 " & CPECurrentOffers & ";"

        'Send(MyCommon.QueryStr)
        'Exit Sub
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

            If (Not ExcludedFromOffer And ((Not OptOutOffer) Or (OptOutOffer And AllowOptOut))) Then
              'Find the name of the associated (and non-excluding) location group
              MyCommon.QueryStr = "SELECT OL.OfferID,OL.LocationGroupID,OL.Excluded,LG.Name FROM OfferLocations AS OL " & _
                                  "LEFT JOIN LocationGroups AS LG on LG.LocationGroupID=OL.LocationGroupID " & _
                                  "WHERE OL.OfferID=" & OfferID & " and OL.Excluded=0"
              rst3 = MyCommon.LRT_Select

              'So finally we have all the info we need. Now display it.
              Send("    <div class=""offer"">")
              Send("    <h3 class=""name"" title=""Offer ID:" & OfferID & """ alt=""OfferID: " & OfferID & """ >" & OfferDesc & "</h2>")
              Send("    Begins " & FormatDateTime(OfferStart, DateFormat.LongDate) & "<br />")
              Send("    Ends " & FormatDateTime(OfferEnd, DateFormat.LongDate) & "<br />")

              If (OfferDaysLeft > 1) Then
                Send("    It will expire in " & OfferDaysLeft & " days.<br />")
              ElseIf (OfferDaysLeft = 1) Then
                Send("    It will expire tomorrow.<br />")
              ElseIf (OfferDaysLeft = 0) Then
                Send("    It expires today.<br />")
              ElseIf (OfferDaysLeft = -1) Then
                Send("    It expired yesterday.<br />")
              ElseIf (OfferDaysLeft < -1) Then
                Send("    It expired " & Math.Abs(OfferDaysLeft) & " days ago.<br />")
              End If
              Send("    <br class=""half"" />")
              If (OfferOdds > 0) Then
                Send("    <b>Odds of winning:</b> 1:" & OfferOdds & "<br />")
              End If

              GraphicsFileName = RetrieveGraphicPath(OfferID)
              If (GraphicsFileName <> "") Then
                Send("    <center><img src=""images\" & GraphicsFileName & """ /></center>")
              End If

              Send("    <table cellspacing=""0"" cellpadding=""0"" style=""width:100%;"">")
              Send("      <tr><td><form action=""nhome.aspx"" method=""post"">")
              Send("          <input type=""hidden"" name=""identifier"" id=""identifier"" value=""" & Identifier & """ />")
              Send("          <input type=""hidden"" name=""customerPK"" id=""customerPK"" value=""" & CustomerPK & """ />")
              Send("          <input type=""hidden"" name=""customergroupID"" id=""customergroupID"" value=""" & RewardGroupID & """ />")
              If (OptOutOffer) Then
                Send("          <input type=""submit"" id=""remove"" name=""remove"" value=""Remove"" /><br />")
              Else
                Send("          <input type=""submit"" id=""join"" name=""join"" value=""Join"" /><br />")
              End If
              Send("    </form></td><td style=""text-align:right;"">")
              PrintedMessage = RetrievePrintedMessage(OfferID)
              If (PrintedMessage <> "") Then
                Send_DetailsLink(Handheld, Identifier, OfferID)
              End If
              Send("    </td></tr></table>")
              Send("    </div>")
              Send("    <hr />")
              Send("")
            End If
          Next
        Else
          Send("    <br /><p>Currently there are no eligible offers available to join.</p>")
        End If
      Else
        'No customer was found, so...
        Send("    <p>We can't find card number " & Identifier & ".</p>")
      End If
    Else
      'No identifier was found, so...
      Sendb("    <p>You must be logged in to see your offers.</p>")
    End If
    Send("    <br />")
    Send("   </div>")
  End Sub

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
        GraphicsFileName += If(MyCommon.NZ(rst.Rows(0).Item("ImageType"), 1) = 2, "gif", "jpg")
        GraphicsFilePath += GraphicsFileName
        If (File.Exists(GraphicsFilePath)) Then
          GraphicsNewFilePath = Server.MapPath("nhome.aspx")
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

  Private Function IsAnimatedGIF(ByVal GraphicsFilePath As String) As Boolean
    Dim IsAnimGif As Boolean = False

    If (File.Exists(GraphicsFilePath)) Then
      'Create an image object from a file on disk
      Dim MyImage As System.Drawing.Image = System.Drawing.Image.FromFile(GraphicsFilePath)

      'Create a new FrameDimension object from this image
      Dim FrameDimensions As System.Drawing.Imaging.FrameDimension = New System.Drawing.Imaging.FrameDimension(MyImage.FrameDimensionsList(0))

      'Determine the number of frames in the image
      'Note that all images contain at least 1 frame, but an animated GIF
      'will contain more than 1 frame.
      Dim NumberOfFrames As Integer = MyImage.GetFrameCount(FrameDimensions)

      IsAnimGif = (NumberOfFrames > 1)
    End If

    Return IsAnimGif
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

  Public Sub Send_Stores(ByVal Handheld As Boolean)
    Dim rst As System.Data.DataTable
    Dim row As System.Data.DataRow
    Dim LocationID As Long
    Dim ExtLocationCode As String
    Dim LocationName As String
    Dim Address1 As String
    Dim Address2 As String
    Dim City As String
    Dim State As String
    Dim Zip As String
    Dim CountryID As Integer
    Dim CountryName As String
    Dim TestingLocation As Boolean
    Dim ContactName As String
    Dim PhoneNumber As String

    MyCommon.Open_LogixRT()
    MyCommon.Open_LogixXS()

    If (Request.QueryString("zip") <> "") Then
      MyCommon.QueryStr = "SELECT L.LocationID,L.ExtLocationCode,L.LocationName,L.Address1,L.Address2,L.City,L.State,L.Zip,L.CountryID,C.CountryName," & _
                          "L.TestingLocation,L.Deleted,L.StatusFlag,L.ContactName,L.PhoneNumber FROM Locations AS L WITH (NoLock) " & _
                          "INNER JOIN Countries AS C ON L.CountryID=C.CountryID " & _
                          "WHERE Deleted = 0 AND TestingLocation = 0 AND Zip = '" & Request.QueryString("zip") & "' ORDER BY ExtLocationCode"
      rst = MyCommon.LRT_Select
      Send("  <div id=""storeslist"">")
      Send("   <h1>Stores</h1>")
      Send("   <p>The ZIP Code you searched for was " & Request.QueryString("zip") & ".<br />")
      If (rst.Rows.Count = 0) Then
        Send("   There are no Adam's Natural Foods stores in that area.</p>")
      ElseIf (rst.Rows.Count = 1) Then
        Send("   There is one Adam's Natural Foods store in that area:</p>")
      Else
        Send("   There are " & rst.Rows.Count & " Adam's Natural Foods stores in that area:</p>")
      End If
    Else
      MyCommon.QueryStr = "SELECT L.LocationID,L.ExtLocationCode,L.LocationName,L.Address1,L.Address2,L.City,L.State,L.Zip,L.CountryID,C.CountryName," & _
                          "L.TestingLocation,L.Deleted,L.StatusFlag,L.ContactName,L.PhoneNumber FROM Locations AS L WITH (NoLock) " & _
                          "INNER JOIN Countries AS C ON L.CountryID=C.CountryID " & _
                          "WHERE Deleted = 0 AND TestingLocation = 0 ORDER BY ExtLocationCode"
      rst = MyCommon.LRT_Select
      Send("  <div id=""storeslist"">")
      Send("   <h1>Stores</h1>")
      Send("   <p>Adam's Natural Foods currently has " & rst.Rows.Count & " retail locations.</p>")
    End If

    Send("  <form action=""stores.aspx"" method=""get"">")
    Send("   Find a store by ZIP Code:<br />")
    Send("   <input id=""zip"" name=""zip"" /> <input id=""search"" name=""search"" type=""submit"" value=""Search"" /><br />")
    Send("   <br />")
    Send("  </form>")

    For Each row In rst.Rows
      LocationID = row.Item("LocationID")
      ExtLocationCode = MyCommon.NZ(row.Item("ExtLocationCode"), "")
      LocationName = MyCommon.NZ(row.Item("LocationName"), "")
      Address1 = MyCommon.NZ(row.Item("Address1"), "")
      Address2 = MyCommon.NZ(row.Item("Address2"), "")
      City = MyCommon.NZ(row.Item("City"), "")
      State = MyCommon.NZ(row.Item("State"), "")
      Zip = MyCommon.NZ(row.Item("Zip"), "")
      CountryID = MyCommon.NZ(row.Item("CountryID"), "")
      CountryName = MyCommon.NZ(row.Item("CountryName"), "")
      TestingLocation = MyCommon.NZ(row.Item("TestingLocation"), "")
      ContactName = MyCommon.NZ(row.Item("ContactName"), "")
      PhoneNumber = MyCommon.NZ(row.Item("PhoneNumber"), "")

      ' List store details
      Sendb("   <b>Store " & ExtLocationCode & "</b>")
      If (LocationName <> "") And (LocationName <> ("Store " & ExtLocationCode)) Then
        Send(" (" & LocationName & ")<br />")
      Else
        Send("<br />")
      End If
      If (Address1 <> "") Then
        Send("   " & Address1 & "<br />")
      End If
      If (Address2 <> "") Then
        Send("   " & Address2 & "<br />")
      End If
      If (City <> "") And (State <> "") And (Zip <> "") Then
        Send("   " & City & ", " & State & " " & Zip & "<br />")
      End If
      If (CountryID > 1) Then
        Send("   " & CountryName & "<br />")
      End If
      If (PhoneNumber <> "") Then
        Send("   " & PhoneNumber & "<br />")
      End If
      If (ContactName <> "") Then
        Send("   Manager: " & ContactName & "<br />")
      End If

      ' Create a link to Google maps
      If (City = "") Then
      Else
        If (Address2 <> "") Then
          Send("   [<a href=""http://maps.google.com/maps?f=q&hl=en&q=" & Address2.Replace(" ", "+") & ",+" & City.Replace(" ", "+") & ",+" & State & "&layer=&ie=UTF8&z=12&om=1"" target=""_blank""><b>Map</b></a>]<br />")
        Else
          Send("   [<a href=""http://maps.google.com/maps?f=q&hl=en&q=" & Address1.Replace(" ", "+") & ",+" & City.Replace(" ", "+") & ",+" & State & "&layer=&ie=UTF8&z=12&om=1"" target=""_blank""><b>Map</b></a>]<br />")
        End If
      End If
      Send("   <br />")
    Next
    Send("  </div>")
    Send("")
  End Sub

  Public Sub Send_MainEnd(ByVal Handheld As Boolean)
    Send("")
    If (Handheld) Then
      Send("  </td> <!-- End main -->")
      Send("</tr>")
    Else
      Send("</div> <!-- End main -->")
      Send("<br clear=""left"" />")
    End If
    Send("")
  End Sub

  Public Sub Send_Footer(ByVal Handheld As Boolean)
    Send("<div id=""footer"">")
    Send("</div>")
    Send("")
  End Sub

  Public Sub Send_InnerWrapEnd(ByVal Handheld As Boolean)
    Send("</div> <!-- Innerwrap ends -->")
    Send("")
  End Sub

  Public Sub Send_Legal(ByVal Handheld As Boolean)
    Send("<div id=""legal"">")
    Send("  <span id=""copyright"">")
    Send("     &copy; Copyright 2011, NCR Supermarkets, Inc.")
    Send("  </span>")
    Send("  <span id=""pptou"">")
    Send("    <a href=""http://www.ncr.com/utility/privacy_policy/ncr_privacy_policy.jsp?lang=EN"">Privacy Info/Terms of Use</a>")
    Send("  </span>")
    Send("</div>")
    Send("")
  End Sub

  Public Sub Send_BodyEnd(ByVal Handheld As Boolean)
    Send("")
    If (Handheld) Then
      Send("</table>")
    Else
      Send("<a id=""bottom"" name=""bottom""></a>")
      Send("</div> <!-- End wrap -->")
    End If
    Send("</body>")
    Send("</html>")
  End Sub

  Public Sub Send_DetailsLink(ByVal Handheld As Boolean, ByVal Identifier As String, ByVal OfferID As Long)
    If (Handheld) Then
      Send("     <a href=""prntmsg.aspx?identifier=" & Identifier & "&OfferID=" & OfferID & """>Details</a>")
    Else
      Send("     <a href=""javascript:openNamedPopup('prntmsg.aspx?identifier=" & Identifier & "&OfferID=" & OfferID & "', 'PrntMsg');"">Details</a>")
    End If
  End Sub

  Public Sub Send_PmsgDetails(ByVal Handheld As Boolean)
    Dim PrintedMessage As String = ""
    Dim OfferID As Long

    OfferID = If(Request.QueryString("OfferID") <> "", CLng(Request.QueryString("OfferID")), 0)
    If (OfferID > 0) Then
      PrintedMessage = RetrievePrintedMessage(OfferID)
      Send(GenerateMessagePreview(PrintedMessage, OfferID))
    Else
      Send("<center><b>No printed message found for this offer.</b></center>")
    End If

  End Sub

  Public Sub Send_SelectStoresLink(ByVal Handheld As Boolean, ByVal Identifier As String, ByVal OfferID As Long)
    Send("")
    If (Handheld) Then
      Send("Available at <a href=""selectstores.aspx?identifier=" & Identifier & "&OfferID=" & OfferID & """>select stores</a>.")
    Else
      Send("Available at <a href=""javascript:openNamedPopup('selectstores.aspx?identifier=" & Identifier & "&OfferID=" & OfferID & "', 'SelStore');"">select stores</a>.")
    End If
  End Sub

  Public Sub Send_SelectStores(ByVal Handheld As Boolean)
    Dim Identifier As String = ""
    Dim OfferID As String = ""
    Dim rst As System.Data.DataTable
    Dim row As System.Data.DataRow
    Dim LocationID As Long
    Dim ExtLocationCode As String
    Dim LocationName As String
    Dim Address1 As String
    Dim Address2 As String
    Dim City As String
    Dim State As String
    Dim Zip As String
    Dim CountryID As String
    Dim CountryName As String
    Dim TestingLocation As String
    Dim ContactName As String
    Dim PhoneNumber As String

    Identifier = Request.QueryString("identifier")
    OfferID = Request.QueryString("OfferID")

    MyCommon.Open_LogixRT()
    MyCommon.QueryStr = "SELECT L.LocationID,L.ExtLocationCode,L.LocationName,L.Address1,L.Address2,L.City,L.State,L.Zip,L.CountryID,C.CountryName, " & _
                        "L.TestingLocation,L.Deleted,L.StatusFlag,L.ContactName,L.PhoneNumber " & _
                        "FROM Locations AS L WITH (NoLock) inner join countries c on L.CountryID = c.CountryID " & _
                        "inner join LocGroupItems LGI with (nolock) on L.LocationID = LGI.LocationID " & _
                        "inner join LocationGroups LG with (nolock) on LGI.LocationGroupID = LG.LocationGroupID " & _
                        "inner join OfferLocations OL with (nolock) on OL.LocationGroupID = LG.LocationGroupID " & _
                        "where OL.OfferID = " & OfferID & " and OL.Deleted = 0 and LG.Deleted = 0 and LGI.Deleted =0 and L.Deleted = 0 and L.TestingLocation = 0 ORDER BY ExtLocationCode;"

    rst = MyCommon.LRT_Select
    Send("  <div id=""storeslist"">")
    Send("   <h1>Selected Stores</h1>")

    For Each row In rst.Rows
      LocationID = row.Item("LocationID")
      ExtLocationCode = MyCommon.NZ(row.Item("ExtLocationCode"), "")
      LocationName = MyCommon.NZ(row.Item("LocationName"), "")
      Address1 = MyCommon.NZ(row.Item("Address1"), "")
      Address2 = MyCommon.NZ(row.Item("Address2"), "")
      City = MyCommon.NZ(row.Item("City"), "")
      State = MyCommon.NZ(row.Item("State"), "")
      Zip = MyCommon.NZ(row.Item("Zip"), "")
      CountryID = MyCommon.NZ(row.Item("CountryID"), "")
      CountryName = MyCommon.NZ(row.Item("CountryName"), "")
      TestingLocation = MyCommon.NZ(row.Item("TestingLocation"), "")
      ContactName = MyCommon.NZ(row.Item("ContactName"), "")
      PhoneNumber = MyCommon.NZ(row.Item("PhoneNumber"), "")

      ' List store details
      Sendb("   <b>Store " & ExtLocationCode & "</b>")
      If (LocationName <> "") And (LocationName <> ("Store " & ExtLocationCode)) Then
        Send(" (" & LocationName & ")<br />")
      Else
        Send("<br />")
      End If
      If (Address1 <> "") Then
        Send("   " & Address1 & "<br />")
      End If
      If (Address2 <> "") Then
        Send("   " & Address2 & "<br />")
      End If
      If (City <> "") And (State <> "") And (Zip <> "") Then
        Send("   " & City & ", " & State & " " & Zip & "<br />")
      End If
      If (CountryID > 1) Then
        Send("   " & CountryName & "<br />")
      End If
      If (PhoneNumber <> "") Then
        Send("   " & PhoneNumber & "<br />")
      End If
      If (ContactName <> "") Then
        Send("   Manager: " & ContactName & "<br />")
      End If

      ' Create a link to Google maps
      If (City = "") Then
      Else
        If (Address2 <> "") Then
          Send("   [<a href=""http://maps.google.com/maps?f=q&hl=en&q=" & Address2.Replace(" ", "+") & ",+" & City.Replace(" ", "+") & ",+" & State & "&layer=&ie=UTF8&z=12&om=1"" target=""_blank""><b>Map</b></a>]<br />")
        Else
          Send("   [<a href=""http://maps.google.com/maps?f=q&hl=en&q=" & Address1.Replace(" ", "+") & ",+" & City.Replace(" ", "+") & ",+" & State & "&layer=&ie=UTF8&z=12&om=1"" target=""_blank""><b>Map</b></a>]<br />")
        End If
      End If
      Send("   <br />")
    Next
    MyCommon.Close_LRTsp()
  End Sub

  Private Function GenerateMessagePreview(ByVal newText As String, ByVal OfferID As Integer) As String
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
    Dim tempDate As Date

    ' this 2d array will hold the replace and replacement for real time engine tags just add or modify
    ' here
    Dim SubTags(,) As String = { _
                                {"|CUSTOMERID|", "|TSD|", "|LYTS|", "|CURRDATE|", "|OFFERSTART|", "|OFFEREND|", "|TOTALPOIINTS|", "|ACCUMANT|", "|REMAINAMT|"}, _
                                {"###################", "000.00", "000.00", Now.ToString("M/d/yyyy"), "xx/xx/xxxx", "xx/xx/xxxx", "xx", "000.00", "000.00"}}

    ' replace the offer start and offer end dates with the actual dates if the tags are present in the message
    If (OfferID > 0 AndAlso (newText.IndexOf("|OFFERSTART|") > -1 OrElse newText.IndexOf("|OFFEREND|") > -1)) Then
      If (MyCommon.LRTadoConn.State <> Data.ConnectionState.Open) Then MyCommon.Open_LogixRT()
      MyCommon.QueryStr = "select StartDate, EndDate from CPE_Incentives I with (NoLock) where I.IncentiveID = " & OfferID & " " & _
                          "union " & _
                          "select ProdStartDate as StartDate, ProdEndDate as EndDate from Offers O with (NoLock) where O.OfferID = " & OfferID
      rst = MyCommon.LRT_Select
      If (rst.Rows.Count > 0) Then
        If (Not IsDBNull(rst.Rows(0).Item("StartDate"))) Then
          tempDate = rst.Rows(0).Item("StartDate")
          SubTags(1, 4) = tempDate.ToString("M/d/yyyy")
        End If
        If (Not IsDBNull(rst.Rows(0).Item("EndDate"))) Then
          tempDate = rst.Rows(0).Item("EndDate")
          SubTags(1, 5) = tempDate.ToString("M/d/yyyy")
        End If
      Else
        SubTags(1, 4) = "OfferStart1"
        SubTags(1, 5) = "OfferEnd1"
      End If
    Else
      SubTags(1, 4) = "OfferStart2"
      SubTags(1, 5) = "OfferEnd2"
    End If

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

  Public Sub Send_Pmsg(ByVal Handheld As Boolean)
    Dim PrintedMessage As String = ""
    Dim OfferID As Long

    OfferID = If(Request.QueryString("OfferID") <> "", CLng(Request.QueryString("OfferID")), 0)
    If (OfferID > 0) Then
      PrintedMessage = RetrievePrintedMessage(OfferID)
      If (PrintedMessage <> "") Then
        Send(GenerateMessagePreview(PrintedMessage, OfferID))
      End If
    Else
      Send("<center><b>No printed message found for this offer.</b></center>")
    End If

  End Sub

  Public Sub Send_OffersForGroup(ByVal Handheld As Boolean)
    Dim rst2 As System.Data.DataTable
    Dim rst3 As System.Data.DataTable
    Dim rstWeb As System.Data.DataTable
    Dim rstExcluded As System.Data.DataTable
    Dim row2 As System.Data.DataRow
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
    Dim retString As New StringBuilder

    CustomerGroupID = MyCommon.Extract_Val(Request.QueryString("cg"))

    If (CustomerGroupID > 0) Then
      If (MyCommon.LRTadoConn.State <> Data.ConnectionState.Open) Then MyCommon.Open_LogixRT()

      ' find all the offers for the given web site offer reward customer group
      MyCommon.QueryStr = "select distinct O.OfferID,O.ExtOfferID,O.IsTemplate,O.CMOADeployStatus,O.StatusFlag,O.OddsOfWinning,O.InstantWin, " & _
                          "O.Name,O.Description,O.ProdStartDate,O.ProdEndDate,LinkID,OID.EngineID from Offers as O " & _
                          "LEFT JOIN OfferConditions as OC on OC.OfferID=O.OfferID " & _
                          "INNER JOIN OfferIDs as OID on OID.OfferID=O.OfferID " & _
                          "where O.IsTemplate=0 and O.Deleted=0 and O.CMOADeployStatus=1 and OC.Deleted=0 and OC.ConditionTypeID=1 " & _
                          "and O.DisabledOnCFW=0 and ProdEndDate>'" & Today.AddDays(-1).ToString & "' and LinkID=" & CustomerGroupID & _
                          "union all " & _
                          "select distinct I.IncentiveID,I.ClientOfferID,I.IsTemplate,I.CPEOADeployStatus,I.StatusFlag,0 as OddsOfWinning,0 as InstantWin, " & _
                          "I.IncentiveName, Convert(nvarchar(2000),I.Description) as Description,I.StartDate,I.EndDate,ICG.CustomerGroupID,OID.EngineID from CPE_Incentives I " & _
                          "LEFT JOIN CPE_RewardOptions RO on I.IncentiveID=RO.IncentiveID and RO.TouchResponse=0 and RO.Deleted=0 " & _
                          "LEFT JOIN CPE_IncentiveCustomerGroups ICG on RO.RewardOptionID=ICG.RewardOptionID and ICG.ExcludedUsers = 0 " & _
                          "INNER JOIN OfferIDs as OID on OID.OfferID=I.IncentiveID " & _
                          "where (I.IsTemplate=0 and I.Deleted=0 and ICG.Deleted=0) " & _
                          "and I.DisabledOnCFW=0 and I.EndDate>'" & Today.AddDays(-1).ToString & "' and CustomerGroupID=" & CustomerGroupID
      rst2 = MyCommon.LRT_Select

      If (rst2.Rows.Count > 0) Then
        retString.Append("<h2>Associated Offers</h2>")
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
                              "where OfferID = " & OfferID & " and ExcludedID in (" & CustomerGroupID & ") " & _
                              "union " & _
                              "select CustomerGroupID as ExcludedID from CPE_IncentiveCustomerGroups ICG with (NoLock) " & _
                              "inner join CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID = ICG.RewardOptionID " & _
                              "where ICG.Deleted=0 and RO.Deleted=0 and RO.IncentiveID = " & OfferID & " and ExcludedUsers=1 " & _
                              "and CustomerGroupID in (" & CustomerGroupID & ");"
          rstExcluded = MyCommon.LRT_Select

          If (rstWeb.Rows.Count = 0 AndAlso rstExcluded.Rows.Count = 0) Then

            'Find the name of the associated (and non-excluding) location group
            MyCommon.QueryStr = "SELECT OL.OfferID,OL.LocationGroupID,OL.Excluded,LG.Name FROM OfferLocations AS OL " & _
                                "LEFT JOIN LocationGroups AS LG on LG.LocationGroupID=OL.LocationGroupID " & _
                                "WHERE OL.OfferID=" & OfferID & " and OL.Excluded=0"
            rst3 = MyCommon.LRT_Select

            retString.Append("<div class=""offer"" id=""offer" & OfferID & """>" & vbCrLf)
            retString.Append("  <h3 class=""name"" alt=""Offer# " & OfferID & ", CGroup# " & CustomerGroupID & """ title=""Offer " & OfferID & ", CGroup " & CustomerGroupID & """>" & OfferName & "</h3>" & vbCrLf)
            GraphicsFileName = RetrieveGraphicPath(OfferID)
            If (GraphicsFileName <> "") Then
              retString.Append("    <img src=""images\" & GraphicsFileName & """ align=""right"" />")
            End If
            retString.Append("  <div class=""description"">" & vbCrLf)
            retString.Append("    " & OfferDesc & vbCrLf)
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
            retString.Append("</div>")
            retString.Append("<br />")
            retString.Append("<hr />")
            retString.Append("")
          End If
        Next
      End If

    End If

    Send(retString.ToString)
  End Sub

End Class