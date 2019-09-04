<%@ Page Language="vb" Debug="true" CodeFile="cwCB.vb" Inherits="cwCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.commonShared" %>
<%
  Dim CopientFileName As String = Page.Request.Url.Segments(Page.Request.Url.Segments.GetUpperBound(0))
  Dim CopientFileVersion As String = "7.3.1.138972" 
  Dim CopientProject As String = "Copient Logix"
  Dim CopientNotes As String = ""
  
  Dim MyCommon As New Copient.CommonInc
  Dim MyCryptLib As New Copient.CryptLib
  Dim Logix As New Copient.LogixInc
  Dim rst As DataTable
  Dim row As DataRow
  Dim CurrentOffers As String = ""
  Dim infoMessage As String = ""
  Dim Popup As Boolean = False
  Dim Framed As Boolean = False
  Dim Handheld As Boolean = False
  Dim OfferID As Integer
  Const DEFAULT_USER_ID As Integer = 1
  Dim PrimaryExtID as string = ""
  Dim CustomerPK As Long = 0

  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "home.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  
  'First things first - determine if they should be allowed to login
  If (Request.Form("password") <> "" And Request.Form("identifier") <> "" And Session("customerpk") <> vbNull) Then
    Session.Remove("customerpk")

    'Time to see if the password is correct for the user logging in
    'There is, so first off we grab the customer's information
    If (IsNumeric(Request.Form("identifier"))) Then
      PrimaryExtID=  MyCommon.Pad_ExtCardID(Request.Form("identifier"), CardTypes.CUSTOMER)
      'The identifier the customer supplied is all numbers, so assume it's a card number
            MyCommon.QueryStr = "select top 1 CustomerPK from CardIDs with (NoLock) where ExtCardID='" & MyCryptLib.SQL_StringEncrypt(PrimaryExtID) & "'"
      rst = MyCommon.LXS_Select
      If rst.Rows.Count > 0 Then
        CustomerPK = MyCommon.NZ(rst.Rows(0).Item("CustomerPK"), 0)
      End If
    Else
      'The identifier the customer supplied isn't just numbers, so assume an email address
            MyCommon.QueryStr = "select CustomerPK from CustomerExt with (NoLock) " & _
                                "  WHERE Email='" & MyCryptLib.SQL_StringEncrypt(Request.Form("identifier")) & "'"
      rst = MyCommon.LXS_Select
      If rst.Rows.Count > 0 Then
        CustomerPK = MyCommon.NZ(rst.Rows(0).Item("CustomerPK"), 0)
      End If
    End If

    If CustomerPK > 0 Then
      MyCommon.QueryStr = "select CustomerPK from Customers with (NoLock) where CustomerPK=" & CustomerPK & " " & _
                          "  and Password ='" & MyCryptLib.SQL_StringEncrypt(Request.Form("password")) & "'"
      rst = MyCommon.LXS_Select
      If rst.Rows.Count > 0 Then
        'We've retrieved a record, so we're logged in
        'Session("customerpk") = rst.Rows(0).Item("CustomerPK")
        Session.Add("customerpk", CustomerPK)
      End If
    End If
    
  ElseIf (Request.Form("opt-type") <> "") Then
  ElseIf (Request.QueryString("Logout") = "logout") Then
    Session.Abandon()
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "index.aspx")
  Else
    Response.Status = "301 Moved Permanently"
    Response.AddHeader("Location", "index.aspx")
  End If
  
  ' Find the Language preference for the Logix Default User to use for logging purposes
  MyCommon.QueryStr = "select LanguageID from AdminUsers where AdminUserID=1;"
  rst = MyCommon.LRT_Select
  If (rst.Rows.Count > 0) Then
    LanguageID = MyCommon.NZ(rst.Rows(0).Item("LanguageID"), 1)
  Else
    LanguageID = 1
  End If
  
  'Manage the "optin" and "optout" actions
  
  If Request.Form("opt-type") <> "" Then
    Dim HouseholdingEnabled As Boolean = IIf(MyCommon.Fetch_SystemOption(50) = 1, True, False)
    Dim IDType As Integer = 0
   
    Dim ClientUserID1 As String = ""
    Dim CustomerGroupID As Long = Request.Form("opt-cgroupid")
    Dim ExcludedGroupID As Long = MyCommon.Extract_Val(Request.Form("opt-xgroupid"))
    Dim RewardGroupID As Long = MyCommon.NZ(Request.Form("opt-rgroupid"), 0)
    Dim CustomerTypeID As Integer = MyCommon.Extract_Val(Request.Form("opt-cardtypeid"))
    Dim ROID As Long = 0
    Dim EngineID As Integer = -1
    Dim ConditionID As Integer = -1
    ClientUserID1 = MyCommon.Pad_ExtCardID(MyCommon.Extract_Val(Request.Form("opt-extid")), CardTypes.CUSTOMER)
   
    If Request.Form("opt-type") = "out" Then
      If (CustomerGroupID <= 2) Then
        ' THIS CODE SHOULD NOT BE NEEDED AS YOU ARE NOW OPTING OUT OF THE REWARD CUSTOMER GROUP NOT THE INDIVIDUAL OFFER      
        '  If (ExcludedGroupID = 0) Then
        '    ' find the excluded customer group for this offer, if one doesn't exist
        '    ' then create one.
        '    OfferID = MyCommon.Extract_Val(Request.Form("opt-offerid"))
          
        '    MyCommon.QueryStr = "select 0 as EngineID, ExcludedID, ConditionID, -1 as ROID from OfferConditions OC with (NoLock) " & _
        '                        "where OfferID = " & OfferID & " and Deleted=0 and ConditionTypeID = 1 and ExcludedID > 0 " & _
        '                        "union all " & _
        '                        "select 2 as EngineID, ICG.CustomerGroupID as ExcludedID, -1 as ConditionID, ICG.RewardOptionID as ROID from CPE_IncentiveCustomerGroups ICG with (NoLock) " & _
        '                        "inner join CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID = ICG.RewardOptionID " & _
        '                        "where ICG.Deleted = 0 and RO.Deleted = 0 and ICG.ExcludedUsers = 1 and RO.IncentiveID = " & OfferID & ";"
        '    rst = MyCommon.LRT_Select
        '    If (rst.Rows.Count > 0) Then
        '      ExcludedGroupID = MyCommon.NZ(rst.Rows(0).Item("ExcludedID"), 0)
        '      ROID = MyCommon.NZ(rst.Rows(0).Item("ROID"), -1)
        '      EngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), -1)
        '      ConditionID = MyCommon.NZ(rst.Rows(0).Item("ConditionID"), 0)
        '    Else
        '      MyCommon.QueryStr = "select EngineID from OfferIDs with (NoLock) where OfferID=" & OfferID
        '      rst = MyCommon.LRT_Select
        '      If (rst.Rows.Count > 0) Then
        '        EngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), -1)
        '      End If
            
        '      ' need to create a new customer group to use for the excluded customer group of the offer
        '      MyCommon.QueryStr = "dbo.pt_CustomerGroups_Insert"
        '      MyCommon.Open_LRTsp()
        '      MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 200).Value = "Excluded customers from offer " & OfferID
        '      MyCommon.LRTsp.Parameters.Add("@CustomerGroupID", SqlDbType.Int).Direction = ParameterDirection.Output
        '      MyCommon.LRTsp.ExecuteNonQuery()
        '      ExcludedGroupID = MyCommon.LRTsp.Parameters("@CustomerGroupID").Value
        '      MyCommon.Activity_Log(4, ExcludedGroupID, DEFAULT_USER_ID, Copient.PhraseLib.Lookup("history.cgroup-create", LanguageID))
        '    End If

        '    If (ExcludedGroupID > 0) Then
        '      ' now lets add the excluded group to the offer
        '      If (EngineID = 0) Then
        '        ' find the condition ID for the offer
        '        If (ConditionID <= 0) Then
        '          MyCommon.QueryStr = "select ConditionID from OfferConditions where ConditionTypeID=1 and OfferID=" & OfferID
        '          rst = MyCommon.LRT_Select
        '          If (rst.Rows.Count > 0) Then
        '            ConditionID = MyCommon.NZ(rst.Rows(0).Item("ConditionID"), -1)
        '            MyCommon.QueryStr = "update OfferConditions with (RowLock) set ExcludedID=" & ExcludedGroupID & _
        '                                " where ConditionID=" & ConditionID
        '            MyCommon.LRT_Execute()
        '          End If
        '        End If
        '      ElseIf (EngineID = 2) Then
        '        ' find the ROID for the offer
        '        If (ROID <= 0) Then
        '          MyCommon.QueryStr = "select RewardOptionID from CPE_RewardOptions where IncentiveID=" & OfferID
        '          rst = MyCommon.LRT_Select
        '          If (rst.Rows.Count > 0) Then
        '            ROID = MyCommon.NZ(rst.Rows(0).Item("RewardOptionID"), -1)
        '          End If
        '          MyCommon.QueryStr = "insert into CPE_IncentiveCustomerGroups with (RowLock) (RewardOptionID,CustomerGroupID,ExcludedUsers,Deleted,LastUpdate,RequiredFromTemplate) " & _
        '                              "values(" & ROID & "," & ExcludedGroupID & ",1,0,getdate(),0)"
        '          MyCommon.LRT_Execute()
        '        End If
        '      End If
        '    End If

        '  End If
        
        '  ' The group is Any Customer or Any Cardholder, so add the customer to the accompanying excluded group
        '  MyCommon.QueryStr = "dbo.pt_GroupMembership_Insert"
        '  MyCommon.Open_LXSsp()
        '  MyCommon.LXSsp.Parameters.Add("@ExtCustomerID", SqlDbType.NVarChar, 26).Value = ClientUserID1
        '  MyCommon.LXSsp.Parameters.Add("@CustomerTypeID", SqlDbType.Int).Value = 0
        '  MyCommon.LXSsp.Parameters.Add("@CustomerGroupID", SqlDbType.BigInt).Value = ExcludedGroupID
        '  MyCommon.LXSsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
        '  MyCommon.LXSsp.ExecuteNonQuery()
        '  MyCommon.Activity_Log(4, CustomerGroupID, DEFAULT_USER_ID, Copient.PhraseLib.Lookup("history.cgroup-optout", LanguageID) & " " & ClientUserID1)
        '  MyCommon.QueryStr = "Update CustomerGroups with (RowLock) set LastUpdate=getdate() where CustomerGroupID=" & ExcludedGroupID
        '  MyCommon.LRT_Execute()
      Else

        ' The group is a list of IDs, so remove the customer's from it
        MyCommon.QueryStr = "dbo.pt_GroupMembership_Delete"
        MyCommon.Open_LXSsp()
                MyCommon.LXSsp.Parameters.Add("@ExtCardID", SqlDbType.NVarChar, 26).Value = MyCryptLib.SQL_StringEncrypt(ClientUserID1)
                MyCommon.LXSsp.Parameters.Add("@CardTypeID", SqlDbType.Int).Value = CustomerTypeID
                MyCommon.LXSsp.Parameters.Add("@CustomerGroupID", SqlDbType.BigInt).Value = RewardGroupID
                'MyCommon.LXSsp.Parameters.Add("@CustomerGroupID", SqlDbType.BigInt).Value = CustomerGroupID
        MyCommon.LXSsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
        MyCommon.LXSsp.ExecuteNonQuery()
        MyCommon.Activity_Log(4, CustomerGroupID, DEFAULT_USER_ID, Copient.PhraseLib.Lookup("history.cgroup-optout", LanguageID) & " " & ClientUserID1)
        MyCommon.QueryStr = "Update CustomerGroups with (RowLock) set LastUpdate=getdate() where CustomerGroupID=" & CustomerGroupID
        MyCommon.LRT_Execute()
      End If

    ElseIf Request.Form("opt-type") = "in" Then
      MyCommon.QueryStr = "dbo.pt_GroupMembership_Insert"
      MyCommon.Open_LXSsp()
            MyCommon.LXSsp.Parameters.Add("@ExtCardID", SqlDbType.NVarChar, 26).Value = MyCryptLib.SQL_StringEncrypt(ClientUserID1)
            MyCommon.LXSsp.Parameters.Add("@CardTypeID", SqlDbType.Int).Value = CustomerTypeID
      MyCommon.LXSsp.Parameters.Add("@CustomerTypeID", SqlDbType.Int).Value = CustomerTypeID
      MyCommon.LXSsp.Parameters.Add("@CustomerGroupID", SqlDbType.BigInt).Value = RewardGroupID
      MyCommon.LXSsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
      MyCommon.LXSsp.ExecuteNonQuery()
      MyCommon.Activity_Log(4, CustomerGroupID, DEFAULT_USER_ID, Copient.PhraseLib.Lookup("history.cgroup-optin", LanguageID) & " " & ClientUserID1)
      MyCommon.QueryStr = "Update CustomerGroups with (RowLock) set LastUpdate=getdate() where CustomerGroupID=" & RewardGroupID
      MyCommon.LRT_Execute()
    End If
    MyCommon.Close_LXSsp()
  End If
  
  Send_HeadBegin(Handheld, "Customer Website: Home")
  Send_Metas(Handheld)
  Send_Links(Handheld)
%>
<script type="text/javascript" language="javascript">

document.write = function(str){ 
    var moz = !window.opera && !/Apple/.test(navigator.vendor); 
        
    // Watch for writing out closing tags, we just
    // ignore these (as we auto-generate our own)
    if ( str.match(/^<\//) ) return;
    
    // Make sure & are formatted properly, but Opera
    // messes this up and just ignores it
    if ( !window.opera )
        str = str.replace(/&(?![#a-z0-9]+;)/g, "&amp;");
    
    // Watch for when no closing tag is provided
    // (Only does one element, quite weak)
    str = str.replace(/<([a-z]+)(.*[^\/])>$/, "<$1$2></$1>");
    
    // Mozilla assumes that everything in XHTML innerHTML
    // is actually XHTML - Opera and Safari assume that it's XML
    if ( !moz )
        str = str.replace(/(<[a-z]+)/g, "$1 xmlns='http://www.w3.org/1999/xhtml'");
    
    // The HTML needs to be within a XHTML element
    var div = document.createElementNS("http://www.w3.org/1999/xhtml","div");
    div.innerHTML = str;
    
    // Find the last element in the document
    var pos;
    
    // Opera and Safari treat getElementsByTagName("*") accurately
    // always including the last element on the page
    if ( !moz ) {
        pos = document.getElementsByTagName("*");
        pos = pos[pos.length - 1];
        // Mozilla does not, we have to traverse manually
    } else {
        pos = document;
        while ( pos.lastChild && pos.lastChild.nodeType == 1 )
            pos = pos.lastChild;
    }
    
    // Add all the nodes in that position
    var nodes = div.childNodes;
    while ( nodes.length )
        pos.parentNode.appendChild( nodes[0] );
};

function xmlhttpPost(strURL,inMode,frameName) {
    var xmlHttpReq = false;
    var xmlHttpReq2 = false;
    var self = this;
    var processingPage = "<div class=\"loading\"><br /><img id=\"clock\" src=\"../images/clock22.png\" \/><br \/>" + 'Loading ... please wait.<\/div>';
    
    //alert(strURL + ", " + inMode + ", " + frameName)
    
    document.getElementById(frameName).innerHTML = processingPage;
    
    if (inMode=='cust'){
        // Mozilla/Safari
        if (window.XMLHttpRequest) {
            self.xmlHttpReq = new XMLHttpRequest();
        }
        // IE
        else if (window.ActiveXObject) {
            self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
            req = self.xmlHttpReq
        }
        self.xmlHttpReq.open('POST', strURL, true);
        self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
        self.xmlHttpReq.onreadystatechange = function() {
            if (self.xmlHttpReq.readyState == 4) {
                updatepage(self.xmlHttpReq.responseText, frameName);
            }
        }
        self.xmlHttpReq.send(getquerystring(inMode));
    }
    else if(inMode=='info'){
            // Mozilla/Safari
        if (window.XMLHttpRequest) {
            self.xmlHttpReq2 = new XMLHttpRequest();
        }
        // IE
        else if (window.ActiveXObject) {
            self.xmlHttpReq2 = new ActiveXObject("Microsoft.XMLHTTP");
            req = self.xmlHttpReq2
        }
        self.xmlHttpReq2.open('POST', strURL, true);
        self.xmlHttpReq2.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
        self.xmlHttpReq2.onreadystatechange = function() {
            if (self.xmlHttpReq2.readyState == 4) {
                updatepage(self.xmlHttpReq2.responseText, frameName);
            }
        }
        self.xmlHttpReq2.send(getquerystring(inMode));
    }
}

function getquerystring(inMode) {
    qstr = 'mode=' + inMode; //+ escape(word);   NOTE: no '?' before querystring
    qstr = qstr + '&Transform=HTML';
    return qstr;
}

function updatepage(str, frameName) {
    document.getElementById(frameName).innerHTML = str;
}
</script>
<%
  Send_HeadEnd(Handheld)
  Send_BodyBegin(Handheld, Popup)
  Send_InnerWrapBegin(Handheld)
  Send_Logo(Handheld)
  Send_Menu(Handheld)
  Send_Submenu(Handheld)
  Send_SidebarBegin(Handheld)
  Send_SidebarEnd(Handheld)
  Send_Gutter(Handheld)
  Send_MainBegin(Handheld)
  Send_MainEnd(Handheld)
  Send_Footer(Handheld)
  Send_InnerWrapEnd(Handheld)
  Send_Legal(Handheld)
%>
<script type="text/javascript" language="javascript">
  xmlhttpPost('cwfeed.aspx',"info","sidebar")
  xmlhttpPost('cwfeed.aspx',"cust","main")
</script>
<%
  Send_BodyEnd(Handheld)
  
done:
  MyCommon.Close_LogixXS()
  MyCommon.Close_LogixRT()
  Logix = Nothing
  MyCommon = Nothing
%>
