<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: CPEoffer-con-customer.aspx 
  ' *~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' * Copyright © 2002 - 2009.  All rights reserved by:
  ' *
  ' * NCR Corporation
  ' * 1435 Win Hentschel Boulevard
  ' * West Lafayette, IN  47906
  ' * voice: 888-346-7199  fax: 765-464-1369
  ' *~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  ' *
  ' * PROJECT : NCR Advanced Marketing Solution
  ' *
  ' * MODULE  : Logix
  ' *
  ' * PURPOSE : 
  ' *
  ' * NOTES   : 
  ' *
  ' * Version : 7.3.1.138972 
  ' *
  ' *****************************************************************************
  
  Dim CopientFileName As String = Page.Request.Url.Segments(Page.Request.Url.Segments.GetUpperBound(0))
  Dim CopientFileVersion As String = "7.3.1.138972" 
  Dim CopientProject As String = "Copient Logix"
  Dim CopientNotes As String = ""
  
  Dim AdminUserID As Long
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim rst As DataTable
  'Dim rstBannerCgs As DataTable = Nothing
  Dim dst As DataTable
  Dim row As DataRow
  Dim OfferID As Long
  Dim Name As String = ""
  Dim ConditionID As String
  Dim IsTemplate As Boolean = False
  Dim FromTemplate As Boolean = False
  Dim Disallow_Edit As Boolean = True
  Dim Household As Boolean = False
  Dim DisabledAttribute As String = ""
  Dim i As Integer
  Dim roid As Integer
  Dim historyString As String
  Dim CloseAfterSave As Boolean = False
  Dim Ids() As String
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim RequireCG As Boolean = False
  Dim HasRequiredCG As Boolean = False
  Dim EngineID As Integer = 2
  Dim BannersEnabled As Boolean = True
  Dim NewCardholdersID As Integer = 0
  Dim FullListSelect As New StringBuilder()
  Dim AllCAM As Integer = 0
  Dim AnyCustomerEnabled As Boolean = False
  Dim RECORD_LIMIT As Integer = GroupRecordLimit '500
  Dim EligibleIncludedcustomergroups As String = String.Empty
  Dim EligibleExcludedcustomergroups As String = String.Empty
  Dim IsEligibilityConditionExistForOffer As String = "False"
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "CPEoffer-con-customer.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  OfferID = CLng(Server.HtmlEncode(Request.QueryString("OfferID")))
  ConditionID = Server.HtmlEncode(Request.QueryString("ConditionID"))
  If (Request.QueryString("EngineID") <> "") Then
    EngineID = MyCommon.Extract_Val(Server.HtmlEncode(Request.QueryString("EngineID")))
  Else
    MyCommon.QueryStr = "select EngineID from OfferIDs where OfferID=" & OfferID & ";"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
      EngineID = MyCommon.NZ(rst.Rows(0).Item("EngineID"), 0)
    End If
  End If
  
  AnyCustomerEnabled = (MyCommon.Fetch_CPE_SystemOption(125) = "1")
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  ' get all the any banner cardholder customer groups for this offer 
  'If (BannersEnabled) Then
  '  MyCommon.QueryStr = "select CustomerGroupID, Name from CustomerGroups with (NoLock) " & _
  '                      "where BannerID in (select BannerID from BannerOffers with (NoLock) where OfferID=" & OfferID & ") and deleted =0;"
  '  rstBannerCgs = MyCommon.LRT_Select
  'End If

  MyCommon.QueryStr = "select RewardOptionID from CPE_RewardOptions with (NoLock) where IncentiveID=" & OfferID & " and touchresponse=0 and deleted=0;"
  rst = MyCommon.LRT_Select
  
  If rst.Rows.Count > 0 Then
    roid = rst.Rows(0).Item("RewardOptionID")
  End If
  
  CMS.AMS.CurrentRequest.Resolver.AppName = "CPEoffer-con-customer.aspx"
  Dim m_offer1 As CMS.AMS.Contract.IOffer = CMS.AMS.CurrentRequest.Resolver.Resolve(Of CMS.AMS.Contract.IOffer)()
  Dim objOffer As CMS.AMS.Models.Offer = m_offer1.GetOffer(OfferID, CMS.AMS.Models.LoadOfferOptions.EligibilityCustomerCondition)
  If (objOffer IsNot Nothing) Then
    IsEligibilityConditionExistForOffer = IIf(objOffer.IsOptable, "TRUE", "FALSE")
    If objOffer.IsOptable Then
      For Each item In (From p In objOffer.EligibleCustomerGroupConditions.IncludeCondition
                                         Where p.Deleted = False
                                         Select p.CustomerGroup).ToList()
        If (EligibleIncludedcustomergroups <> "") Then
          EligibleIncludedcustomergroups = EligibleIncludedcustomergroups & ","
        End If
        EligibleIncludedcustomergroups = EligibleIncludedcustomergroups & item.CustomerGroupID
      Next
      For Each item In (From p In objOffer.EligibleCustomerGroupConditions.ExcludeCondition
                                                 Where p.Deleted = False
                                                 Select p.CustomerGroup).ToList()
        If (EligibleExcludedcustomergroups <> "") Then
          EligibleExcludedcustomergroups = EligibleExcludedcustomergroups & ","
        End If
        EligibleExcludedcustomergroups = EligibleExcludedcustomergroups & item.CustomerGroupID
      Next
    End If
  End If
  
  ' Check to see if someone is saving the condition
  If (Request.QueryString("save") <> "") Then
    ' Find out the roid
    If roid > 0 Then
      If (Request.QueryString("selGroups") = "") AndAlso (Request.QueryString("require_cg") = "") Then
        infoMessage = "Please select a customer group."
      Else
        ' Need to save the hhenable choice
        Dim form_Household As Integer = IIf(Request.QueryString("household") = "on", 1, 0)
        MyCommon.QueryStr = "update CPE_RewardOptions with (RowLock) set HHEnable=" & form_Household & " where TouchResponse=0 and IncentiveID=" & OfferID
        MyCommon.LRT_Execute()
        
        ' Check to see if a customer condition is required by the template, if applicable
        MyCommon.QueryStr = "select CustomerGroupID from CPE_IncentiveCustomerGroups with (NoLock) where RewardOptionID=" & roid & _
                            " and RequiredFromTemplate=1 and deleted=0 and ExcludedUsers=0;"
        rst = MyCommon.LRT_Select
        HasRequiredCG = (rst.Rows.Count > 0)
        
        ' We got some selected groups so let's blow out all the existing ones
        MyCommon.QueryStr = "update CPE_IncentiveCustomerGroups with (RowLock) set deleted=1, TCRMAStatusFlag=3 where RewardOptionID=" & roid & _
                            " and deleted=0 and ExcludedUsers=0"
        MyCommon.LRT_Execute()
        
        ' Let's handle the selected groups first
        If (Request.QueryString("selGroups") <> "") Then
          historyString = Copient.PhraseLib.Lookup("history.con-customer-edit", LanguageID) & ": " & Server.HtmlEncode(Request.QueryString("selGroups"))
          Ids = Server.HtmlEncode(Request.QueryString("selGroups")).Split(",")
          For i = 0 To Ids.Length - 1
            MyCommon.QueryStr = "insert into CPE_IncentiveCustomerGroups with (RowLock) (RewardOptionID,CustomerGroupID,ExcludedUsers,Deleted,LastUpdate,RequiredFromTemplate,TCRMAStatusFlag) " & _
                                "values(" & roid & "," & MyCommon.Extract_Val(Ids(i)) & ",0,0,getdate()," & IIf(HasRequiredCG, "1", "0") & ",3)"
            MyCommon.LRT_Execute()
          Next
        ElseIf HasRequiredCG Then
          MyCommon.QueryStr = "insert into CPE_IncentiveCustomerGroups with (RowLock) (RewardOptionID,ExcludedUsers,Deleted,LastUpdate,RequiredFromTemplate,TCRMAStatusFlag) " & _
                              "values(" & roid & ",0,0,getdate(),1,3)"
          MyCommon.LRT_Execute()
        End If
        
        ' Check to see if a customer condition is required by the template, if applicable
        MyCommon.QueryStr = "select CustomerGroupID from CPE_IncentiveCustomerGroups with (NoLock) where RewardOptionID=" & roid & _
                            " and RequiredFromTemplate=1 and deleted=0 and ExcludedUsers=1;"
        rst = MyCommon.LRT_Select
        HasRequiredCG = (rst.Rows.Count > 0)
        
        ' We got some selected groups so let's blow out all the existing ones
        MyCommon.QueryStr = "update CPE_IncentiveCustomerGroups with (RowLock) set deleted=1, TCRMAStatusFlag=3 where RewardOptionID=" & roid & _
                            " and deleted=0 and ExcludedUsers=1"
        MyCommon.LRT_Execute()
        
        ' Now let's handle the excluded groups
        If (Request.QueryString("exGroups") <> "") Then
          historyString = historyString & " " & Copient.PhraseLib.Lookup("term.excluding", LanguageID) & ": " & Server.HtmlEncode(Request.QueryString("exGroups"))
          Ids = Request.QueryString("exGroups").Split(",")
          For i = 0 To Ids.Length - 1
            MyCommon.QueryStr = "insert into CPE_IncentiveCustomerGroups with (RowLock) (RewardOptionID,CustomerGroupID,ExcludedUsers,Deleted,LastUpdate,RequiredFromTemplate, TCRMAStatusFlag) " & _
                                "values(" & roid & "," & MyCommon.Extract_Val(Ids(i)) & ",1,0,getdate()," & IIf(HasRequiredCG, "1", "0") & ",3)"
            'Send(MyCommon.QueryStr)
            MyCommon.LRT_Execute()
          Next
        ElseIf HasRequiredCG Then
          MyCommon.QueryStr = "insert into CPE_IncentiveCustomerGroups with (RowLock) (RewardOptionID,ExcludedUsers,Deleted,LastUpdate,RequiredFromTemplate, TCRMAStatusFlag) " & _
                              "values(" & roid & ",1,0,getdate(),1,3)"
          MyCommon.LRT_Execute()
        End If
        
        ' Now let's handle the excluded eligible groups
        If (objOffer IsNot Nothing AndAlso objOffer.IsOptable AndAlso Request.QueryString("mode") <> "optout") Then
          If (Request.QueryString("tomoveeligibleincludedgroups") <> "") Then
            MyCommon.QueryStr = "UPDATE [CustomerConditionDetails] SET [Excluded]=1 WHERE [ConditionID]= " & objOffer.EligibleCustomerGroupConditions.ConditionID & " AND [CustomerGroupID] IN (" & Server.HtmlEncode(Request.QueryString("tomoveeligibleincludedgroups")) & ")"
            MyCommon.LRT_Execute()
          End If
          If (Request.QueryString("toaddeligibleexcludedgroups") <> "") Then
            Ids = Server.HtmlEncode(Request.QueryString("toaddeligibleexcludedgroups")).Split(",")
            For i = 0 To Ids.Length - 1
              MyCommon.QueryStr = "INSERT INTO [CustomerConditionDetails] ([ConditionID],[CustomerGroupID],[Excluded]) VALUES(" & objOffer.EligibleCustomerGroupConditions.ConditionID & "," & MyCommon.Extract_Val(Ids(i)) & ",1)"
              'Send(MyCommon.QueryStr)
              MyCommon.LRT_Execute()
            Next
          End If
        End If
        
        ' Finally, set the "AllowOptOut" bit on the offer to 0 (False) if it's Any Customer or Any Cardholder with no exclusions
        If (Request.QueryString("exGroups") = "") Then
          If (Request.QueryString("selGroups") = "1") OrElse (Request.QueryString("selGroups") = "2") Then
            MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set AllowOptOut=0 where IncentiveID=" & OfferID
            MyCommon.LRT_Execute()
          End If
        End If
        
        MyCommon.QueryStr = "update CPE_Incentives with (RowLock) set LastUpdate=getdate(), LastUpdatedByAdminID=" & AdminUserID & ", StatusFlag=1 where IncentiveID=" & OfferID
        MyCommon.LRT_Execute()
        MyCommon.Activity_Log(3, OfferID, AdminUserID, historyString)
      End If
    Else
      infoMessage = Copient.PhraseLib.Lookup("ueoffer-con-customer.ErrorNoROID", LanguageID)
    End If
    
    If (infoMessage = "") Then
      CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
    Else
      CloseAfterSave = False
    End If
    '!
    'If CloseAfterSave = False Then
    '  Response.Redirect("CPEoffer-con-customer.aspx?OfferID=" & OfferID)
    'End If
    
  End If
  
  ' Dig the offer info out of the database
  ' No one clicked anything
  MyCommon.QueryStr = "Select IncentiveID,IsTemplate,ClientOfferID,IncentiveName as Name,CPE.Description,PromoClassID,CRMEngineID,Priority," & _
                      "StartDate,EndDate,EveryDOW,EligibilityStartDate,EligibilityEndDate,TestingStartDate,TestingEndDate,P1DistQtyLimit,P1DistTimeType,P1DistPeriod," & _
                      "P2DistQtyLimit,P2DistTimeType,P2DistPeriod,P3DistQtyLimit,P3DistTimeType,P3DistPeriod,EnableImpressRpt,EnableRedeemRpt,CreatedDate,CPE.LastUpdate,CPE.Deleted, CPEOARptDate, CPEOADeploySuccessDate, CPEOADeployRpt," & _
                      "CRMRestricted,StatusFlag,OC.Description as CategoryName,IsTemplate,FromTemplate from CPE_Incentives as CPE with (NoLock) left join OfferCategories as OC with (NoLock) on CPE.PromoClassID=OfferCategoryID where IncentiveID=" & OfferID & ";"
  rst = MyCommon.LRT_Select
  For Each row In rst.Rows
    Name = MyCommon.NZ(row.Item("Name"), "")
    IsTemplate = MyCommon.NZ(row.Item("IsTemplate"), False)
    FromTemplate = MyCommon.NZ(row.Item("FromTemplate"), False)
  Next
  
  ' Update the templates permission if necessary
  If (Request.QueryString("save") <> "" AndAlso Request.QueryString("IsTemplate") = "IsTemplate" AndAlso infoMessage = "") Then
    ' Update the status bits for the templates
    Dim form_Disallow_Edit As Integer = 0
    Dim form_Require_CG As Integer = 0
    
    If (Request.QueryString("Disallow_Edit") = "on") Then
      form_Disallow_Edit = 1
    End If
    
    If (Request.QueryString("require_cg") <> "") Then
      form_Require_CG = 1
    End If
    
    ' Both requiring and locking the customer group is not permitted 
    If (form_Disallow_Edit = 1 AndAlso form_Require_CG = 1) Then
      infoMessage = Copient.PhraseLib.Lookup("offer-con.lockeddenied", LanguageID)
      MyCommon.QueryStr = "update CPE_IncentiveCustomerGroups with (RowLock) set DisallowEdit=1, RequiredFromTemplate=0 " & _
                          " where RewardOptionID=" & roid & " and deleted = 0;"
    Else
      MyCommon.QueryStr = "update CPE_IncentiveCustomerGroups with (RowLock) set DisallowEdit=" & form_Disallow_Edit & _
                          ", RequiredFromTemplate=" & form_Require_CG & " " & _
                          " where RewardOptionID=" & roid & " and deleted = 0;"
    End If
    MyCommon.LRT_Execute()
    
    ' If necessary, create an empty customer condition
    If (form_Require_CG = 1) Then
      MyCommon.QueryStr = "select CustomerGroupID from CPE_IncentiveCustomerGroups with (NoLock) where RewardOptionID=" & roid & " and deleted = 0;"
      rst = MyCommon.LRT_Select
      If (rst.Rows.Count = 0) Then
        MyCommon.QueryStr = "insert into CPE_IncentiveCustomerGroups with (RowLock) (RewardOptionID,ExcludedUsers,Deleted,LastUpdate, RequiredFromTemplate, TCRMAStatusFlag) " & _
            " values(" & roid & ",0,0,getdate(),1,3)"
        MyCommon.LRT_Execute()
      End If
    End If
    
    If (infoMessage = "") Then
      CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
    Else
      CloseAfterSave = False
    End If
  End If
  
  If (IsTemplate OrElse FromTemplate) Then
    ' Dig the permissions if it's a template
    MyCommon.QueryStr = "select DisallowEdit, RequiredFromTemplate from CPE_IncentiveCustomerGroups with (NoLock) where RewardOptionID=" & roid & " and deleted = 0;"
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then
      Disallow_Edit = MyCommon.NZ(rst.Rows(0).Item("DisallowEdit"), False)
      RequireCG = MyCommon.NZ(rst.Rows(0).Item("RequiredFromTemplate"), False)
    Else
      Disallow_Edit = False
    End If
  End If
  
  MyCommon.QueryStr = "select HHEnable from CPE_RewardOptions with (NoLock) where IncentiveID=" & OfferID & " and touchresponse=0 and deleted=0;"
  rst = MyCommon.LRT_Select
  
  If rst.Rows.Count > 0 Then
    Household = MyCommon.NZ(rst.Rows(0).Item("HHEnable"), False)
  End If
  
  If Not isTemplate Then
    DisabledAttribute = IIf(Logix.UserRoles.EditOffer And Not (FromTemplate AndAlso Disallow_Edit), "", " disabled=""disabled""")
  Else
    DisabledAttribute = IIf(Logix.UserRoles.EditTemplates, "", " disabled=""disabled""")
  End If
  
  Send_HeadBegin("term.offer", "term.customercondition", OfferID)
  If (Request.QueryString("mode") = "optout") Then
    Send("<base target='_self'/>")
  End If
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
%>

<%
    Send("<script type=""text/javascript"">")

    Send("// This is the javascript array holding the function list")
    Send("// The PrintJavascriptArray ASP function can be used to print this array.")


    FullListSelect.Append("<select class=""longer"" id=""functionselect"" name=""functionselect"" multiple=""multiple"" size=""12"">")
    MyCommon.QueryStr = "select CustomerGroupID from CustomerGroups with (NoLock) where NewCardholders=1 and Deleted=0;"
    rst = MyCommon.LRT_Select

    If rst.Rows.Count > 0 Then
        NewCardholdersID = MyCommon.NZ(rst.rows(0).item("CustomerGroupID"), -1)
    Else
        NewCardholdersID = -1
    End If

    MyCommon.QueryStr = "select CustomerGroupID from CustomerGroups with (NoLock) where AnyCAMCardholder=1 and Deleted=0;"
    rst = MyCommon.LRT_Select

    If rst.Rows.Count > 0 Then
        AllCAM = MyCommon.NZ(rst.rows(0).item("CustomerGroupID"), -1)
    Else
        AllCAM = -1
    End If

    'Populate the Javascript array that holds the list of selectable customer groups
    If (EngineID <> 6) Then
        'MyCommon.QueryStr = "Select CustomerGroupID,Name from CustomerGroups with (NoLock) where Deleted=0 and AnyCustomer<>1 and CustomerGroupID<>2 and CustomerGroupID is not null and BannerID is null and NewCardholders=0 and CAMCustomerGroup<>1"
        MyCommon.QueryStr = "SELECT DISTINCT CG.CustomerGroupID, CG.Name " &
               "FROM CustomerGroups CG With (NOLOCK) " &
               "LEFT JOIN ExtSegmentMap ER on CG.CustomerGroupID = ER.InternalId " &
               "WHERE CG.Deleted = 0 And AnyCustomer <> 1 And CustomerGroupID <> 2 " &
               "AND CustomerGroupID IS NOT NULL AND BannerID IS NULL " &
               "And NewCardholders = 0 And CAMCustomerGroup <> 1 AND CG.Deleted = 0 " &
               "AND ((ER.SegmentTypeID IS NULL OR ER.SegmentTypeID = 1 OR ER.SegmentTypeID = 3) AND " &
               "(ER.IncentiveId = " & OfferID & " OR ER.ExtSegmentID > 0 OR ER.ExtSegmentID IS NULL)) " &
               " And CG.isOptInGroup = 0 "
        If Request.QueryString("mode") = "optout" Then
            MyCommon.QueryStr = MyCommon.QueryStr & " and CG.IsOptInGroup=0 ORDER BY CG.Name;"
        Else
            MyCommon.QueryStr = MyCommon.QueryStr & " ORDER BY CG.Name;"
        End If
    Else
        MyCommon.QueryStr = "Select CustomerGroupID,Name from CustomerGroups with (NoLock) where Deleted=0 and AnyCustomer<>1 and CustomerGroupID<>2 and CustomerGroupID is not null and BannerID is null and NewCardholders=0 and CAMCustomerGroup=1 and AnyCAMCardholder<>1 order by Name;"
    End If
    rst = MyCommon.LRT_Select
    If (rst.Rows.Count > 0) Then

        Sendb("var exceptlist = new Array(")
        If (EngineID <> 6) Then
            MyCommon.QueryStr = "select CustomerGroupID from CustomerGroups with (NoLock) where (AnyCardholder=1 or AnyCustomer=1) and Deleted=0 order by CustomerGroupID;"
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count > 0) Then
                i = 1
                For Each row In rst.Rows
                    If (i > 1) Then Sendb(",")
                    Sendb(MyCommon.NZ(row.Item("CustomerGroupID"), 0))
                    i += 1
                Next
            End If
            Send(");")
        Else
            MyCommon.QueryStr = "select CustomerGroupID from CustomerGroups with (NoLock) where AnyCAMCardholder=1 and Deleted=0 order by CustomerGroupID;"
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count > 0) Then
                Sendb("-1,")
                For Each row In rst.Rows
                    Sendb(MyCommon.NZ(row.Item("CustomerGroupID"), 0))
                Next
            Else
                Sendb("-1,-2")
            End If
            Send(");")
        End If
    Else

        Sendb("var exceptlist = new Array(")
        If (EngineID <> 6) Then
            MyCommon.QueryStr = "select CustomerGroupID from CustomerGroups with (NoLock) where (AnyCardholder=1 or AnyCustomer=1) and Deleted=0 order by CustomerGroupID;"
            rst = MyCommon.LRT_Select
            If (rst.Rows.Count > 0) Then
                i = 1
                For Each row In rst.Rows
                    If (i > 1) Then Sendb(",")
                    Sendb(MyCommon.NZ(row.Item("CustomerGroupID"), 0))
                    i += 1
                Next
                Send(");")
            Else
                Send("""" & "-99" & """);")
            End If
        Else
            MyCommon.QueryStr = "select CustomerGroupID from CustomerGroups with (NoLock) where AnyCAMCardholder=1 and Deleted=0 order by CustomerGroupID;"
            If (rst.Rows.Count > 0) Then
                rst = MyCommon.LRT_Select
                For Each row In rst.Rows
                    Sendb(MyCommon.NZ(row.Item("CustomerGroupID"), 0))
                Next
                Send(");")
            Else
                Send("""" & "-99" & """);")
            End If
        End If
    End If

%>
var isFireFox = (navigator.appName.indexOf('Mozilla') !=-1) ? true : false;
var timer;
function xmlPostTimer(strURL,mode)
{
  clearTimeout(timer);
  timer=setTimeout("xmlhttpPost('" + strURL + "','" + mode + "')", 250);
}

function xmlhttpPost(strURL,mode) {
  var xmlHttpReq = false;
  var self = this;
  
  document.getElementById("searchLoadDiv").innerHTML = '<% Sendb(Copient.PhraseLib.Lookup("message.loading", LanguageID)) %>';
 
  // Mozilla/Safari
  if (window.XMLHttpRequest) {
    self.xmlHttpReq = new XMLHttpRequest();
  }
  // IE
  else if (window.ActiveXObject) {
    self.xmlHttpReq = new ActiveXObject("Microsoft.XMLHTTP");
  }
  var qryStr = getgroupquery(mode);
  self.xmlHttpReq.open('POST', strURL + "?" + qryStr, true);
  self.xmlHttpReq.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
  self.xmlHttpReq.onreadystatechange = function() {
    if (self.xmlHttpReq !=null && self.xmlHttpReq.readyState == 4) {
      updatepage(self.xmlHttpReq.responseText);
    }
  }

  self.xmlHttpReq.send(qryStr);
  //self.xmlHttpReq.send(getquerystring());
}

function getgroupquery(mode) {
  var radioString;
  if(document.getElementById('functionradio2').checked) {
    radioString = 'functionradio2';
  }
  else {
    radioString = 'functionradio1';
  }
  var select = document.getElementById('selected').options;
  var exclude = document.getElementById('excluded').options;
  
  var opt = 0;
  var selectedGroups = '';
  for(opt = 0; opt < select.length; opt++){
    if(select[opt] != null){
      if(selectedGroups != '') {
        selectedGroups += ",";
      }
      selectedGroups += select[opt].value;
    }
  }
  
  var excludedGroups = '';
  for(opt = 0; opt < exclude.length; opt++){
    if(exclude[opt] != null){
      if(excludedGroups != '') {
        excludedGroups += ",";
      }
      excludedGroups += exclude[opt].value;
    }
  }
  return "Mode=" + mode + "&Search=" + document.getElementById('functioninput').value + "&EngineID=" + '<% Sendb(EngineID)%>' + "&AnyCustomerEnabled=" + '<% Sendb(AnyCustomerEnabled)%>' + 
  "&SelectedGroups=" + selectedGroups + "&ExcludedGroups=" + excludedGroups + "&OfferID=" + '<% Sendb(OfferID)%>' + "&SearchRadio=" + radioString;
 
}

function updatepage(str) {
  if(str.length > 0){
    if(!isFireFox){
      document.getElementById("cgList").innerHTML = '<select class="longer" id="functionselect" name="functionselect" size="12" multiple="multiple">' + str + '</select>';
    }
    else{
      document.getElementById("functionselect").innerHTML = str;
    }
    document.getElementById("searchLoadDiv").innerHTML = '&nbsp;';
    if (document.getElementById("functionselect").options.length > 0) {
      document.getElementById("functionselect").options[0].selected = true;
    }
  }
  else if(str.length == 0){
    if(!isFireFox){
       document.getElementById("cgList").innerHTML = '';
    }
    else{
      document.getElementById("functionselect").innerHTML = '<select class="longer" id="functionselect" name="functionselect" size="12" multiple="multiple"></select>';
    }
    document.getElementById("searchLoadDiv").innerHTML = '&nbsp;';
  }
}

// This is the function that refreshes the list after a keypress.
// The maximum number to show can be limited to improve performance with
// huge lists (1000s of entries).
// The function clears the list, and then does a linear search through the
// globally defined array and adds the matches back to the list.
function handleKeyUp(maxNumToShow) {
  var selectObj, textObj, functionListLength;
  var i,  numShown;
  var searchPattern;
  var optVal;

  document.getElementById("functionselect").size = "12";
  
  // Set references to the form elements
  selectObj = document.forms[0].functionselect;
  textObj = document.forms[0].functioninput;
  
  // Remember the function list length for loop speedup
  functionListLength = functionlist.length;
  
  // Set the search pattern depending
  if(document.forms[0].functionradio[0].checked == true) {
    searchPattern = "^"+textObj.value;
  } else {
    searchPattern = textObj.value;
  }
  searchPattern = cleanRegExpString(searchPattern);
  
  // Create a regular expression
  re = new RegExp(searchPattern,"gi");
  
  // Loop through the array and re-add matching options
  numShown = 0;
  
  if (textObj.value == '') {
    document.getElementById("cgList").innerHTML = fullList;
  } else {
    var newList = '<select class="longer" id="functionselect" name="functionselect" size="12" multiple="multiple">';
    for(i = 0; i < functionListLength; i++) {
      if(functionlist[i].search(re) != -1) {
        optVal = vallist[i<% Sendb(IIf(EngineID=6, "-1", "")) %>];
        if (optVal != "") {
          if (optVal == 1 || optVal == 2 || optVal == <%Sendb(NewCardholdersID)%>) {
            selectObj[numShown].style.fontWeight = 'bold';
            selectObj[numShown].style.color = 'brown';
            newList += '<option value="' + optVal + '" style="color:brown;font-weight:bold;"> ' + functionlist[i] + '<\/option>';
          } else {
            newList += '<option value="' + optVal + '"> ' + functionlist[i] + '<\/option>';
          }
          numShown++;
        }
      }
      // Stop when the number to show is reached
      if(numShown == maxNumToShow) {
        break;
      }
    }
    newList += '<\/select>'
    document.getElementById("cgList").innerHTML = newList;
  }
  
  removeUsed(true);
  
  // When options list whittled to one, select that entry
  if(selectObj.length == 1) {
    selectObj.options[0].selected = true;
  }
}

function removeUsed(bSkipKeyUp) {
  //if (!bSkipKeyUp) handleKeyUp(99999);
  if(!bSkipKeyUp) xmlhttpPost('OfferFeeds.aspx','ConditionCustomerGroups');
  // this function will remove items from the functionselect box that are used in 
  // selected and excluded boxes
  
  var funcSel = document.getElementById('functionselect');
  var exSel = document.getElementById('selected');
  var elSel = document.getElementById('excluded');
  var i,j;
  
  for (i = elSel.length - 1; i>=0; i--) {
    for(j=funcSel.length-1;j>=0; j--) {
      if(funcSel.options[j].value == elSel.options[i].value){
        funcSel.options[j] = null;
      }
    }
  }
  
  for (i = exSel.length - 1; i>=0; i--) {
    for(j=funcSel.length-1;j>=0; j--) {
      if(funcSel.options[j].value == exSel.options[i].value){
        funcSel.options[j] = null;
      }
    }
  }
}
function IsSelectedForGrantMembership(selectedValue){

    <%
        Dim dt As DataTable
        Dim strID As String = ""
        MyCommon.QueryStr = "select CGT.CustomerGroupID from CPE_DeliverableCustomerGroupTiers CGT " &
               " inner Join CPE_Deliverables CD On CGT.DeliverableID = CD.DeliverableID " &
                "inner Join CPE_RewardOptions CRO on CRO.RewardOptionID = CD.RewardOptionID " &
                "inner Join CPE_Incentives CI On CI.IncentiveID = CRO.IncentiveID " &
                "where CD.Deleted = 0 And CI.EngineID = 2 And CI.IncentiveID =" & OfferID
        dt = MyCommon.LRT_Select()
        If dt IsNot Nothing And dt.Rows.Count > 0 Then
            For Each row In dt.Rows
                strID &= ", " & MyCommon.NZ(row.Item("CustomerGroupID"), 0)
            Next
        End If

     %>
    var idsstr = '<%= strID %>'
    var ids = idsstr.split(',');
    for (var i = 0; i < ids.length; i++) {
        if (ids[i].trim() == selectedValue) {
            return true;
        }
    }
    return false;
}
// this function gets the selected value and loads the appropriate
// php reference page in the display frame
// it can be modified to perform whatever action is needed, or nothing
function handleSelectClick(itemSelected) {
 <%
   CMS.AMS.CurrentRequest.Resolver.AppName = "CPEoffer-con-customer.aspx"
   Dim m_offer As CMS.AMS.Contract.IOffer = CMS.AMS.CurrentRequest.Resolver.Resolve(Of CMS.AMS.Contract.IOffer)()
   Dim m_defaultcust As CMS.AMS.Models.CustomerGroup = m_offer.GetOfferDefaultCustomerGroup(OfferID, EngineID)
   Dim CustGroupId As String = ""
   If (m_defaultcust IsNot Nothing) Then
     CustGroupId = m_defaultcust.CustomerGroupID
   End If
      
   %>
 var IsEligibleConditionExist = '<%= m_offer.IsOfferOptable(OfferID)%>';
 var defaultcusgroupid= '<%=CustGroupId%>'
 var IsOptout = '<%= Request.QueryString("mode")%>'
  var funcSel = document.getElementById('functionselect');
  var exSel = document.getElementById('selected');
  var elSel = document.getElementById('excluded');
  var i,j;
  
  textObj = document.forms[0].functioninput;
  
  selectObj = document.forms[0].functionselect;
  selectedValue = document.getElementById("functionselect").value;
  if(selectedValue != ""){ selectedText = selectObj[document.getElementById("functionselect").selectedIndex].text; }
  
  selectboxObj = document.forms[0].selected;
  selectedboxValue = document.getElementById("selected").value;
  if(selectedboxValue != ""){ selectedboxText = selectboxObj[document.getElementById("selected").selectedIndex].text; }
  
  excludedbox = document.forms[0].excluded;
  excludedboxValue = document.getElementById("excluded").value;
  if(excludedboxValue != ""){ excludeboxText = excludedbox[document.getElementById("excluded").selectedIndex].text; }
  
  if (itemSelected == "select1") {
    if (selectedValue != "") {
     if(IsSelectedForGrantMembership(selectedValue))
     {
        alert('<% Sendb(Copient.PhraseLib.Lookup("error.includedcustomer", LanguageID))%>');
        return;
     }
     if(IsEligibleConditionExist == "True" && IsOptout != 'optout')
    {       
      if(selectedValue == 1 || selectedValue == 2)
      {
          alert('<% Sendb(Copient.PhraseLib.Lookup("term.invalidgroupforoptin", LanguageID))%>');
          return;
      }
    }
      // add items to selected box
      document.getElementById('deselect1').disabled=false;
      document.getElementById('select2').disabled=false;
      document.getElementById('save').disabled=false;
      
      while (selectObj.selectedIndex != -1) {
        selectedText = selectObj.options[selectObj.selectedIndex].text;
        selectedValue = selectObj.options[selectObj.selectedIndex].value;
        if(selectedValue==1 || selectedValue == 2 || selectedValue == <%Sendb(AllCAM) %>){
          document.getElementById('select1').disabled=true;  
          document.getElementById('select2').disabled=false;
          // someone's adding all customers we need to empty the select box
          for (i = selectboxObj.length - 1; i>=0; i--){
            selectboxObj.options[i] = null;
          }
          selectboxObj[selectboxObj.length] = new Option(selectedText,selectedValue);
          selectboxObj[selectboxObj.length-1].style.color = 'brown';
          selectboxObj[selectboxObj.length-1].style.fontWeight = 'bold';
          selectObj[selectObj.selectedIndex].selected = false;

          // any customer currently does not support exclusions
          if (selectedValue==1) {
            for (i = excludedbox.length - 1; i>=0; i--){
              excludedbox.options[i] = null;
            }
            document.getElementById('select2').disabled=true;
          }

          break;
        } else {
          selectboxObj[selectboxObj.length] = new Option(selectedText,selectedValue);
          if (selectedValue== "<%Sendb(NewCardholdersID)%>") {
            selectboxObj[selectboxObj.length-1].style.color = 'brown';
            selectboxObj[selectboxObj.length-1].style.fontWeight = 'bold';
          }
          selectObj[selectObj.selectedIndex].selected = false;
        }
      }
    }
  }
  
  if (itemSelected == "deselect1") {
    if (selectedboxValue != "") {

    //If all items are selected for remove and exclude item exist
        if ($("#selected").children().length == $("#selected :selected").length && $("#excluded").children().length>0)
        {
          alert('<% Sendb(Copient.PhraseLib.Lookup("term-ValidationOnDeleteForAllSelected", LanguageID))%>');
          return;
        }
      
        //if Eligibility condition exist then verify that default group should not be in selected list
        if (IsEligibleConditionExist == "True" && IsOptout != 'optout')
        {
          var defaultfound=0;
         $("#selected :selected").each
         (
            function()
            {
              if (this.value == defaultcusgroupid)
              {
                defaultfound=1;
                return;
              }
            }
          );
        
          if(defaultfound == 1)
          {
              alert('<% Sendb(Copient.PhraseLib.Lookup("term.deselectdefaultgroup", LanguageID))%>');
              return;
          }
        }
       
    
      // remove items from selected box
      while (document.getElementById("selected").selectedIndex != -1) {
          document.getElementById("selected").remove(document.getElementById("selected").selectedIndex);
      }
      
      if (selectedboxValue == 2 || selectedboxValue == <%Sendb(AllCAM) %>) {
        document.getElementById('select1').disabled=true;
        document.getElementById('select2').disabled=false;
        document.getElementById('deselect2').disabled= true;
      }
      
      if (selectboxObj.length == 0) {
        // nothing in the select box so disable deselect
        document.getElementById('deselect1').disabled=true;
        document.getElementById('select2').disabled=true;
        document.getElementById('deselect2').disabled=true;
        // this being the case, let's also disable the page's save button
        // (but not if it's a template with "required" checked)
        if (document.getElementById("require_cg") != null) {
          if (document.getElementById("require_cg").checked == false) {
            document.getElementById('save').disabled=true;
          }
        } else {
          document.getElementById('save').disabled=true;
        }
      }
      
      if (selectedboxValue != "") {
        if (excludedbox.length == 0) {
          document.getElementById('select1').disabled=false;
        }
        document.getElementById('select2').disabled=false;
      }
      
      if (selectboxObj.length == 0 || selectedboxValue == "") {
        document.getElementById('select1').disabled=false; 
        document.getElementById('select2').disabled=true;
      } else if (selectedboxValue != 0 || selectboxObj.length != 0) {
        document.getElementById('select2').disabled=false;
      } else {
        document.getElementById('select1').disabled=false; 
        document.getElementById('select2').disabled=false;
      }
    }
  }
  
  if (itemSelected == "select2") {
    if (selectedValue != "") {
      var AnyCardholder = "<% Sendb(Copient.PhraseLib.Lookup("term.anycardholder", LanguageID))%>";
      if (selectedValue != '1' && selectedValue !='2' && selectedValue != <%Sendb(NewCardholdersID)%>) {
        // add items to excluded box
        while(document.getElementById("functionselect").selectedIndex != -1){
          selectedText = selectObj.options[selectObj.selectedIndex].text;
          selectedValue = selectObj.options[selectObj.selectedIndex].value;
          excludedbox[excludedbox.length] = new Option(selectedText,selectedValue);
          selectObj[selectObj.selectedIndex].selected = false;
        }
        if (excludedbox.length == 1) {
          document.getElementById('select2').disabled=false;
          // need to disable deselection on the selected box also since we added an excluded group
          document.getElementById('deselect1').disabled=false;
          document.getElementById('deselect2').disabled=false;
        }
      } else if (selectedValue=='1') {
        alert("<% Sendb(Copient.PhraseLib.Lookup("offer-con.anycustomerexcluded", LanguageID))%>");
      } else if (selectedValue=='2') {
        alert("<% Sendb(Copient.PhraseLib.Lookup("offer-con.anycardholderexcluded", LanguageID))%>");
      } else {
        alert("<% Sendb(Copient.PhraseLib.Lookup("offer-con.newcardholderexcluded", LanguageID))%>");
      }
    } 
  }
  
  if (itemSelected == "deselect2") {
    if (excludedboxValue != "") {
      // remove items from excluded box
      while (document.getElementById("excluded").selectedIndex != -1) {
        document.getElementById("excluded").remove(document.getElementById("excluded").selectedIndex);
      }
      if (excludedbox.length == 0) {
        document.getElementById('select2').disabled=false;
        document.getElementById('deselect1').disabled=false;
        document.getElementById('deselect2').disabled=true;
      }
    }
  }
  
  updateExceptionButtons();
  
  // remove items from large list that are in the other lists
  removeUsed(false);
  return true;
}

function saveForm() {
  var funcSel = document.getElementById('functionselect');
  var exSel = document.getElementById('excluded');
  var elSel = document.getElementById('selected');
  var objinfobar = document.getElementById('infobar');
  var isOfferOptable="FALSE";
  var eligibleIncluded="";
  var eligibleExcluded="";
  var notFoundInEligibleExcludedList="";
  var needstoremovefromeligibleincludedlist="";
  var needstoremovefromeligibleincludedlistText="";
  var i,j;
  var selectList = "";
  var excludededList = "";
  var htmlContents = "";
  var isFoundInIncludedEligibleCondition="";
  var pagemode = '<%= Request.QueryString("mode")%>'

  objinfobar.innerHTML= '';
  objinfobar.style.display='none';
  isOfferOptable = document.getElementById('HdnIsOfferOptable').value;
  eligibleIncluded = document.getElementById('Hdnincludedeligiblegroup').value;
  eligibleExcluded = document.getElementById('Hdnexcludedeligiblegroup').value;
  
  // assemble the list of values from the selected box
  for (i = elSel.length - 1; i>=0; i--) {
    if(elSel.options[i].value != ""){
      if(selectList != "") { selectList = selectList + ","; }
      selectList = selectList + elSel.options[i].value;
    }
  }
  for (i = exSel.length - 1; i>=0; i--) {
    if(exSel.options[i].value != ""){
      
      if(excludededList != "") { excludededList = excludededList + ","; }
      excludededList = excludededList + exSel.options[i].value;
      
      //validation for eligiblegroups if offer is optable
      if(isOfferOptable=="TRUE" && pagemode != 'optout')
      {
        isFoundInIncludedEligibleCondition="";
        if(eligibleIncluded.indexOf(exSel.options[i].value) != -1)
        {
          isFoundInIncludedEligibleCondition="TRUE";

          if(needstoremovefromeligibleincludedlist != "")
          { needstoremovefromeligibleincludedlist = needstoremovefromeligibleincludedlist + ","; }
          
          if(needstoremovefromeligibleincludedlistText !="")
          {needstoremovefromeligibleincludedlistText = needstoremovefromeligibleincludedlistText + ",";}

          needstoremovefromeligibleincludedlist = needstoremovefromeligibleincludedlist + exSel.options[i].value;
          needstoremovefromeligibleincludedlistText = needstoremovefromeligibleincludedlistText + exSel.options[i].text;
        }

        if(isFoundInIncludedEligibleCondition == "" && eligibleExcluded.indexOf(exSel.options[i].value) == -1 )
        {
              if(notFoundInEligibleExcludedList != "")
              { notFoundInEligibleExcludedList = notFoundInEligibleExcludedList + ","; }
              notFoundInEligibleExcludedList = notFoundInEligibleExcludedList + exSel.options[i].value;
        }
      }
    }
  }

  if(isOfferOptable=="TRUE" && pagemode != 'optout')
  {
    if(needstoremovefromeligibleincludedlist!="")
    {
       var confmsg="";
        
       if(needstoremovefromeligibleincludedlist.split(",").length >= eligibleIncluded.split(",").length)
       {
         confmsg= "<% Sendb(Copient.PhraseLib.Lookup("term.msgsavediscarded", LanguageID))%>";
         objinfobar.innerHTML=confmsg.format([needstoremovefromeligibleincludedlistText]);
         objinfobar.style.display='block';
         return false;
       }
       else
       {
         confmsg = "<% Sendb(Copient.PhraseLib.Lookup("term.confirmtodeleligiblegroup", LanguageID))%>";
         confmsg= confmsg.format([needstoremovefromeligibleincludedlistText])
         if(!confirm(confmsg))
         {
         return false;
         }
       }
    }
  }

  // time to build up the hidden variables to pass for saving
  htmlContents = "<input type=\"hidden\" name=\"selGroups\" value=" + selectList + "> ";
  htmlContents = htmlContents + "<input type=\"hidden\" name=\"exGroups\" value=" + excludededList + ">";
  htmlContents = htmlContents + "<input type=\"hidden\" name=\"tomoveeligibleincludedgroups\" value=" + needstoremovefromeligibleincludedlist + ">";
  htmlContents = htmlContents + "<input type=\"hidden\" name=\"toaddeligibleexcludedgroups\" value=" + notFoundInEligibleExcludedList + ">";
  document.getElementById("hiddenVals").innerHTML = htmlContents;
  
 // alert(htmlContents);
  return true;
}

function updateButtons(){
  

  if(document.getElementById('selected').length > 0){
    var selectedValue = document.getElementById('selected').options[0].value;
    if(selectedValue==1 || selectedValue == 2 || selectedValue == <%Sendb(AllCAM) %>) {
      // all customers is in the selected box so disable adding another
      document.getElementById('select1').disabled=true; 
      if(document.forms[0].selected.length == 1) {
      //since one or more cg can be excluded, do not disable selection
        document.getElementById('select2').disabled=false; 
        if(document.forms[0].excluded.length == 0) { 
          // nothing is excluded so allow excluding one
          document.getElementById('select2').disabled=false; 
          document.getElementById('deselect2').disabled=true; 
          document.getElementById('deselect1').disabled=false; 
        } else {
          document.getElementById('select2').disabled=false; 
          document.getElementById('deselect2').disabled=false; 
          document.getElementById('deselect1').disabled=true; 
        }
      }
    } else {
      // something is selected but its not all customers
      document.getElementById('select1').disabled=false; 
      document.getElementById('deselect1').disabled=false;
      document.getElementById('select2').disabled=false; 
      document.getElementById('deselect2').disabled=true;
    }
  } else {
    // nothing is selected 
    document.getElementById('select1').disabled=false; 
    document.getElementById('deselect1').disabled=false;
    document.getElementById('select2').disabled=true; 
    document.getElementById('deselect2').disabled=true;
  }
  
  <%
   If Not isTemplate Then   
      If Not (Logix.UserRoles.EditOffer And Not (FromTemplate AndAlso Disallow_Edit)) Then
        Send("  disableAll();")
      End If
   Else
     If Not (Logix.UserRoles.EditTemplates) Then
      Send("  disableAll();")
     End If
   End If        
  %>
}

function updateExceptionButtons() {
  var exSel = document.getElementById('excluded');
  var elSel = document.getElementById('selected');
  var bEligible = false;
  
  // check if there already is an excluded group, if so disable select and enable deselect
  if (exSel != null && elSel != null && exSel.options.length == 0) {
    // check if a exception-qualify customer group is in the selected list
    for (var i=0; i < elSel.options.length && !bEligible; i++) {
      bEligible = isExceptionGroup(elSel.options[i].value)
    }
  } else if (exSel != null && exSel.options.length > 0) {
    document.getElementById('select2').disabled=false; 
    document.getElementById('deselect2').disabled=false; 
  } else {
    document.getElementById('select2').disabled=true; 
    document.getElementById('deselect2').disabled=true; 
  }

  // exclusions are not currently supported for Any Customer selection
  if (elSel != null && elSel.options.length > 0 && elSel.options[0].value=='1') {
    document.getElementById('select2').disabled=true; 
    document.getElementById('deselect2').disabled=true; 
  }

}

function isExceptionGroup(groupID) {
  var bRetVal = false;
  
  for (var i=0; i < exceptlist.length && !bRetVal; i++) {
    bRetVal = (exceptlist[i] == groupID) 
  }
  
  return bRetVal;
}

function handleRequiredToggle() {
  if(document.forms[0].selected.length == 0) {
    if (document.getElementById("require_cg").checked == false) {
      document.getElementById('save').disabled=true;
    } else {
      document.getElementById('save').disabled=false;
    }
  }
  if (document.getElementById("require_cg").checked == true) {
    document.getElementById("Disallow_Edit").checked=false;
  }
}

function disableAll() {
  document.getElementById('select1').disabled=true;
  document.getElementById('select2').disabled=true;
  document.getElementById('deselect1').disabled=true;
  document.getElementById('deselect2').disabled=true;
  document.getElementById('functionselect').disabled=true;
  document.getElementById('selected').disabled=true;
  document.getElementById('excluded').disabled=true;
}
</script>
<%
  Send("<script type=""text/javascript"">")
  Send("function ChangeParentDocument() { ")
  If (EngineID = 3) Then
    Send("  opener.location = 'web-offer-con.aspx?OfferID=" & OfferID & "'; ")
  ElseIf (EngineID = 5) Then
    Send("  opener.location = 'email-offer-con.aspx?OfferID=" & OfferID & "'; ")
  ElseIf (EngineID = 6) Then
    Send("  opener.location = 'CAM/CAM-offer-con.aspx?OfferID=" & OfferID & "'; ")
  Else
    Send("  opener.location = 'CPEoffer-con.aspx?OfferID=" & OfferID & "'; ")
  End If
  Send("} ")
  Send("</script>")
  Send_HeadEnd()
  
  If Request.QueryString("mode") = "optout" Then
    If (IsTemplate) Then
      Send_BodyBegin(13)
    Else
      Send_BodyBegin(3)
    End If
  Else
    If (IsTemplate) Then
      Send_BodyBegin(12)
    Else
      Send_BodyBegin(2)
    End If
  End If
  
  If (Logix.UserRoles.AccessOffers = False AndAlso Not IsTemplate) Then
    Send_Denied(2, "perm.offers-access")
    GoTo done
  End If
  If (Logix.UserRoles.AccessTemplates = False AndAlso IsTemplate) Then
    Send_Denied(2, "perm.offers-access-templates")
    GoTo done
  End If
  If (BannersEnabled AndAlso Not Logix.IsAccessibleOffer(AdminUserID, OfferID)) Then
    Send("<script type=""text/javascript"" language=""javascript"">")
    Send("  function ChangeParentDocument() { return true; } ")
    Send("</script>")
    Send_Denied(1, "banners.access-denied-offer")
    Send_BodyEnd()
    GoTo done
  End If
%>
<form action="#" name="mainform" id="mainform" onsubmit="return saveForm();">
  <span id="hiddenVals"></span>
  <input type="hidden" id="OfferID" name="OfferID" value="<% Sendb(OfferID) %>" />
  <input type="hidden" id="Name" name="Name" value="<% Sendb(Name) %>" />
  <input type="hidden" id="ConditionID" name="ConditionID" value="<% Sendb(ConditionID) %>" />
  <input type="hidden" id="EngineID" name="EngineID" value="<% Sendb(EngineID) %>" />
   <input type="hidden" id="mode" name="mode" value="<%=Request.QueryString("mode") %>" />
   <input type="hidden" id="Hdnincludedeligiblegroup" value="<% Sendb(EligibleIncludedcustomergroups ) %>" />
   <input type="hidden" id="Hdnexcludedeligiblegroup" value="<% Sendb(EligibleExcludedcustomergroups ) %>" />
   <input type="hidden" id="HdnIsOfferOptable" value="<% Sendb(IsEligibilityConditionExistForOffer ) %>" />
  <input type="hidden" id="IsTemplate" name="IsTemplate" value="<% 
        If (IsTemplate) Then
        Sendb("IsTemplate")
        Else
        Sendb("Not")
        End If
        %>" />
  <div id="intro">
    <%If (IsTemplate) Then
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.template", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.customercondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
      Else
        Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.offer", LanguageID) & " #" & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.customercondition", LanguageID), VbStrConv.Lowercase) & "</h1>")
      End If
    %>
    <div id="controls">
      <% If (IsTemplate) Then%>
      <span class="temp">
        <input type="checkbox" class="tempcheck" id="Disallow_Edit" name="Disallow_Edit"<% if(disallow_edit)then sendb(" checked=""checked""") %> />
        <label for="Disallow_Edit"><% Sendb(Copient.PhraseLib.Lookup("term.locked", LanguageID))%></label>
      </span>
      <% End If%>
      <%
      If Not IsTemplate Then
        If (Logix.UserRoles.EditOffer And Not (FromTemplate And Disallow_Edit)) Then
          Send_Save()
        End If
      Else
        If (Logix.UserRoles.EditTemplates) Then
          Send_Save()
        End If
      End If
      %>
    </div>
  </div>
  <div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
     <% If (infoMessage = "") Then Send("<div id=""infobar"" class=""red-background"" style='display:none'></div>")%>
    <div id="column1">
      <div class="box" id="selector">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.customergroup", LanguageID))%>
          </span>
          <% If (IsTemplate) Then%>
          <span class="tempRequire">
            <input type="checkbox" class="tempcheck" id="require_cg" name="require_cg" onclick="handleRequiredToggle();"<% if(requirecg)then sendb(" checked=""checked""") %> />
            <label for="require_cg"><% Sendb(Copient.PhraseLib.Lookup("term.required", LanguageID))%></label>
          </span>
          <% ElseIf (FromTemplate And RequireCG) Then%>
          <span class="tempRequire">*
            <%Sendb(Copient.PhraseLib.Lookup("term.required", LanguageID))%>
          </span>
          <% End If%>
        </h2>
        <input type="radio" id="functionradio1" name="functionradio" checked="checked"<% sendb(disabledattribute) %> /><label for="functionradio1"><% Sendb(Copient.PhraseLib.Lookup("term.startingwith", LanguageID))%></label>
        <input type="radio" id="functionradio2" name="functionradio"<% sendb(disabledattribute) %> /><label for="functionradio2"><% Sendb(Copient.PhraseLib.Lookup("term.containing", LanguageID))%></label><br />
       <%-- <input type="text" class="medium" id="functioninput" name="functioninput" maxlength="100" onkeyup="handleKeyUp(99999);" value=""<% sendb(disabledattribute) %> /><br />--%>
        <input type="text" class="medium" id="functioninput" name="functioninput" maxlength="100" onkeyup="javascript:xmlPostTimer('OfferFeeds.aspx','ConditionCustomerGroups');" value=""<% sendb(disabledattribute) %> /><br />
        <div id="searchLoadDiv" style="display:block;" >&nbsp;</div>
        <div id="cgList">
          <select class="longer" id="functionselect" name="functionselect" size="12"<% sendb(disabledattribute) %> >
            <%
                If (EngineID <> 6) Then
                    'add "special" customer groups
                    If AnyCustomerEnabled Then
                        'see if the offer conditions/rewards allow us to display AnyCustomer group.  (This is not allowed if conditions/rewards require a known customer ex: Points, Stored Value, etc.)
                        MyCommon.QueryStr = "dbo.pa_Check_AnyCustomer_Violation"
                        MyCommon.Open_LRTsp()
                        MyCommon.LRTsp.Parameters.Add("@IncentiveID", SqlDbType.BigInt).Value = OfferID
                        Dim anyCustomerDT As DataTable = MyCommon.LRTsp_select
                        MyCommon.Close_LRTsp()
                        If anyCustomerDT.Rows.Count = 0 Then
                            Send("<option value=""1"" style=""color:brown;font-weight:bold;"">" & StrConv(Copient.PhraseLib.Lookup("term.anycustomer", LanguageID), VbStrConv.ProperCase) & "</option>")
                        End If
                        dst = Nothing
                    End If

                    Send("<option value=""2"" style=""color:brown;font-weight:bold;"">" & Copient.PhraseLib.Lookup("term.anycardholder", LanguageID) & "</option>")
                    Send("<option value=""" & NewCardholdersID & """ style=""color:brown;font-weight:bold;"">" & Copient.PhraseLib.Lookup("term.newcardholders", LanguageID) & "</option>")
                Else
                    Send("<option value=""" & AllCAM & """ style=""color:brown;font-weight:bold;"">" & Copient.PhraseLib.Lookup("term.allcam", LanguageID) & "</option>")
                End If

                Dim topString As String = ""
                If (RECORD_LIMIT > 0) Then topString = " top " & RECORD_LIMIT & " "

                If (EngineID <> 6) Then
                    'MyCommon.QueryStr = "Select " & topString & " CustomerGroupID,Name from CustomerGroups with (NoLock) where Deleted=0 and AnyCustomer<>1 and CustomerGroupID<>2 and CustomerGroupID is not null and BannerID is null and NewCardholders=0 and CAMCustomerGroup<>1 order by CustomerGroupID desc, Name;"
                    MyCommon.QueryStr = "SELECT DISTINCT " & topString & " CG.CustomerGroupID, CG.Name " &
                       "FROM CustomerGroups CG With (NOLOCK) " &
                       "LEFT JOIN ExtSegmentMap ER on CG.CustomerGroupID = ER.InternalId " &
                       "WHERE CG.Deleted = 0 And AnyCustomer <> 1 And CustomerGroupID <> 2 " &
                       "AND CustomerGroupID IS NOT NULL AND BannerID IS NULL " &
                       "And NewCardholders = 0 And CAMCustomerGroup <> 1 AND CG.Deleted = 0 " &
                       "AND ((ER.SegmentTypeID IS NULL OR ER.SegmentTypeID = 1 OR ER.SegmentTypeID = 3) AND " &
                       "(ER.IncentiveId = " & OfferID & " OR ER.ExtSegmentID > 0 OR ER.ExtSegmentID IS NULL)) " &
                       " And CG.isOptInGroup = 0 ORDER BY CG.Name"
                Else
                    'MyCommon.QueryStr = "Select " & topString & " CustomerGroupID,Name from CustomerGroups with (NoLock) where Deleted=0 and AnyCustomer<>1 and CustomerGroupID<>2 and CustomerGroupID is not null and BannerID is null and NewCardholders=0 and CAMCustomerGroup=1 and AnyCAMCardholder<>1 order by CustomerGroupID desc, Name;"
                    MyCommon.QueryStr = "SELECT DISTINCT " & topString & " CG.CustomerGroupID, CG.Name " &
                           "FROM CustomerGroups CG With (NOLOCK) " &
                           "LEFT JOIN ExtSegmentMap ER on CG.CustomerGroupID = ER.InternalId " &
                           "WHERE CG.Deleted = 0 And AnyCustomer <> 1 And CustomerGroupID <> 2 " &
                           "AND CustomerGroupID IS NOT NULL AND BannerID IS NULL " &
                           "And NewCardholders = 0  and CAMCustomerGroup=1 and AnyCAMCardholder<>1 AND CG.Deleted = 0 " &
                           "AND ((ER.SegmentTypeID IS NULL OR ER.SegmentTypeID = 1 OR ER.SegmentTypeID = 3) AND " &
                           "(ER.IncentiveId = " & OfferID & " OR ER.ExtSegmentID > 0 OR ER.ExtSegmentID IS NULL)) " &
                           " And CG.isOptInGroup = 0 ORDER BY CG.Name"
                End If

                Dim groupDT As DataTable = MyCommon.LRT_Select()
                If (groupDT.Rows.Count <= RECORD_LIMIT AndAlso groupDT.Rows.Count > 0) Then
                    For Each groupRow As DataRow In groupDT.Rows
                        Send("<option value=""" & MyCommon.NZ(groupRow.Item("CustomerGroupID"), 0) & """>" & MyCommon.NZ(groupRow.Item("Name"), "").ToString().Replace("'", "\'") & "</option>")
                    Next
                End If

            %>        
          </select>
        </div>
        <%If (RECORD_LIMIT > 0) Then
            Send(Copient.PhraseLib.Lookup("groups.display", LanguageID) & ": " & RECORD_LIMIT.ToString() & "<br />")
          End If
        %>
       
        <br class="half" />
        <b><% Sendb(Copient.PhraseLib.Lookup("term.selectedcustomers", LanguageID))%>:</b>
        <br />
        <input type="button" class="regular select" id="select1" name="select1" value="&#9660; <% Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID))%>" onclick="handleSelectClick('select1');" />&nbsp;
        <input type="button" class="regular deselect" id="deselect1" name="deselect1" value="<% Sendb(Copient.PhraseLib.Lookup("term.deselect", LanguageID))%> &#9650;" disabled="disabled" onclick="handleSelectClick('deselect1');" /><br />
        <br class="half" />
        <select class="longer" id="selected" name="selected" multiple="multiple" size="2"<% sendb(disabledattribute) %>>
          <%
            ' alright lets find the currently selected groups on page load
            MyCommon.QueryStr = "select CG.CustomerGroupID,Name from CPE_IncentiveCustomerGroups as ICG with (NoLock) left join CustomerGroups as CG with (NoLock) " & _
                " on CG.CustomerGroupID=ICG.CustomerGroupID where RewardOptionID=" & roid & _
                " and ICG.deleted=0 and ExcludedUsers=0 and ICG.CustomerGroupID is not null"
            If Request.QueryString("mode") = "optout" Then
              MyCommon.QueryStr = MyCommon.QueryStr & " and IsOptInGroup=0"
            End If
            rst = MyCommon.LRT_Select
            For Each row In rst.Rows
              If MyCommon.NZ(row.Item("CustomerGroupID"), 0) = 1 Then
                Send("<option value=""1"" style=""color:brown;font-weight:bold;"">" & StrConv(Copient.PhraseLib.Lookup("term.anycustomer", LanguageID), VbStrConv.ProperCase) & "</option>")
              ElseIf MyCommon.NZ(row.Item("CustomerGroupID"), 0) = 2 Then
                Send("<option value=""2"" style=""color:brown;font-weight:bold;"">" & Copient.PhraseLib.Lookup("term.anycardholder", LanguageID) & "</option>")
              ElseIf MyCommon.NZ(row.Item("CustomerGroupID"), 0) = NewCardholdersID Then
                Send("<option value=""" & NewCardholdersID & """ style=""color:brown;font-weight:bold;"">" & Copient.PhraseLib.Lookup("term.newcardholders", LanguageID) & "</option>")
              Else
                Send("<option value=""" & MyCommon.NZ(row.Item("CustomerGroupID"), 0) & """>" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
              End If
            Next
          %>
        </select>
        <br />
        <br class="half" />
        <b><% Sendb(Copient.PhraseLib.Lookup("term.excludedcustomers", LanguageID))%>:</b>
        <br />
        <input type="button" class="regular select" id="select2" name="select2" value="&#9660; <% Sendb(Copient.PhraseLib.Lookup("term.select", LanguageID))%>" disabled="disabled" onclick="handleSelectClick('select2');" />&nbsp;
        <input type="button" class="regular deselect" id="deselect2" name="deselect2" value="<% Sendb(Copient.PhraseLib.Lookup("term.deselect", LanguageID))%> &#9650;" disabled="disabled" onclick="handleSelectClick('deselect2');" /><br />
        <br class="half" />
        <select class="longer" id="excluded" name="excluded" multiple="multiple" size="2"<% sendb(disabledattribute) %>>
          <%
            ' alright lets find the currently selected groups on page load
            MyCommon.QueryStr = "select CG.CustomerGroupID,Name from CPE_IncentiveCustomerGroups as ICG with (NoLock) left join CustomerGroups as CG with (NoLock) " & _
                                " on CG.CustomerGroupID=ICG.CustomerGroupID where RewardOptionID=" & roid & _
                                " and ICG.deleted=0 and ExcludedUsers=1 and ICG.CustomerGroupID is not null"
            If Request.QueryString("mode") = "optout" Then
              MyCommon.QueryStr = MyCommon.QueryStr & " and IsOptInGroup=0"
            End If
            rst = MyCommon.LRT_Select
            For Each row In rst.Rows
              Send("<option value=""" & MyCommon.NZ(row.Item("CustomerGroupID"), 0) & """>" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
            Next
          %>
        </select>
        <hr class="hidden" />
      </div>
    </div>
    
    <div id="gutter">
    </div>
    <% If (EngineID <> 6) Then%>
    <div id="column2">
      <div class="box" id="hhoptions">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.household", LanguageID))%>
          </span>
        </h2>
        <input type="checkbox" class="tempcheck" id="household" name="household"<% If (Household) Then Sendb(" checked=""checked""") %><% Sendb(DisabledAttribute) %> />&nbsp;
        <label for="household"><% Sendb(Copient.PhraseLib.Lookup("term.enable", LanguageID))%>&nbsp;<% Sendb(StrConv(Copient.PhraseLib.Lookup("term.householding", LanguageID), VbStrConv.Lowercase))%> </label>
      </div>
    </div>
    <% End If%>
  </div>
</form>

<script type="text/javascript">
<% If (CloseAfterSave) Then %>
    window.close();
<% Else %>
    removeUsed(false);
    updateButtons();
    updateExceptionButtons();
<% End If %>
</script>

<%
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd("mainform", "functioninput")
  MyCommon = Nothing
  Logix = Nothing
%>
