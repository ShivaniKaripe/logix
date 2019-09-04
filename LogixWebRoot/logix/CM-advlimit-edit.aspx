<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Globalization" %>
<%
  ' *****************************************************************************
  ' * FILENAME: CM-advlimit-edit.aspx 
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
  Dim row As System.Data.DataRow
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim dstPrograms As System.Data.DataTable
  Dim dstAssociated As System.Data.DataTable = Nothing
  Dim dstAssociatedPropagate As System.Data.DataTable = Nothing
  Dim iAssociated As Integer = 0
  Dim iAssociatedPropagate As Integer = 0
  Dim iNeedOfferUpdate As Integer = 0
  Dim rst As DataTable
  Dim pgDescription As String
  Dim Desc As String
  Dim pgPromoVarID As String
  Dim pgCreated As String
  Dim pgUpdated As String
  Dim pgName As String
  Dim Name As String = ""
  Dim PromoVarID As String
  Dim LimitID As Long
  Dim assocName As String
  Dim assocID As String
  Dim l_pgID As String
  Dim longDate As New DateTime
  Dim longDateString As String
  Dim NameTitle As String = ""
  Dim statusMessage As String = ""
  Dim ShowActionButton As Boolean = False
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim AutoDelete As Boolean = False
  Dim OfferUpdateNum As Integer = 0
  Dim OfferDeployNum As Integer = 0

  Dim i As Decimal = 0
  Dim counter As Integer = 1
  Dim MaxLength As Integer = 0
  Dim alTypeID As Integer
  Dim alValue As Decimal
  Dim alPeriod As Integer
  Dim alStartDate As Date
  Dim alLastUpdate As Date
  Dim alLastOfferUpdate As Date
  Dim sLimitPeriod As String
  Dim sStartDateStr As String = ""
    Dim XID As String = ""
    Dim LimitPeriodVal As Integer = 0

  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  
  Response.Expires = 0
  MyCommon.AppName = "CM-advlimit-edit.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  If (Request.QueryString("new") <> "") Then
    Response.Redirect("CM-advlimit-edit.aspx")
  End If
  
  l_pgID = MyCommon.Extract_Val(Request.QueryString("LimitID"))
  If (l_pgID > 0) Then
    MyCommon.QueryStr = "pa_CM_AdvancedLimitOffersGet"
    MyCommon.Open_LRTsp()
    MyCommon.LRTsp.Parameters.Add("@LimitID", SqlDbType.BigInt).Value = l_pgID
    dstAssociated = MyCommon.LRTsp_select()
    MyCommon.Close_LRTsp()
    iAssociated = dstAssociated.Rows.Count
    
    MyCommon.QueryStr = "pa_CM_AdvancedLimitOffersPropagateGet"
    MyCommon.Open_LRTsp()
    MyCommon.LRTsp.Parameters.Add("@LimitID", SqlDbType.BigInt).Value = l_pgID
    dstAssociatedPropagate = MyCommon.LRTsp_select()
    MyCommon.Close_LRTsp()
    iAssociatedPropagate = dstAssociatedPropagate.Rows.Count
  End If

  
  ' any GET parms inbound?
  If (Request.QueryString("Delete") <> "") Then
    If (Not dstAssociated Is Nothing) AndAlso (dstAssociated.Rows.Count > 0) Then
      infoMessage = Copient.PhraseLib.Lookup("sv-edit.inuse", LanguageID)
    else
      ' check that there are no deployed offers that use this Advanced Limt
      MyCommon.QueryStr = "pa_CM_AdvancedLimitOffersGet_ST"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@LimitID", SqlDbType.BigInt).Value = l_pgID
      rst = MyCommon.LRTsp_select()
      MyCommon.Close_LRTsp()
      
      If (rst.Rows.Count > 0) Then
        infoMessage = Copient.PhraseLib.Lookup("term.inusedeployment", LanguageID) & " : ("
        For each row in rst.Rows
          infoMessage &= MyCommon.NZ(row.Item("Name"), "") & ","
        Next
        infoMessage = infoMessage.TrimEnd(",") & ")"
      Else
        ' expunge record if there is one
        If (MyCommon.Extract_Val(Request.QueryString("LimitID")) <> "") Then
          MyCommon.QueryStr = "pt_AdvancedLimits_Delete"
          MyCommon.Open_LRTsp()
          MyCommon.LRTsp.Parameters.Add("@LimitID", SqlDbType.BigInt).Value = l_pgID
          MyCommon.LRTsp.ExecuteNonQuery()
          MyCommon.Close_LRTsp()
        End If
        'Record history
        MyCommon.Activity_Log(44, l_pgID, AdminUserID, Copient.PhraseLib.Lookup("history.advlimit-delete", LanguageID))
        Response.Status = "301 Moved Permanently"
        Response.AddHeader("Location", "CM-advlimit-list.aspx")
        GoTo done
      End If
    End If
  ElseIf (Request.QueryString("save") <> "" AndAlso MyCommon.Extract_Val(Request.QueryString("LimitID")) = 0) Then
    ' add a record
    Name = MyCommon.Parse_Quotes(Request.QueryString.Item("name"))
    Name = Logix.TrimAll(Name)
    Desc = MyCommon.Parse_Quotes(Request.QueryString.Item("desc"))
    Desc = Logix.TrimAll(Desc)
    If (Name = "") Then
      infoMessage = Copient.PhraseLib.Lookup("sv-no-name", LanguageID)
    Else
      MyCommon.QueryStr = "select LimitID from CM_AdvancedLimits with (NoLock) where Name = '" & Name & "' and Deleted=0;"
      rst = MyCommon.LRT_Select
      If (rst.Rows.Count > 0) Then
        infoMessage = Copient.PhraseLib.Lookup("layout.duplicatecell", LanguageID)
      Else
        If Request.QueryString.Item("limitperiod") Is Nothing Then
          sLimitPeriod = MyCommon.Parse_Quotes(Request.QueryString.Item("ImpliedPeriod"))
        Else
                    sLimitPeriod = MyCommon.Parse_Quotes(Request.QueryString.Item("limitperiod"))
                    If (Int32.TryParse(sLimitPeriod, LimitPeriodVal)) Then
                        sLimitPeriod = LimitPeriodVal
                    Else
                        sLimitPeriod = 0
                        infoMessage = Copient.PhraseLib.Lookup("term.validNoOfDays", LanguageID)
                    End If
                End If
        
        alStartDate = Now()
        sStartDateStr = alStartDate.ToShortDateString()

        MyCommon.QueryStr = "dbo.pt_AdvancedLimits_Insert"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 200).Value = Request.QueryString.Item("name")
        MyCommon.LRTsp.Parameters.Add("@Description", SqlDbType.NVarChar, 1000).Value = Desc        
        MyCommon.LRTsp.Parameters.Add("@AutoDelete", SqlDbType.Bit).Value = IIf(Request.QueryString("AutoDelete") = "1", 1, 0)
        MyCommon.LRTsp.Parameters.Add("@LimitID", SqlDbType.BigInt).Direction = ParameterDirection.Output
        MyCommon.LRTsp.ExecuteNonQuery()
        LimitID = MyCommon.LRTsp.Parameters("@LimitID").Value
        MyCommon.Close_LRTsp()
        MyCommon.Activity_Log(44, LimitID, AdminUserID, Copient.PhraseLib.Lookup("history.advlimit-create", LanguageID))

        MyCommon.QueryStr = "dbo.pc_AdvancedLimitsVar_Create"
        MyCommon.Open_LXSsp()
        MyCommon.LXSsp.Parameters.Add("@LimitID", SqlDbType.BigInt).Value = LimitID
        MyCommon.LXSsp.Parameters.Add("@VarID", SqlDbType.BigInt).Direction = ParameterDirection.Output
        MyCommon.LXSsp.ExecuteNonQuery()
        PromoVarID = MyCommon.LXSsp.Parameters("@VarID").Value
        MyCommon.Close_LXSsp()

        MyCommon.QueryStr = "update CM_AdvancedLimits with (RowLock) set " & _
                            "Description=N'" & MyCommon.Parse_Quotes(Logix.TrimAll(Request.QueryString.Item("desc"))) & "'," & _
                            "LimitTypeID=" & MyCommon.Parse_Quotes(Request.QueryString.Item("limittype")) & ", " & _
                            "LimitValue=" & MyCommon.Extract_Decimal(GetCgiValue("limitvalue"), MyCommon.GetAdminUser.Culture).ToString(CultureInfo.InvariantCulture) & ", " & _
                            "LimitPeriod=" & sLimitPeriod & ", " & _
                            "PromoVarID = " & PromoVarID & ", " & _
                            "StartDate = '" & sStartDateStr & "' " & _
                            "where LimitID = " & LimitID & ";"
        MyCommon.LRT_Execute()
        
        Response.Status = "301 Moved Permanently"
        Response.AddHeader("Location", "CM-advlimit-edit.aspx?LimitID=" & LimitID)
        GoTo done
      End If
    End If
  ElseIf (((Request.QueryString("save") <> "") Or (Request.QueryString("SaveProp") <> "") Or (Request.QueryString("SavePropDeploy") <> "")) AndAlso MyCommon.Extract_Val(Request.QueryString("LimitID")) > 0) Then
    ' somebody clicked save
    l_pgID = MyCommon.Extract_Val(Request.QueryString("LimitID"))

    AutoDelete = IIf(Request.QueryString("AutoDelete") = "1", True, False)
    
    If MyCommon.Parse_Quotes(Logix.TrimAll(Request.QueryString.Item("name"))) = "" Then
      infoMessage = Copient.PhraseLib.Lookup("sv-no-name", LanguageID)
    Else
      MyCommon.QueryStr = "select LimitID, Name from CM_AdvancedLimits " & _
                          "where Name = '" & MyCommon.Parse_Quotes(Request.QueryString.Item("name")) & "' " & _
                          "and LimitID <> " & l_pgID & " " & _
                          "and Deleted = 0;"
      rst = MyCommon.LRT_Select
      If (rst.Rows.Count > 0) Then
        infoMessage = Copient.PhraseLib.Lookup("point-edit.nameused", LanguageID)
      Else
        If Request.QueryString.Item("startdate") Is Nothing Then
          sStartDateStr = alStartDate.ToShortDateString()
        Else
          Dim StartDate As Date
                    If (Date.TryParse(GetCgiValue("startdate"), MyCommon.GetAdminUser.Culture, System.Globalization.DateTimeStyles.None, StartDate)) Then
                        sStartDateStr = StartDate.ToShortDateString()
                    Else
                        sStartDateStr = Now().ToShortDateString()
                        infoMessage = Copient.PhraseLib.Lookup("Logix-js.EnterValidDate", LanguageID)
                    End If
                End If
         
                If Request.QueryString.Item("limitperiod") Is Nothing Then
                    sLimitPeriod = MyCommon.Parse_Quotes(Request.QueryString.Item("ImpliedPeriod"))
                Else
                    sLimitPeriod = MyCommon.Parse_Quotes(Request.QueryString.Item("limitperiod"))
                    If (Int32.TryParse(sLimitPeriod, LimitPeriodVal)) Then
                        sLimitPeriod = LimitPeriodVal
                    Else
                        sLimitPeriod = 0
                        infoMessage = Copient.PhraseLib.Lookup("term.validNoOfDays", LanguageID)
                    End If
                End If
        
        Dim adName As String = String.Empty
        Dim addesc As String = String.Empty
        MyCommon.QueryStr = "select LimitID, Name, Description from CM_AdvancedLimits where LimitID = " & l_pgID & " " & _
                            "and Deleted = 0;"
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
          adName = MyCommon.NZ(rst.Rows(0).Item("Name"), "")
          addesc = MyCommon.NZ(rst.Rows(0).Item("Description"), "")
        End If


                Dim LimitValue As Decimal = MyCommon.Extract_Decimal(GetCgiValue("limitvalue"), MyCommon.GetAdminUser.Culture)
                MyCommon.QueryStr = "update CM_AdvancedLimits with (RowLock) set " & _
                                    "Name=N'" & MyCommon.Parse_Quotes(Logix.TrimAll(Request.QueryString.Item("name"))) & "'," & _
                                    "Description=N'" & MyCommon.Parse_Quotes(Logix.TrimAll(Request.QueryString.Item("desc"))) & "'," & _
                                    "LimitTypeID=" & MyCommon.Parse_Quotes(Request.QueryString.Item("limittype")) & "," & _
                                    "LimitValue=" & LimitValue.ToString(CultureInfo.InvariantCulture) & "," & _
                                    "LimitPeriod=" & sLimitPeriod & ", " & _
                                    "StartDate = '" & sStartDateStr & "', " & _
                                    "AutoDelete=" & IIf(Request.QueryString("autodelete") = "1", 1, 0)
                If (Int32.TryParse(MyCommon.Parse_Quotes(Request.QueryString.Item("limitperiod")), LimitPeriodVal)) Then
                    If (Integer.Parse(MyCommon.Parse_Quotes(Request.QueryString.Item("limittype"))) <> Integer.Parse(MyCommon.Parse_Quotes(Request.QueryString.Item("OriginalType")))) Or _
                       (LimitValue <> Decimal.Parse(MyCommon.Parse_Quotes(Request.QueryString.Item("OriginalValue")), NumberStyles.AllowDecimalPoint, MyCommon.GetAdminUser.Culture)) Or _
                       (adName <> MyCommon.Parse_Quotes(Request.QueryString.Item("name"))) Or (addesc <> MyCommon.Parse_Quotes(Request.QueryString.Item("desc"))) Or _
                       (Integer.Parse(MyCommon.Parse_Quotes(Request.QueryString.Item("limitperiod"))) <> Integer.Parse(MyCommon.Parse_Quotes(Request.QueryString.Item("OriginalPeriod")))) Then
                        MyCommon.QueryStr &= ",LastUpdate=getDate()"
                    End If
                End If
                MyCommon.QueryStr &= " where LimitID=" & MyCommon.Parse_Quotes(Request.QueryString.Item("LimitID")) & ";"
                MyCommon.LRT_Execute()
                MyCommon.Activity_Log(44, l_pgID, AdminUserID, Copient.PhraseLib.Lookup("history.advlimit-edit", LanguageID))
            End If
        End If

    If (Request.QueryString("SaveProp") <> "") Then
      If (iAssociated > 0) Then
        MyCommon.QueryStr = "dbo.pa_CM_AdvancedLimitPropagate"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@LimitID", SqlDbType.BigInt).Value = l_pgID
        MyCommon.LRTsp.Parameters.Add("@Redeploy", SqlDbType.Bit).Value = 0
        MyCommon.LRTsp.Parameters.Add("@UserId", SqlDbType.Int).Value = AdminUserID
        MyCommon.LRTsp.Parameters.Add("@UpdateHist", SqlDbType.VarChar, 100).Value = Copient.PhraseLib.Lookup("history.advlimit-update", LanguageID)
        MyCommon.LRTsp.Parameters.Add("@DeployHist", SqlDbType.VarChar, 100).Value = Copient.PhraseLib.Lookup("history.advlimit-deploy", LanguageID)
        MyCommon.LRTsp.Parameters.Add("@UpdateNum", SqlDbType.Int).Direction = ParameterDirection.Output
        MyCommon.LRTsp.Parameters.Add("@DeployNum", SqlDbType.Int).Direction = ParameterDirection.Output
        MyCommon.LRTsp.ExecuteNonQuery()
        OfferUpdateNum = MyCommon.LRTsp.Parameters("@UpdateNum").Value
        OfferDeployNum = MyCommon.LRTsp.Parameters("@DeployNum").Value
        MyCommon.Close_LRTsp()
      End If
        
      MyCommon.QueryStr = "update CM_AdvancedLimits with (RowLock) set LastOfferUpdate=getDate() where LimitID = " & l_pgID & ";"
      MyCommon.LRT_Execute()

      statusMessage = OfferUpdateNum & " " & Copient.PhraseLib.Lookup("sv-edit.OffersUpdated", LanguageID)

    ElseIf (Request.QueryString("SavePropDeploy") <> "") Then
      If (iAssociatedPropagate > 0) Then
        MyCommon.QueryStr = "dbo.pa_CM_AdvancedLimitPropagate"
        MyCommon.Open_LRTsp()
        MyCommon.LRTsp.Parameters.Add("@LimitID", SqlDbType.BigInt).Value = l_pgID
        MyCommon.LRTsp.Parameters.Add("@Redeploy", SqlDbType.Bit).Value = 1
        MyCommon.LRTsp.Parameters.Add("@UserId", SqlDbType.Int).Value = AdminUserID
        MyCommon.LRTsp.Parameters.Add("@UpdateHist", SqlDbType.VarChar, 100).Value = Copient.PhraseLib.Lookup("history.advlimit-update", LanguageID)
        MyCommon.LRTsp.Parameters.Add("@DeployHist", SqlDbType.VarChar, 100).Value = Copient.PhraseLib.Lookup("history.advlimit-deploy", LanguageID)
        MyCommon.LRTsp.Parameters.Add("@UpdateNum", SqlDbType.Int).Direction = ParameterDirection.Output
        MyCommon.LRTsp.Parameters.Add("@DeployNum", SqlDbType.Int).Direction = ParameterDirection.Output
        MyCommon.LRTsp.ExecuteNonQuery()
        OfferUpdateNum = MyCommon.LRTsp.Parameters("@UpdateNum").Value
        OfferDeployNum = MyCommon.LRTsp.Parameters("@DeployNum").Value
        MyCommon.Close_LRTsp()
      End If
        
      MyCommon.QueryStr = "update CM_AdvancedLimits with (RowLock) set LastOfferUpdate=getDate() where LimitID = " & l_pgID & ";"
      MyCommon.LRT_Execute()

      statusMessage = OfferUpdateNum & " " & Copient.PhraseLib.Lookup("sv-edit.OffersUpdated", LanguageID)
      statusMessage = statusMessage & " " & OfferDeployNum & " " & Copient.PhraseLib.Lookup("sv-edit.OffersDeployed", LanguageID)
    End If
  ElseIf (Request.QueryString("LimitID") <> "") Then
    ' simple edit/search mode
    l_pgID = MyCommon.NZ(Request.QueryString("LimitID"), "0")
  ElseIf (Request.Form("LimitID") <> "") Then
    l_pgID = MyCommon.Extract_Val(Request.Form("LimitID"))
  Else
    ' no group id passed ... what now ?
    l_pgID = "0"
  End If
  
  MyCommon.QueryStr = "select AL.LimitID, AL.Name, AL.CreatedDate, AL.LastUpdate, AL.Description, " & _
                      "AL.LimitTypeID, AL.LimitValue, AL.LimitPeriod, AL.PromoVarID, AL.AutoDelete, AL.StartDate, " & _
                      "AL.LastUpdate, AL.LastOfferUpdate, AL.ExternalID " & _
                      "from CM_AdvancedLimits AS AL with (NoLock) where Deleted=0 and LimitID='" & l_pgID & "';"
  dstPrograms = MyCommon.LRT_Select
  If (dstPrograms.Rows.Count > 0) Then
    pgDescription = MyCommon.NZ(dstPrograms.Rows(0).Item("Description"), "")
    pgCreated = MyCommon.NZ(dstPrograms.Rows(0).Item("CreatedDate"), "")
    pgUpdated = MyCommon.NZ(dstPrograms.Rows(0).Item("LastUpdate"), "")
    pgPromoVarID = MyCommon.NZ(dstPrograms.Rows(0).Item("PromoVarID"), 0)
    l_pgID = MyCommon.NZ(dstPrograms.Rows(0).Item("LimitID"), 0)
    alTypeID = MyCommon.NZ(dstPrograms.Rows(0).Item("LimitTypeID"), 0)
    alValue = MyCommon.NZ(dstPrograms.Rows(0).Item("LimitValue"), 0)
    alPeriod = MyCommon.NZ(dstPrograms.Rows(0).Item("LimitPeriod"), 0)
    pgName = MyCommon.NZ(dstPrograms.Rows(0).Item("Name"), "")
    AutoDelete = MyCommon.NZ(dstPrograms.Rows(0).Item("AutoDelete"), True)
    alStartDate = MyCommon.NZ(dstPrograms.Rows(0).Item("StartDate"), Now())
    sStartDateStr = Logix.ToShortDateString(alStartDate, MyCommon)
    alLastUpdate = MyCommon.NZ(dstPrograms.Rows(0).Item("LastUpdate"), Now())
    alLastOfferUpdate = MyCommon.NZ(dstPrograms.Rows(0).Item("LastOfferUpdate"), "01/01/2010")
    XID = MyCommon.NZ(dstPrograms.Rows(0).Item("ExternalID"), "")
    
  ElseIf (Request.QueryString("new") <> "New") And (l_pgID > 0) Then
    ' check if this is a deleted Advanced Limit
    MyCommon.QueryStr = "select Name from CM_AdvancedLimits with (NoLock) where LimitID=" & l_pgID & " and deleted =1"
    rst = MyCommon.LRT_Select()
    If (rst.Rows.Count > 0) Then
      pgName = MyCommon.NZ(rst.Rows(0).Item("Name"), "")
    Else
      pgName = ""
    End If
    
    Send_HeadBegin("term.advlimits", , l_pgID)
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send_Scripts()
    Send_Scripts(New String() {"datePicker.js"})
    Send_HeadEnd()
    Send_BodyBegin(1)
    Send_Bar(Handheld)
    Send_Help(CopientFileName)
    Send_Logos()
    Send_Tabs(Logix, 5)
    Send_Subtabs(Logix, 54, 5, , l_pgID)
    Send("")
    Send("<div id=""intro"">")
    Send("    <h1 id=""title"">" & Copient.PhraseLib.Lookup("term.advlimits", LanguageID) & " #" & l_pgID & ": " & pgName & "</h1>")
    Send("</div>")
    Send("<div id=""main"">")
    Send("    <div id=""infobar"" class=""red-background"">")
    Send("        " & Copient.PhraseLib.Lookup("error.deleted", LanguageID))
    Send("    </div>")
    Send("</div>")
    Send_BodyEnd()
    GoTo done
  Else
    pgPromoVarID = 0
    l_pgID = "0"
    pgDescription = ""
    pgCreated = ""
    pgUpdated = ""
    pgName = ""
  End If
  
  Send_HeadBegin("term.advlimits", , l_pgID)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
  Send_Scripts(New String() {"datePicker.js"})
  Send_HeadEnd()
  Send_BodyBegin(1)
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
  Send_Tabs(Logix, 5)
  Send_Subtabs(Logix, 54, 5, , l_pgID)
  
  If (Logix.UserRoles.AccessStoredValuePrograms = False) Then
    Send_Denied(1, "perm.storedvalue-access")
    Send_BodyEnd()
    GoTo done
  End If
%>

<script type="text/javascript" language="javascript">
  var datePickerDivID = "datepicker";
  
  if (window.captureEvents){
    window.captureEvents(Event.CLICK);
    window.onclick=handlePageClick;
  } else {
    document.onclick=handlePageClick;
  }
  
<% Send_Calendar_Overrides(MyCommon) %>

  function handlePageClick(e) {
    var calFrame = document.getElementById('calendariframe');
    var el=(typeof event!=='undefined')? event.srcElement : e.target
    
    if (el != null) {
      var pickerDiv = document.getElementById(datePickerDivID);
      if (pickerDiv != null && pickerDiv.style.visibility == "visible") {
        if (el.id!="start-date-picker") {
          if (!isDatePickerControl(el.className)) {
            pickerDiv.style.visibility = "hidden";
            pickerDiv.style.display = "none";
            if (calFrame != null) {
              calFrame.style.visibility = 'hidden';
              calFrame.style.display = "none";
            }
          }
        } else  {
          pickerDiv.style.visibility = "visible";            
          pickerDiv.style.display = "block";            
          if (calFrame != null) {
            calFrame.style.visibility = 'visible';
            calFrame.style.display = "block";
          }
        }
      }
    }
    if (el.id != 'actions') {
      if (document.getElementById("actionsmenu") != null) {
        var  bOpen = (document.getElementById("actionsmenu").style.visibility == 'visible');
        if (bOpen) {
          toggleDropdown();
        }
      }
    }
  }
  
  function isDatePickerControl(ctrlClass) {
    var retVal = false;
    
    if (ctrlClass != null && ctrlClass.length >= 2) {
      if (ctrlClass.substring(0,2) == "dp") {
        retVal = true;
      }
    }
    
    return retVal;
  }


  <% If (Logix.UserRoles.EditStoredValuePrograms And l_pgID > 0) Then %>
  window.onunload= function(){
    if (document.mainform.name.value != document.mainform.name.defaultValue || document.mainform.desc.value != document.mainform.desc.defaultValue) {
      saveChanges = confirm('<% Sendb(Copient.PhraseLib.Lookup("sv-edit.ChangesMade", LanguageID)) %>');
      if (saveChanges) {
        if (document.mainform.elements['Save'] == null) {
          saveElem = document.createElement("input");
          saveElem.type = 'hidden';
          saveElem.id = 'Save';
          saveElem.name = 'Save';
          saveElem.value = 'save';
          document.mainform.appendChild(saveElem);
        }
        handleAutoFormSubmit();
      }
    }
  };
  <% End If %>
  
  // callback function for save changes on unload during navigate away
  function handleAutoFormSubmit() {
    window.onunload = null;
    document.mainform.action="CM-advlimit-edit.aspx"
    document.mainform.submit();
  }
  
  function disableSaveCheck() {
    window.onunload = null;
    return true;
  }
  
  function toggleDropdown() {
    var bNeedOfferUpdate = false;
  
    if (document.getElementById("actionsmenu") != null) {
      bNeedOfferUpdate = checkNeedOfferUpdate();
      setUpdateDeploy(bNeedOfferUpdate);
      bOpen = (document.getElementById("actionsmenu").style.visibility != 'visible')
      if (bOpen) {
        document.getElementById("actionsmenu").style.visibility = 'visible';
        document.mainform.actions.value = '<% Sendb(Copient.PhraseLib.Lookup("term.actions", LanguageID)) %>▲';
      } else {
        document.getElementById("actionsmenu").style.visibility = 'hidden';
        document.mainform.actions.value = '<% Sendb(Copient.PhraseLib.Lookup("term.actions", LanguageID)) %>▼';
      }
    }
  }
   
  function checkNeedOfferUpdate() {
    var elemCurrentValue = document.getElementById("limitvalue");
    var elemOriginalValue = document.getElementById("OriginalValue");
    var elemCurrentType = document.getElementById("limittype");
    var elemOriginalType = document.getElementById("OriginalType");
    var elemCurrentPeriod = document.getElementById("limitperiod");
    var elemOriginalPeriod = document.getElementById("OriginalPeriod");
    var elemNeedOfferUpdate = document.getElementById("NeedOfferUpdate");
    var bNeedUpdate = false;
 
    if (elemNeedOfferUpdate != null) {
      if (elemNeedOfferUpdate.value != 0) {
        bNeedUpdate = true;
      }
    }
    if (elemCurrentValue != null && elemOriginalValue != null) {
      if (elemCurrentValue.value != elemOriginalValue.value) {
        bNeedUpdate = true;
      }
    }
    if (elemCurrentType != null && elemOriginalType != null) {
      if (elemCurrentType.value != elemOriginalType.value) {
        bNeedUpdate = true;
      }
    }
    if (elemCurrentPeriod != null && elemOriginalPeriod != null) {
      if (elemCurrentPeriod.value != elemOriginalPeriod.value) {
        bNeedUpdate = true;
      }
    }
    return bNeedUpdate;
  }
  
  function setUpdateDeploy(bNeedUpdate) {
    var elemProp = document.getElementById("SaveProp");
    var elemDeploy = document.getElementById("SavePropDeploy");
    
    if (elemProp != null && elemDeploy != null) {
      if (bNeedUpdate) {
        elemProp.disabled = false;
        elemDeploy.disabled = false;
      }
      else {
        elemProp.disabled = true;
        elemDeploy.disabled = true;
      }
    }
  }
  
  function padLeft(str, totalLength) {
    var pd = '';
    
    str = str.toString();
    if (totalLength > str.length) {
      for (var i=0; i < (totalLength-str.length); i++) {
        pd += '0';
      }      
    }
    return pd + str.toString();
  }
  
  function setperiodsection(bSelect) {
    var elemSelectDay = document.getElementById("selectday");
    var elemPeriod=document.getElementById("limitperiod");
    var elemOriginalPeriod=document.getElementById("OriginalPeriod");
    var elemImpliedPeriod=document.getElementById("ImpliedPeriod");
    var elemStartDate=document.getElementById("StartDateDiv");

    if (elemSelectDay != null && (elemSelectDay.value == '2') || (elemSelectDay.value == '3')) {
      if (elemPeriod != null) {
        elemPeriod.style.visibility = 'hidden';
      }
      if (elemStartDate != null) {
        elemStartDate.style.visibility = 'hidden';
      }
      if (elemSelectDay.value == '2') {
        elemImpliedPeriod.value = '0';
        elemPeriod.value = '0';
      }
      else {
        elemImpliedPeriod.value = '-1';
        elemPeriod.value = '-1';
      }
    }
    else {
      if (elemPeriod != null) {
        if (bSelect && elemOriginalPeriod != null) {
          if ((elemOriginalPeriod.value == '-1') || (elemOriginalPeriod.value == '0')) {
            elemPeriod.value = '0';
          }
          else {
            elemPeriod.value = elemOriginalPeriod.value;
            elemImpliedPeriod.value = elemOriginalPeriod.value;
          }
        }
        elemPeriod.style.visibility = 'visible';
      }
      if (elemStartDate != null) {
        elemStartDate.style.visibility = 'visible';
      }
    }
  }

</script>

<form action="#" method="get" id="mainform" name="mainform" onsubmit="return disableSaveCheck();">
  <div id="intro">
    <h1 id="title">
      <%
        If l_pgID = 0 Then
          Sendb(Copient.PhraseLib.Lookup("term.advlimit", LanguageID))
        Else
          Sendb(Copient.PhraseLib.Lookup("term.advlimit", LanguageID) & " #" & l_pgID & ": ")
          MyCommon.QueryStr = "select LimitID,Name from CM_advancedLimits with (NoLock) where LimitId = " & l_pgID & ";"
          rst = MyCommon.LRT_Select
          If (rst.Rows.Count > 0) Then
            NameTitle = MyCommon.NZ(rst.Rows(0).Item("Name"), "")
          End If
          Sendb(MyCommon.TruncateString(NameTitle, 35))
        End If
      %>
    </h1>
    <div id="controls">
      <%
        If (l_pgID = 0) Then
          If (Logix.UserRoles.CreateStoredValuePrograms) Then
            Send_Save()
          End If
        Else
          ShowActionButton = (Logix.UserRoles.CreateStoredValuePrograms) OrElse (Logix.UserRoles.EditStoredValuePrograms) OrElse (Logix.UserRoles.DeleteStoredValuePrograms)
          If (ShowActionButton) Then
            Send("<input type=""button"" class=""regular"" id=""actions"" name=""actions"" value=""" & Copient.PhraseLib.Lookup("term.actions", LanguageID) & " ▼"" onclick=""toggleDropdown();"" />")
            Send("<div class=""actionsmenu"" id=""actionsmenu"">")
            If (Logix.UserRoles.EditStoredValuePrograms) Then
              Send_Save("")
              If (iAssociated > 0 And Logix.UserRoles.UpdateOffersUsingStoredValueProgram) Then
                Send_SV_Propagate("")
                If Date.Compare(alLastUpdate, alLastOfferUpdate) > 0 Then
                  iNeedOfferUpdate = 1
                  statusMessage = Copient.PhraseLib.Lookup("advlimit.need-update", LanguageID)
                End If
              End If
              If (iAssociatedPropagate > 0 And Logix.UserRoles.RedeployOffersUsingStoredValueProgram) Then
                Send_SV_Deploy("")
              End If
              If (Logix.UserRoles.DeleteStoredValuePrograms) Then
                Send_Delete("")
              End If
            End If
            If (Logix.UserRoles.CreateStoredValuePrograms) Then
              Send_New()
            End If
            If Request.Browser.Type = "IE6" Then
              Send("<iframe src=""javascript:'';"" id=""actionsiframe"" frameborder=""0"" marginheight=""0"" marginwidth=""0"" scrolling=""no"" style=""height:75px;""></iframe>")
            End If
            Send("</div>")
            If MyCommon.Fetch_SystemOption(75) Then
              If (Logix.UserRoles.AccessNotes) Then
                Send_NotesButton(8, l_pgID, AdminUserID)
              End If
            End If
          End If
        End If
      %>
    </div>
  </div>
  <div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <% If (statusMessage <> "") Then Send("<div id=""statusbar"" class=""green-background"">" & statusMessage & "</div>")%>
    <div id="column1">
      <div class="box" id="identity">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.identification", LanguageID))%>
          </span>
        </h2>
        <input type="hidden" id="LimitID" name="LimitID" value="<% sendb(l_pgID) %>" />
        <input type="hidden" id="PromoVarID" name="PromoVarID" value="<% sendb(pgPromoVarID) %>" />
        <input type="hidden" id="OriginalPeriod" name="OriginalPeriod" value="<% sendb(alPeriod) %>" />
        <input type="hidden" id="ImpliedPeriod" name="ImpliedPeriod" value="<% sendb(alPeriod) %>" />
        <input type="hidden" id="OriginalType" name="OriginalType" value="<% sendb(alTypeID) %>" />
        <input type="hidden" id="OriginalValue" name="OriginalValue" value="<% sendb(alValue.ToString(MyCommon.GetAdminUser.Culture)) %>" />
        <input type="hidden" id="NeedOfferUpdate" name="NeedOfferUpdate" value="<% sendb(iNeedOfferUpdate) %>" />
        <label for="name">
          <% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>
          :</label><br />
        <% If (pgName Is Nothing) Then pgName = ""%>
        <input type="text" class="longest" id="name" name="name" maxlength="50" value="<% Sendb(pgName.Replace("""", "&quot;")) %>" /><br />
        <label for="desc">
          <% Sendb(Copient.PhraseLib.Lookup("term.description", LanguageID))%>
          :</label><br />
        <textarea class="longest" id="desc" name="desc" cols="48" rows="3"><% Sendb(pgDescription.Replace("""", "&quot;"))%></textarea><br />
        <br class="half" />
        <small><%Send("(" & Copient.PhraseLib.Lookup("CPEoffergen.description", LanguageID) & ")")%></small><br class="half" /><br class="half" />
        
        <%
          If XID <> "" AndAlso XID <> "0" Then
            Sendb(Copient.PhraseLib.Lookup("term.xid", LanguageID) & ": " & XID & "<br />")
          End If
          
            If pgCreated <> "" Then
                Sendb(Copient.PhraseLib.Lookup("term.created", LanguageID) & " ")
                longDate = pgCreated
                longDateString = longDate.ToString("dddd, d MMMM yyyy, HH:mm:ss")
                Send(longDateString)
            End If
          
            If pgUpdated <> "" Then
                Send("<br />")
                Sendb(Copient.PhraseLib.Lookup("term.edited", LanguageID) & " ")
                longDate = pgUpdated
                longDateString = longDate.ToString("dddd, d MMMM yyyy, HH:mm:ss")
                Send(longDateString)
            End If
          
          
          Dim iExtPromoVarId As Integer = 0
          If IsNumeric(MyCommon.Fetch_CM_SystemOption(42)) Then
           If (MyCommon.Fetch_CM_SystemOption(42) = 1) Then
             MyCommon.QueryStr = "select ExternalID from PromoVariables with (NoLock) where PromoVarID=" & pgPromoVarID & ";"
             rst = MyCommon.LXS_Select
             If (rst.Rows.Count > 0) Then
               iExtPromoVarId = MyCommon.NZ(rst.Rows(0).Item(0), 0)
             End If
           End If
            End If
            If pgPromoVarID > 0 Then
                Send("<br />")
                Send("<br class=""half"" />")
                Sendb(Copient.PhraseLib.Lookup("term.contains", LanguageID) & " ")
                If iExtPromoVarId > 0 Then
                    Sendb(Copient.PhraseLib.Lookup("term.var", LanguageID) & ": " & pgPromoVarID & "  (" & Copient.PhraseLib.Lookup("term.xid", LanguageID) & ": " & iExtPromoVarId & ")" & "<br />")
                Else
                    Sendb(Copient.PhraseLib.Lookup("term.var", LanguageID) & ": " & pgPromoVarID & "<br />")
                End If
            End If
          
        %>
        <hr class="hidden" />
      </div>
      <div class="box" id="general">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.general", LanguageID))%>
          </span>
        </h2>
        <%
          MyCommon.QueryStr = "select PhraseID, TypeID from CM_AdvancedLimitTypes with (NoLock) where TypeId > 0 order by TypeID;"
          rst = MyCommon.LRT_Select
          Send("<div" & IIf(l_pgID <> 0, " style=""display:none;""", "") & ">")
          Send("<label for=""limittype"" style=""position:relative;"">" & Copient.PhraseLib.Lookup("term.type", LanguageID) & ":</label> ")
          Send("<select id=""limittype"" name=""limittype""  >")
          For Each row In rst.Rows
            If (MyCommon.NZ(row.Item("TypeID"), -1) = alTypeID) Then
              Send("<option value=""" & MyCommon.NZ(row.Item("TypeID"), -1) & """ selected=""selected"">" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & "</option>")
            Else
              Send("<option value=""" & MyCommon.NZ(row.Item("TypeID"), -1) & """>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & "</option>")
            End If
          Next
          Send("</select><br />")
          Send("</div>")
          If l_pgID <> 0 Then
            Sendb(Copient.PhraseLib.Lookup("term.type", LanguageID) & ": ")
            For Each row In rst.Rows
              If (MyCommon.NZ(row.Item("TypeID"), -1) = alTypeID) Then
                Send(Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseID"), 0), LanguageID) & "<br />")
              End If
            Next
          End If
          Send("<br />")
          Send("<label for=""limitvalue"" style=""position:relative;"">" & Copient.PhraseLib.Lookup("term.limit", LanguageID) & " :   </label>")
          Send("<input type=""text"" class=""shorter"" id=""limitvalue"" name=""limitvalue"" value=""" & alValue.ToString(MyCommon.GetAdminUser.Culture) & """ />")
          Send("<label for=""limitperiod"" style=""position:relative;"">  " & Copient.PhraseLib.Lookup("term.per", LanguageID) & "   </label>")
          Send("<input type=""text"" class=""shorter"" id=""limitperiod"" name=""limitperiod"" value=""" & alPeriod & """ />")
          Send("<select id=""selectday"" name=""selectday"" style=""position:relative;"" onchange=""setperiodsection(true);"" >")
          Select Case alPeriod
            Case -1
              ' per Offer
              Send("<option value=""1"" >" & Copient.PhraseLib.Lookup("term.days", LanguageID) & "</option>")
              Send("<option value=""2"" >" & Copient.PhraseLib.Lookup("term.transaction", LanguageID) & "</option>")
              Send("<option value=""3"" selected=""selected"" >" & Copient.PhraseLib.Lookup("term.customer", LanguageID) & "</option>")
            Case 0
              ' per Transaction
              Send("<option value=""1"" >" & Copient.PhraseLib.Lookup("term.days", LanguageID) & "</option>")
              Send("<option value=""2"" selected=""selected"" >" & Copient.PhraseLib.Lookup("term.transaction", LanguageID) & "</option>")
              Send("<option value=""3"" >" & Copient.PhraseLib.Lookup("term.customer", LanguageID) & "</option>")
            Case Else
              ' per Days
              Send("<option value=""1"" selected=""selected"" >" & Copient.PhraseLib.Lookup("term.days", LanguageID) & "</option>")
              Send("<option value=""2"" >" & Copient.PhraseLib.Lookup("term.transaction", LanguageID) & "</option>")
              Send("<option value=""3"" >" & Copient.PhraseLib.Lookup("term.customer", LanguageID) & "</option>")
          End Select
          Send("</select>")
          Send("<hr class=""hidden""/>")
          Send("<br />")
          Send("<br class=""half"" />")
          
          Send("<div id=""StartDateDiv"">")
          Send(" <label for=""startdate"">")
          Send(Copient.PhraseLib.Lookup("term.startdate", LanguageID) & " :")
          Send(" </label>")
          Send(" <input type=""text"" class=""short"" name=""startdate"" id=""startdate"" value=""" & sStartDateStr & """ />")
          Send(" <img src=""../images/calendar.png"" class=""calendar"" id=""start-date-picker"" alt=""" & Copient.PhraseLib.Lookup("term.datepicker", LanguageID) & """")
          Send("  title=""" & Copient.PhraseLib.Lookup("term.datepicker", LanguageID) & """ onclick=""displayDatePicker('startdate', event);"" />")
          Send("</div>")
          Send("<br class=""half"" />")
          Sendb("<input type=""checkbox"" id=""autodelete"" name=""autodelete"" value=""1""" & IIf(AutoDelete, " checked=""checked""", "") & " />")
          Send("<label for=""autodelete"">" & Copient.PhraseLib.Lookup("sv-edit.AutoDelete", LanguageID) & "</label><br />")
        %>
        <hr class="hidden" />
      </div>
      <%
        Send("<div id=""datepicker"" class=""dpDiv"">")
        Send("</div>")

        If Request.Browser.Type = "IE6" Then
          Send("<iframe src=""javascript:'';"" id=""calendariframe"" frameborder=""0"" marginheight=""0"" marginwidth=""0"" scrolling=""no"" style=""visibility:hidden;display:none;""></iframe>")
        End If
      %>

    </div>
    <div id="gutter">
    </div>
    <div id="column2">
      <% If (l_pgID > 0) Then%>
      <div class="box" id="offers">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.associatedoffers", LanguageID))%>
          </span>
        </h2>
        <div class="boxscroll">
          <% 
            If Not dstAssociated Is Nothing AndAlso dstAssociated.Rows.Count > 0 Then
              For Each row In dstAssociated.Rows
                assocName = row.Item("Name")
                assocID = row.Item("OfferID")

                If (Logix.IsAccessibleOffer(AdminUserID, row.Item("OfferID"))) Then
                  Sendb(" <a href=""offer-redirect.aspx?OfferID=" & assocID & """>" & assocName & "</a><br />")
                Else
                  Sendb(assocName & "<br />")
                End If
              Next
            Else
              Send("  " & Copient.PhraseLib.Lookup("term.none", LanguageID))
            End If
          %>
        </div>
        <hr class="hidden" />
      </div>
      <br clear="all" />
      <% End If%>
    </div>
  </div>
</form>

<script runat="server">
  Public Function AllDigits(ByVal txt As String) As Boolean
    Dim ch As String
    Dim i As Integer
    
    AllDigits = True
    For i = 1 To Len(txt)
      ' See if the next character is a non-digit.
      ch = Mid$(txt, i, 1)
      If ch < "0" Or ch > "9" Then
        ' This is not a digit.
        AllDigits = False
        Exit For
      End If
    Next i
  End Function
</script>

<script type="text/javascript">
<% Send_Date_Picker_Terms() %>

  setperiodsection(false);

  if (window.captureEvents){
    window.captureEvents(Event.CLICK);
    window.onclick=handlePageClick;
  } else {
    document.onclick=handlePageClick;
  }
</script>

<%
  If MyCommon.Fetch_SystemOption(75) Then
    If (l_pgID > 0 And Logix.UserRoles.AccessNotes) Then
      Send_Notes(8, l_pgID, AdminUserID)
    End If
  End If
  Send_BodyEnd("mainform", "name")
done:
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  Logix = Nothing
  MyCommon = Nothing
%>
