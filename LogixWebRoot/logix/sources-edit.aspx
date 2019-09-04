<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.CommonInc" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%
  ' *****************************************************************************
  ' * FILENAME: sources-edit.aspx 
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
  Dim SourceID As Long = -1
  Dim SourceName As String
  Dim SourceTitle As String
  Dim SourceDesc As String
  Dim ExtCode As String
  Dim MaxOffers As Integer = 0
  Dim AutoDeploy As Boolean = False
  Dim AutoSendOutbound As Boolean = False
  Dim LastUpdate As String
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim dst As DataTable
  Dim rst As DataTable
  Dim rst2 As DataTable
  Dim rstAssociated As DataTable = Nothing
  Dim row As DataRow
  Dim bSave As Boolean
  Dim bDelete As Boolean
  Dim bCreate As Boolean
  Dim ShowActionButton As Boolean = False
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim HasAssociatedOffers As Boolean = False
  Dim DefaultAsMfgCoupon As Boolean = False
  Dim DefaultAsLogixID As Boolean = False
  Dim ChangeExtID As Boolean = False
  Dim CurrentOffers As Integer
  Dim EnableIssuance As Boolean = False
  Dim MfgTxt As String
  Dim NonMfgTxt As String
  Dim bDefaultReceiptMessageEnabled As Boolean = False

  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "sources-edit.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  Try
    ' fill in if it was a get method
    If Request.RequestType = "GET" Then
      SourceID = IIf(Request.QueryString("SourceID") = "", -1, MyCommon.Extract_Val(Request.QueryString("SourceID")))
      SourceName = Logix.TrimAll(Request.QueryString("SourceName"))
      SourceDesc = Logix.Remove_crlf(Logix.TrimAll(Request.QueryString("SourceDesc")))
      ExtCode = Logix.TrimAll(Request.QueryString("ExtCode"))
      MaxOffers = MyCommon.Extract_Decimal(GetCgiValue("MaxOffers"), MyCommon.GetAdminUser.Culture)
      AutoDeploy = IIf(Request.QueryString("autodeploy") <> "", True, False)
      AutoSendOutbound = IIf(Request.QueryString("AutoSendOutbound") <> "", True, False)
      DefaultAsMfgCoupon = IIf(Request.QueryString("defaultAsMfg") <> "", True, False)
      DefaultAsLogixID = IIf(Request.QueryString("defaultAsLogixID") <> "", True, False)
      ChangeExtID = IIf(Request.QueryString("extID") <> "", True, False)
      EnableIssuance = IIf(Request.QueryString("enableIssuance") <> "", True, False)
      MfgTxt = Logix.TrimAll(Request.QueryString("MfgTxt"))
      NonMfgTxt = Logix.TrimAll(Request.QueryString("NonMfgTxt"))

      If Request.QueryString("save") = "" Then
        bSave = False
      Else
        bSave = True
      End If
      If Request.QueryString("delete") = "" Then
        bDelete = False
      Else
        bDelete = True
      End If
      If Request.QueryString("mode") = "" Then
        bCreate = False
      Else
        bCreate = True
      End If
    Else
      SourceID = IIf(Request.Form("SourceID") = "", -1, MyCommon.Extract_Val(Request.Form("SourceID")))
      If SourceID <= 0 Then
        SourceID = IIf(Request.QueryString("SourceID") = "", -1, MyCommon.Extract_Val(Request.QueryString("SourceID")))
      End If
      SourceName = Logix.TrimAll(Request.Form("SourceName"))
      SourceDesc = Logix.TrimAll(Request.Form("SourceDesc"))
      ExtCode = Logix.TrimAll(Request.Form("ExtCode"))
      MaxOffers = MyCommon.Extract_Decimal(Logix.TrimAll(GetCgiValue("MaxOffers")), MyCommon.GetAdminUser.Culture)
      AutoDeploy = IIf(Request.Form("autodeploy") <> "", True, False)
      AutoSendOutbound = IIf(Request.Form("AutoSendOutbound") <> "", True, False)
      DefaultAsMfgCoupon = IIf(Request.Form("defaultAsMfg") <> "", True, False)
      ChangeExtID = IIf(Request.Form("extID") <> "", True, False)
      EnableIssuance = IIf(Request.Form("enableIssuance") <> "", True, False)
      MfgTxt = Logix.TrimAll(Request.Form("MfgTxt"))
      NonMfgTxt = Logix.TrimAll(Request.Form("NonMfgTxt"))
      
      If Request.Form("save") = "" Then
        bSave = False
      Else
        bSave = True
      End If
      If Request.Form("delete") = "" Then
        bDelete = False
      Else
        bDelete = True
      End If
      If Request.Form("mode") = "" Then
        bCreate = False
      Else
        bCreate = True
      End If
    End If
    
    If MyCommon.Fetch_InterfaceOption(55) = "1" Then
      bDefaultReceiptMessageEnabled = True
    End If

    ' get the current number of offers for this external source
    MyCommon.QueryStr = "select count(*) as IncentiveCount from AllOffersListView " & _
                        "where Deleted=0 and InboundCRMEngineID=" & SourceID & " and DateAdd(d, 1, ProdEndDate) >= getdate();"
    dst = MyCommon.LRT_Select
    If dst.Rows.Count > 0 Then
      CurrentOffers = (MyCommon.NZ(dst.Rows(0).Item("IncentiveCount"), 0))
    End If

    Send_HeadBegin("term.externalsources", , SourceID)
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send_Scripts()
%>
<script type="text/javascript">
  function handleMaxOffers(selValue) {
    var elem = document.getElementById("maxoffers");
    var elemSub1 = document.getElementById("submission1");
    var elemSub2 = document.getElementById("submission2");
    
    if (elem != null && elemSub1 != null && elemSub2 != null) {
      if (selValue == "-1") {
        elem.value = "-1"
        elem.style.display = 'none';
        elemSub1.style.display = 'inline';      
        elemSub2.style.display = 'none';      
      } else if (selValue == "0") {
        elem.value = "0"
        elem.style.display = 'none';      
        elemSub1.style.display = 'inline';      
        elemSub2.style.display = 'none';      
      } else if (selValue == "1") {
        elem.value = ""
        elem.style.display = 'inline';      
        elemSub1.style.display = 'none';      
        elemSub2.style.display = 'inline';      
      }
    }
  }

</script>
<%    
  Send_HeadEnd()
  Send_BodyBegin(1)
  Send_Bar(Handheld)
  Send_Help(CopientFileName)
  Send_Logos()
  Send_Tabs(Logix, 8)
  Send_Subtabs(Logix, 8, 4)
      
  If (Logix.UserRoles.AccessExternalSources = False) Then
    Send_Denied(1, "perm.access-external")
    GoTo done
  End If
      
  If (Request.QueryString("new") <> "") Then
    Response.Redirect("sources-edit.aspx")
  End If
    
  If bSave Then
    If (SourceName = "") OrElse (ExtCode = "") Then
      infoMessage = Copient.PhraseLib.Lookup("sources.noname", LanguageID)
    Else
      If (SourceID = -1) Then
        MyCommon.QueryStr = "SELECT ExtInterfaceID FROM ExtCRMInterfaces with (NoLock) WHERE Deleted=0 and Name = '" & MyCommon.Parse_Quotes(SourceName) & "'"
        dst = MyCommon.LRT_Select
        If (dst.Rows.Count > 0) Then
          infoMessage = Copient.PhraseLib.Lookup("sources.nameused", LanguageID)
        Else
          MyCommon.QueryStr = "SELECT ExtInterfaceID FROM ExtCRMInterfaces with (NoLock) WHERE Deleted=0 and ExtCode = '" & MyCommon.Parse_Quotes(ExtCode) & "'"
          dst = MyCommon.LRT_Select
          If (dst.Rows.Count > 0) Then
            infoMessage = Copient.PhraseLib.Lookup("sources.codeused", LanguageID)
          Else
            MyCommon.QueryStr = "dbo.pt_ExtCRMInterfaces_Insert"
            MyCommon.Open_LRTsp()
            MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 100).Value = SourceName
            MyCommon.LRTsp.Parameters.Add("@Description", SqlDbType.NVarChar, 1000).Value = SourceDesc
            MyCommon.LRTsp.Parameters.Add("@ExtCode", SqlDbType.NVarChar, 50).Value = ExtCode
            MyCommon.LRTsp.Parameters.Add("@MaxOffers", SqlDbType.Int).Value = MaxOffers
            MyCommon.LRTsp.Parameters.Add("@AutoDeploy", SqlDbType.Bit).Value = IIf(AutoDeploy, 1, 0)
            MyCommon.LRTsp.Parameters.Add("@AutoSendOutbound", SqlDbType.Bit).Value = IIf(AutoSendOutbound, 1, 0)
            MyCommon.LRTsp.Parameters.Add("@ExtInterfaceID", SqlDbType.Int).Direction = ParameterDirection.Output
            MyCommon.LRTsp.ExecuteNonQuery()
            SourceID = MyCommon.LRTsp.Parameters("@ExtInterfaceID").Value
            MyCommon.QueryStr = "Update ExtCRMInterfaces set DefaultAsMfgCoupon=@DefaultAsMfgCoupon ," & _
                                "DefaultAsLogixID=@DefaultAsLogixID , AllowExtOfferIDChange=@AllowExtOfferIDChange ," & _
                                "EnableIssuance=@EnableIssuance ," & _
                                "MfgDefaultReceiptMessage=@MfgDefaultReceiptMessage ," & _
                                "NonMfgDefaultReceiptMessage=@NonMfgDefaultReceiptMessage " & _
                                "where ExtInterfaceID = @ExtInterfaceID ;"
            MyCommon.DBParameters.Add("@DefaultAsMfgCoupon", SqlDbType.Bit).Value = IIf(DefaultAsMfgCoupon, 1, 0)
            MyCommon.DBParameters.Add("@DefaultAsLogixID", SqlDbType.Bit).Value = IIf(DefaultAsLogixID, 1, 0)
            MyCommon.DBParameters.Add("@AllowExtOfferIDChange", SqlDbType.Bit).Value = IIf(ChangeExtID, 1, 0)
            MyCommon.DBParameters.Add("@EnableIssuance", SqlDbType.Bit).Value = IIf(EnableIssuance, 1, 0)
            MyCommon.DBParameters.Add("@MfgDefaultReceiptMessage", SqlDbType.NVarChar).Value = MfgTxt
            MyCommon.DBParameters.Add("@NonMfgDefaultReceiptMessage", SqlDbType.NVarChar).Value = NonMfgTxt
            MyCommon.DBParameters.Add("@ExtInterfaceID", SqlDbType.Int).Value = SourceID
            MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
            MyCommon.Activity_Log(35, SourceID, AdminUserID, Copient.PhraseLib.Lookup("history.source-create", LanguageID))
          End If
        End If
      Else
      ' check if the name is already in use
      If infoMessage = "" Then
        MyCommon.QueryStr = "SELECT ExtInterfaceID FROM ExtCRMInterfaces with (NoLock) WHERE Deleted=0 and Name = '" & MyCommon.Parse_Quotes(SourceName) & "' and ExtInterfaceID <> " & SourceID
        dst = MyCommon.LRT_Select
        If (dst.Rows.Count > 0) Then
          infoMessage = Copient.PhraseLib.Lookup("sources.nameused", LanguageID)
        End If
      End If
        
      ' check if code is already in use
      If infoMessage = "" Then
        MyCommon.QueryStr = "SELECT ExtInterfaceID FROM ExtCRMInterfaces with (NoLock) WHERE Deleted=0 and ExtCode = '" & MyCommon.Parse_Quotes(ExtCode) & "' and ExtInterfaceID <> " & SourceID
        dst = MyCommon.LRT_Select
        If (dst.Rows.Count > 0) Then
          infoMessage = Copient.PhraseLib.Lookup("sources.codeused", LanguageID)
        End If
      End If
        
      ' calculate the current number of offers for this external source
      If infoMessage = "" AndAlso MaxOffers > 0 AndAlso MaxOffers < CurrentOffers Then
        infoMessage = Copient.PhraseLib.Lookup("sources.maxoffers-exceeded", LanguageID) & " (" & CurrentOffers & ")"
      End If
        
      ' if no problems, then save the changes
      If infoMessage = "" Then
          MyCommon.QueryStr = "update ExtCRMInterfaces with (RowLock) set Name=@Name , ExtCode=@ExtCode ," & _
                              "MaxOffers=@MaxOffers , Description=@Description , LastUpdate=getdate()," & _
                              "AutoDeploy=@AutoDeploy , DefaultAsMfgCoupon=@DefaultAsMfgCoupon ," & _
                              "DefaultAsLogixID=@DefaultAsLogixID , EnableIssuance=@EnableIssuance ," & _
                              "AllowExtOfferIDChange=@AllowExtOfferIDChange ," & _
                              "MfgDefaultReceiptMessage=@MfgDefaultReceiptMessage ," & _
                              "NonMfgDefaultReceiptMessage=@NonMfgDefaultReceiptMessage, " & _
                              "AutoSendOutbound=@AutoSendOutbound " & _
                              "where ExtInterfaceID=@ExtInterfaceID;"
                    MyCommon.DBParameters.Add("@Name", SqlDbType.NVarChar, 100).Value = SourceName
                    MyCommon.DBParameters.Add("@ExtCode", SqlDbType.NVarChar, 50).Value = ExtCode
          MyCommon.DBParameters.Add("@MaxOffers", SqlDbType.Int).Value = MaxOffers
                    MyCommon.DBParameters.Add("@Description", SqlDbType.NVarChar, 1000).Value = SourceDesc
          MyCommon.DBParameters.Add("@AutoDeploy", SqlDbType.Bit).Value = IIf(AutoDeploy, 1, 0)
          MyCommon.DBParameters.Add("@DefaultAsMfgCoupon", SqlDbType.Bit).Value = IIf(DefaultAsMfgCoupon, 1, 0)
          MyCommon.DBParameters.Add("@DefaultAsLogixID", SqlDbType.Bit).Value = IIf(DefaultAsLogixID, 1, 0)
          MyCommon.DBParameters.Add("@EnableIssuance", SqlDbType.Bit).Value = IIf(EnableIssuance, 1, 0)
          MyCommon.DBParameters.Add("@AllowExtOfferIDChange", SqlDbType.Bit).Value = IIf(ChangeExtID, 1, 0)
          MyCommon.DBParameters.Add("@MfgDefaultReceiptMessage", SqlDbType.NVarChar).Value = MfgTxt
          MyCommon.DBParameters.Add("@NonMfgDefaultReceiptMessage", SqlDbType.NVarChar).Value = NonMfgTxt
          MyCommon.DBParameters.Add("@ExtInterfaceID", SqlDbType.Int).Value = SourceID
          MyCommon.DBParameters.Add("@AutoSendOutbound", SqlDbType.Bit).Value = IIf(AutoSendOutbound, 1, 0)
          MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
          MyCommon.Activity_Log(35, SourceID, AdminUserID, Copient.PhraseLib.Lookup("history.source-edit", LanguageID))
        End If
      End If
    End If
    If infoMessage = "" Then
      Response.Status = "301 Moved Permanently"
      Response.AddHeader("Location", "sources-edit.aspx?SourceID=" & SourceID)
    End If
    
  ElseIf bDelete Then
    If (SourceID = 0) Then
      infoMessage = Copient.PhraseLib.Lookup("sources.nodelete", LanguageID)
    Else
      MyCommon.QueryStr = "update ExtCRMInterfaces with (RowLock) set Deleted=1 where ExtInterfaceID = " & SourceID
      MyCommon.LRT_Execute()

      MyCommon.Activity_Log(35, SourceID, AdminUserID, Copient.PhraseLib.Lookup("history.source-delete", LanguageID))
      Response.Status = "301 Moved Permanently"
      Response.AddHeader("Location", "sources-list.aspx")
    End If
  End If
    
  LastUpdate = ""
    
  If Not bCreate And infoMessage = "" Then
    ' no one clicked anything
    MyCommon.QueryStr = "select ExtInterfaceID, ExtCode, Name, Description, MaxOffers, LastUpdate, " & _
                        "AutoDeploy, DefaultAsMfgCoupon, DefaultAsLogixID, AllowExtOfferIDChange, EnableIssuance, " & _
                        "MfgDefaultReceiptMessage, NonMfgDefaultReceiptMessage, AutoSendOutbound " & _
                        "from ExtCRMInterfaces with (NoLock) " & _
                        "where Editable = 1 and Active=1 and Deleted=0 and ExtInterfaceID=" & SourceID
    rst = MyCommon.LRT_Select()
    If (rst.Rows.Count > 0) Then
      SourceID = MyCommon.NZ(rst.Rows(0).Item("ExtInterfaceID"), -1)
      SourceName = MyCommon.NZ(rst.Rows(0).Item("Name"), "")
      SourceDesc = MyCommon.NZ(rst.Rows(0).Item("Description"), "")
      ExtCode = MyCommon.NZ(rst.Rows(0).Item("ExtCode"), "")
      MaxOffers = MyCommon.NZ(rst.Rows(0).Item("MaxOffers"), 0)
      AutoDeploy = MyCommon.NZ(rst.Rows(0).Item("AutoDeploy"), False)
      AutoSendOutbound = MyCommon.NZ(rst.Rows(0).Item("AutoSendOutbound"), False)
      DefaultAsMfgCoupon = MyCommon.NZ(rst.Rows(0).Item("DefaultAsMfgCoupon"), False)
      DefaultAsLogixID = MyCommon.NZ(rst.Rows(0).Item("DefaultAsLogixID"), False)
      ChangeExtID = MyCommon.NZ(rst.Rows(0).Item("AllowExtOfferIDChange"), False)
      EnableIssuance = MyCommon.NZ(rst.Rows(0).Item("EnableIssuance"), False)
      MfgTxt = MyCommon.NZ(rst.Rows(0).Item("MfgDefaultReceiptMessage"), "")
      NonMfgTxt = MyCommon.NZ(rst.Rows(0).Item("NonMfgDefaultReceiptMessage"), "")
      
      If (IsDBNull(rst.Rows(0).Item("LastUpdate"))) Then
        LastUpdate = Copient.PhraseLib.Lookup("term.unknown", LanguageID)
      Else
        LastUpdate = Logix.ToLongDateTimeString(rst.Rows(0).Item("LastUpdate"), MyCommon)
      End If
    ElseIf (SourceID > 0) Then
      Send("")
      Send("<div id=""intro"">")
      Send("    <h1 id=""title"">" & Copient.PhraseLib.Lookup("term.externalsources", LanguageID) & " #" & SourceID & "</h1>")
      Send("</div>")
      Send("<div id=""main"">")
      Send("    <div id=""infobar"" class=""red-background"">")
      Send("        " & Copient.PhraseLib.Lookup("error.deleted", LanguageID))
      Send("    </div>")
      Send("</div>")
      GoTo done
    End If
  End If
    
  'MyCommon.QueryStr = "select distinct I.IncentiveID as OfferID, I.IncentiveName as OfferName, I.EndDate as ProdEndDate from CPE_Incentives I " & _
  '                "where I.Deleted=0 and InboundCRMEngineID=" & SourceID
  MyCommon.QueryStr = "select OfferID, Name as OfferName, ProdEndDate from AllOffersListView where Deleted=0 and InboundCRMEngineID=" & SourceID
  rstAssociated = MyCommon.LRT_Select
  HasAssociatedOffers = (rstAssociated.Rows.Count > 0)
%>
<script type="text/javascript">
  function toggleDropdown() {
    if (document.getElementById("actionsmenu") != null) {
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
</script>

<form action="#" id="mainform" name="mainform">
  <input type="hidden" id="SourceID" name="SourceID" value="<% Sendb(SourceID) %>" />
  <div id="intro">
    <%Sendb("<h1 id=""title"">")
      If SourceID = -1 Then
        Sendb(Copient.PhraseLib.Lookup("term.newexternalsource", LanguageID))
      Else
        Sendb(Copient.PhraseLib.Lookup("term.externalsource", LanguageID) & " #" & SourceID & ": ")
        MyCommon.QueryStr = "SELECT Name FROM ExtCRMInterfaces with (NoLock) WHERE ExtInterfaceID = " & SourceID & ";"
        rst2 = MyCommon.LRT_Select
        If (rst2.Rows.Count > 0) Then
          SourceTitle = MyCommon.NZ(rst2.Rows(0).Item("Name"), "")
        End If
        Send(MyCommon.TruncateString(SourceTitle, 40))
      End If
      Sendb("</h1>")
    %>
    <div id="controls">
      <%
        If (SourceID = -1) Then
          If (Logix.UserRoles.CreateExternalSource) Then
            Send_Save()
          End If
        Else
          ShowActionButton = (Logix.UserRoles.EditExternalSource)
          If (ShowActionButton) Then
            Send("<input type=""button"" class=""regular"" id=""actions"" name=""actions"" value=""" & Copient.PhraseLib.Lookup("term.actions", LanguageID) & " ▼"" onclick=""toggleDropdown();"" />")
            Send("<div class=""actionsmenu"" id=""actionsmenu"">")
            If (Logix.UserRoles.EditExternalSource) Then
              Send_Save()
            End If
            If (Logix.UserRoles.DeleteExternalSource) Then
              Send_Delete()
            End If
            If (Logix.UserRoles.CreateExternalSource) Then
              Send_New()
            End If
            Send("</div>")
          End If
          If MyCommon.Fetch_SystemOption(75) Then
            If (Logix.UserRoles.AccessNotes) Then
              Send_NotesButton(33, SourceID, AdminUserID)
            End If
          End If
        End If
      %>
    </div>
  </div>
  <div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <div id="column1">
      <div class="box" id="identification">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.identification", LanguageID))%>
          </span>
        </h2>
        <label id="lblExtCode" for="ExtCode"><% Sendb(Copient.PhraseLib.Lookup("term.code", LanguageID))%>:</label><br style="line-height: 0.1;" />
        <% Sendb("<input type=""text"" class=""longest"" id=""ExtCode"" name=""ExtCode"" maxlength=""50"" value=""" & ExtCode & """" & IIf(HasAssociatedOffers, " readonly style=""color:gray;"" ", "") & " />")%>
        <br class="half" />
        <label for="SourceName"><% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>:</label><br />
        <%If (SourceName Is Nothing) Then SourceName = ""
          Sendb("<input type=""text"" class=""longest"" id=""SourceName"" name=""SourceName"" maxlength=""100"" value=""" & SourceName.Replace("""", "&quot;") & """ />")%>
        <br class="half" />
        <label for="SourceDesc" style="position:relative;"><% Sendb(Copient.PhraseLib.Lookup("term.description", LanguageID))%>:</label><br />
        <textarea class="longest" cols="48" rows="3" id="SourceDesc" name="SourceDesc" oninput="limitText(this,1000);"><% Sendb(SourceDesc)%></textarea><br />
        <br class="half" />
        <br class="half" />
        <% Send(Copient.PhraseLib.Lookup("term.currentoffers", LanguageID) & ": " & CurrentOffers)%>
        <br class="half" />
        <%
          If (SourceID > -1) Then
            Sendb(Copient.PhraseLib.Lookup("term.edited", LanguageID) & " " & LastUpdate)
          End If
        %>
        <hr class="hidden" />
      </div>
      <br class="half" />
      <div class="box" id="options">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.advancedoptions", LanguageID))%>
          </span>
        </h2>
        <select id="submissionType" name="submissionType" onchange="handleMaxOffers(this.value);">
          <option value="-1"<% Sendb(IIf(MaxOffers<0, " selected=""selected""", ""))%>><% Sendb(Copient.PhraseLib.Lookup("term.disabled", LanguageID))%></option>
          <option value="0"<% Sendb(IIf(MaxOffers=0, " selected=""selected""", ""))%>><% Sendb(Copient.PhraseLib.Lookup("term.unlimited", LanguageID))%></option>
          <option value="1"<% Sendb(IIf(MaxOffers>0, " selected=""selected""", ""))%>><% Sendb(Copient.PhraseLib.Lookup("term.limited-to", LanguageID))%></option>
        </select>&nbsp;&nbsp;
        <input type="text" class="shorter" id="maxoffers" name="maxoffers" maxlength="7" value="<% Sendb(IIf(MaxOffers = 0, "", MaxOffers)) %>"
               style="<%Sendb(IIf(MaxOffers>0,"display:inline;", "display:none;"))%>" />
        <span id="submission1" style="<% Sendb(IIf(MaxOffers<=0, " display:inline;", " display:none;"))%>"><% Sendb(Copient.PhraseLib.Lookup("sources.offer-submissions", LanguageID).ToLower)%></span>
        <span id="submission2" style="<% Sendb(IIf(MaxOffers>0, " display:inline;", " display:none;"))%>"><% Sendb(Copient.PhraseLib.Lookup("sources.concurrent-submissions", LanguageID).ToLower)%></span>
        <br />
        <br class="half" />
        <input type="checkbox" name="autodeploy" id="autodeploy" value="1"<% Sendb(IIf(AutoDeploy, " checked=""checked""", "")) %> />
        <label for="autodeploy"><%Sendb(Copient.PhraseLib.Lookup("sources.autodeploy", LanguageID))%></label><br />
        <br class="half" />
        <%if (MyCommon.Fetch_CM_SystemOption(114)) Then %>
          <input type="checkbox" name="autosendoutbound" id="autosendoutbound" value="1"<% Sendb(IIf(AutoSendOutbound, " checked=""checked""", "")) %> />
        <label for="autosendoutbound"><%Sendb(Copient.PhraseLib.Lookup("sources.autosendoutbound", LanguageID))%></label><br />
        <%end if%>
        <br class="half" />
        <input type="checkbox" name="defaultAsMfg" id="defaultAsMfg" value="1"<% Sendb(IIf(DefaultAsMfgCoupon, " checked=""checked""", "")) %> />
        <label for="defaultAsMfg"><%Sendb(Copient.PhraseLib.Lookup("sources.treat-as-mfg-coupon", LanguageID))%></label><br />
        <br class="half" />
        <input type="checkbox" name="defaultAsLogixID" id="defaultAsLogixID" value="1"<% Sendb(IIf(DefaultAsLogixID, " checked=""checked""", "")) %> />
        <label for="defaultAsLogixID"><%Sendb(Copient.PhraseLib.Lookup("sources.use-logixID", LanguageID))%></label><br />
        <br class="half" />
        <input type="checkbox" name="enableIssuance" id="enableIssuance" value="1"<% Sendb(IIf(EnableIssuance, " checked=""checked""", "")) %> />
        <label for="defaultIssuance"><%Sendb(Copient.PhraseLib.Lookup("sources.defaultissuance", LanguageID))%></label><br />
        <br class="half" />
        <input type="checkbox" name="extID" id="extID" value="1" <% Sendb(IIf(ChangeExtID, " checked='checked' ", "")) %>/>
        <label for="extID"><%Sendb(Copient.PhraseLib.Lookup("term.AllowExtOfferChange", LanguageID))%></label><br />
        <br class="half" />
        <label id="Mfg" for="Mfg"><% Sendb(Copient.PhraseLib.Lookup("term.mfgdefaultreceiptmsg", LanguageID))%>:<% If Not bDefaultReceiptMessageEnabled Then Sendb("<font color=""red""> {" & Copient.PhraseLib.Lookup("term.off", LanguageID) & "}</font>")%></label><br />
        <% Sendb("<input type=""text"" class=""longest"" id=""MfgTxt"" name=""MfgTxt"" maxlength=""100"" value=""" & MfgTxt & """ /><br />")%>
        <br class="half" />
        <label for="NonMfg"><% Sendb(Copient.PhraseLib.Lookup("term.nonmfgdefaultreceiptmsg", LanguageID))%>:<% If Not bDefaultReceiptMessageEnabled Then Sendb("<font color=""red""> {" & Copient.PhraseLib.Lookup("term.off", LanguageID) & "}</font>")%></label><br />
        <%If (SourceName Is Nothing) Then SourceName = ""
          Sendb("<input type=""text"" class=""longest"" id=""NonMfgTxt"" name=""NonMfgTxt"" maxlength=""100"" value=""" & NonMfgTxt & """ /><br />")%>
        <hr class="hidden" />
      </div>
    </div>
    
    <div id="gutter">
    </div>
    
    <div id="column2">
      <div class="box" id="offers"<%if(SourceID = -1)then sendb(" style=""visibility: hidden;""") %>>
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.associatedoffers", LanguageID))%>
          </span>
        </h2>
        <div class="boxscroll">
          <% 
            If (SourceID > -1) Then
              If rstAssociated.Rows.Count > 0 Then
                For Each row In rstAssociated.Rows
                  If (Logix.IsAccessibleOffer(AdminUserID, row.Item("OfferID"))) Then
                    Sendb(" <a href=""offer-redirect.aspx?OfferID=" & row.Item("OfferID") & """>" & MyCommon.NZ(row.Item("OfferName"), "") & "</a>")
                  Else
                    Sendb(MyCommon.NZ(row.Item("OfferName"), ""))
                  End If

                  If (MyCommon.NZ(row.Item("ProdEndDate"), Today) < Today) Then
                    Sendb(" (" & Copient.PhraseLib.Lookup("term.expired", LanguageID) & ")")
                  End If
                  Send("<br />")
                Next
              Else
                Send("                " & Copient.PhraseLib.Lookup("term.none", LanguageID) & "<br />")
              End If
            Else
              Send("                " & Copient.PhraseLib.Lookup("term.none", LanguageID) & "<br />")
            End If
          %>
        </div>
        <hr class="hidden" />
      </div>
    </div>
    
    <br clear="all" />
  </div>
</form>

<script type="text/javascript">
    if (window.captureEvents){
        window.captureEvents(Event.CLICK);
        window.onclick=handlePageClick;
    }
    else {
        document.onclick=handlePageClick;
    }
</script>
<%
  If MyCommon.Fetch_SystemOption(75) Then
    If (Logix.UserRoles.AccessNotes) Then
      Send_Notes(33, SourceID, AdminUserID)
    End If
  End If
done:
Finally
  MyCommon.Close_LogixRT()
End Try
Send_BodyEnd("mainform", "ExtCode")
MyCommon = Nothing
Logix = Nothing
%>
