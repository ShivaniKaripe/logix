<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.CommonInc" %>
<%@ Import Namespace="System.Collections.Generic" %>
<script type="text/javascript">

    var nVer = navigator.appVersion;
    var nAgt = navigator.userAgent;
    var browserName = navigator.appName;
    var nameOffset, verOffset, ix;

    var browser = navigator.appName;

    // In Opera, the true version is after "Opera" or after "Version"
    if ((verOffset = nAgt.indexOf("Opera")) != -1) {
        browserName = "Opera";
    }
    // In MSIE, the true version is after "MSIE" in userAgent
    else if ((verOffset = nAgt.indexOf("MSIE")) != -1) {
        browserName = "IE";
    }
    // In Chrome, the true version is after "Chrome" 
    else if ((verOffset = nAgt.indexOf("Chrome")) != -1) {
        browserName = "Chrome";
    }
    // In Safari, the true version is after "Safari" or after "Version" 
    else if ((verOffset = nAgt.indexOf("Safari")) != -1) {
        browserName = "Safari";
    }
    // In Firefox, the true version is after "Firefox" 
    else if ((verOffset = nAgt.indexOf("Firefox")) != -1) {
        browserName = "Firefox";
    }
    // In most other browsers, "name/version" is at the end of userAgent 
    else if ((nameOffset = nAgt.lastIndexOf(' ') + 1) <
          (verOffset = nAgt.lastIndexOf('/'))) {
        browserName = nAgt.substring(nameOffset, verOffset);
        fullVersion = nAgt.substring(verOffset + 1);
        if (browserName.toLowerCase() == browserName.toUpperCase()) {
            browserName = navigator.appName;
        }
    }


    if (browserName == "IE") {
        document.attachEvent("onclick", PageClick);
    }
    else {
        document.onclick = function (evt) {
            var target = document.all ? event.srcElement : evt.target;
            if (target.href) {
                if (IsFormChanged(document.mainform)) {
                    var bConfirm = confirm('Warning:  Unsaved data will be lost.  Are you sure you wish to continue?');
                    return bConfirm;
                }
            }
        };
    }
    function PageClick(evt) {
        var target = document.all ? event.srcElement : evt.target;

        if (target.href) {
            if (IsFormChanged(document.mainform)) {
                var bConfirm = confirm('Warning:  Unsaved data will be lost.  Are you sure you wish to continue?');
                return bConfirm;
            }
        }
    }

   
</script>
<%
    ' *****************************************************************************
    ' * FILENAME: department-edit.aspx 
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
    Dim OptionVal As String
    Dim DeptID As Long = -1
    Dim DeptName As String
    Dim ExternalID As String
    Dim LastUpdate As String
    Dim MyCommon As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
    Dim dst As DataTable
    Dim rst As DataTable
    Dim rst2 As DataTable
    Dim rst3 As DataTable
    Dim rstAssociated As DataTable = Nothing
    Dim row As DataRow
    Dim bSave As Boolean
    Dim bDelete As Boolean
    Dim bCreate As Boolean
    Dim DeptNameTitle As String = ""
    Dim ShowActionButton As Boolean = False
    Dim infoMessage As String = ""
    Dim Handheld As Boolean = False
    Dim BannerID As Integer = 0
    Dim BannerName As String = ""
    Dim BannersEnabled As Boolean = False
    Dim HasAssociatedOffers As Boolean = False

    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If

    MyCommon.AppName = "department-edit.aspx"
    Response.Expires = 0

    Try
        MyCommon.Open_LogixRT()
        AdminUserID = Verify_AdminUser(MyCommon, Logix)

        ' fill in if it was a get method
        If Request.RequestType = "GET" Then
            DeptID = IIf(Request.QueryString("DeptID") = "", -1, MyCommon.Extract_Val(Request.QueryString("DeptID")))
            DeptName = Logix.TrimAll(Request.QueryString("DeptName"))
            ExternalID = Logix.TrimAll(Request.QueryString("ExtID"))
            BannerID = MyCommon.Extract_Val(Request.QueryString("BannerID"))
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
            DeptID = IIf(Request.Form("DeptID") = "", -1, MyCommon.Extract_Val(Request.Form("DeptID")))
            If DeptID <= 0 Then
                DeptID = IIf(Request.QueryString("DeptID") = "", -1, MyCommon.Extract_Val(Request.QueryString("DeptID")))
            End If
            DeptName = Logix.TrimAll(Request.Form("DeptName"))
            ExternalID = Logix.TrimAll(Request.Form("ExtID"))
            If ExternalID = "" Then
                ExternalID = Logix.TrimAll(Request.QueryString("ExtID"))
            End If
            BannerID = MyCommon.Extract_Val(Request.Form("BannerID"))
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

        BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")

        Send_HeadBegin("term.department", , DeptID)
        Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
        Send_Metas()
        Send_Links(Handheld)
        Send_Scripts()
        Send_HeadEnd()
        Send_BodyBegin(1)
        Send_Bar(Handheld)
        Send_Help(CopientFileName)
        Send_Logos()
        Send_Tabs(Logix, 8)
        Send_Subtabs(Logix, 8, 4)

        If (Logix.UserRoles.EditSystemConfiguration = False) Then
            Send_Denied(1, "perm.admin-configuration")
            GoTo done
        End If

        If (Request.QueryString("new") <> "") Then
            Response.Redirect("department-edit.aspx")
        End If

        If bSave Then
            If (DeptName = "") Or (ExternalID = "") Then
                infoMessage = Copient.PhraseLib.Lookup("departments.noname", LanguageID)
            ElseIf (ExternalID = "0000") Then
                infoMessage = Copient.PhraseLib.Lookup("departments.numberused", LanguageID)
            ElseIf (CleanUPC(ExternalID) = "False") Then
                infoMessage = Copient.PhraseLib.Lookup("departments.badcode", LanguageID)
            Else
                If (DeptID = -1) Then
                    MyCommon.QueryStr = "SELECT ChargeBackDeptID FROM ChargeBackDepts with (NoLock) WHERE Name = '" & MyCommon.Parse_Quotes(DeptName) & "'"
                    If (BannersEnabled) Then
                        MyCommon.QueryStr &= " and BannerID=" & BannerID
                    End If
                    dst = MyCommon.LRT_Select
                    If (dst.Rows.Count > 0) Then
                        infoMessage = Copient.PhraseLib.Lookup("departments.nameused", LanguageID)
                    Else
                        MyCommon.QueryStr = "SELECT ChargeBackDeptID FROM ChargeBackDepts with (NoLock) WHERE ExternalID = '" & ExternalID & "'"
                        If (BannersEnabled) Then
                            MyCommon.QueryStr &= " and BannerID=" & BannerID
                        End If
                        dst = MyCommon.LRT_Select
                        If (dst.Rows.Count > 0) Then
                            infoMessage = Copient.PhraseLib.Lookup("departments.numberused", LanguageID)
                        Else
                            MyCommon.QueryStr = "dbo.pt_ChargebackDept_Insert"
                            MyCommon.Open_LRTsp()
                            MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 200).Value = DeptName
                            MyCommon.LRTsp.Parameters.Add("@ExternalID", SqlDbType.NVarChar, 120).Value = ExternalID
                            MyCommon.LRTsp.Parameters.Add("@BannerID", SqlDbType.Int).Value = BannerID
                            MyCommon.LRTsp.Parameters.Add("@ChargebackDeptID", SqlDbType.Int).Direction = ParameterDirection.Output
                            MyCommon.LRTsp.ExecuteNonQuery()
                            DeptID = MyCommon.LRTsp.Parameters("@ChargebackDeptID").Value
                            MyCommon.Activity_Log(17, DeptID, AdminUserID, Copient.PhraseLib.Lookup("history.department-create", LanguageID))
                        End If
                    End If
                Else
                    ' update the existing department
                    MyCommon.QueryStr = "SELECT ChargeBackDeptID FROM ChargeBackDepts with (NoLock) WHERE Name = '" & MyCommon.Parse_Quotes(DeptName) & "' and ChargebackDeptID <> " & DeptID
                    If (BannersEnabled) Then
                        MyCommon.QueryStr &= " and BannerID=" & BannerID
                    End If
                    dst = MyCommon.LRT_Select
                    If (dst.Rows.Count > 0) Then
                        infoMessage = Copient.PhraseLib.Lookup("departments.nameused", LanguageID)
                    Else
                        MyCommon.QueryStr = "SELECT ChargeBackDeptID FROM ChargeBackDepts with (NoLock) WHERE ExternalID = '" & ExternalID & "' and ChargebackDeptID <> " & DeptID
                        If (BannersEnabled) Then
                            MyCommon.QueryStr &= " and BannerID=" & BannerID
                        End If
                        dst = MyCommon.LRT_Select
                        If (dst.Rows.Count > 0) Then
                            infoMessage = Copient.PhraseLib.Lookup("departments.numberused", LanguageID)
                        Else
                            MyCommon.QueryStr = "update ChargeBackDepts with (RowLock) set Name='" & MyCommon.Parse_Quotes(DeptName) & "', ExternalID='" & ExternalID & "', " & _
                                                "LastUpdate=getdate() " & _
                                                "where ChargebackDeptID=" & DeptID
                            MyCommon.LRT_Execute()
                            MyCommon.Activity_Log(17, DeptID, AdminUserID, Copient.PhraseLib.Lookup("history.department-edit", LanguageID))
                        End If
                    End If
                End If
            End If

        ElseIf bDelete Then
            If (DeptID = 0) Then
                infoMessage = Copient.PhraseLib.Lookup("departments.nodelete", LanguageID)
            ElseIf (DeptID > 0) Then
                'Prevent delete if dept is in use in cpe system options
                OptionVal = Convert.ToString(ExternalID)
                MyCommon.QueryStr = "select OptionValue from CPE_SystemOptions with (nolock) where OptionName like '%department%' and OptionValue = '" & OptionVal & "';"
                rst = MyCommon.LRT_Select
                MyCommon.QueryStr = "select OptionValue from InterfaceOptions with (nolock) where OptionName like '%department%' and OptionValue = '" & OptionVal & "';"
                rst2 = MyCommon.LRT_Select
                MyCommon.QueryStr = "select OptionValue from UE_SystemOptions with (nolock) where OptionName like '%department%' and OptionValue = '" & OptionVal & "';"
                rst3 = MyCommon.LRT_Select

                If ((rst.Rows.Count > 0) Or (rst2.Rows.Count > 0) Or (rst3.Rows.Count > 0)) Then
                    infoMessage = Copient.PhraseLib.Lookup("departments.inuse", LanguageID)
                Else
                    MyCommon.QueryStr = "select O.offerid from offers as O with (NoLock) left join offerrewards as OFR with (NoLock) on OFR.offerid=O.offerid " & _
                                        "left join discounts as DISC with (NoLock) on DISC.discountid=OFR.linkid " & _
                                        "where OFR.rewardtypeid=1 and O.prodenddate>getdate() and " & _
                                        "O.deleted=0 and DISC.chargebackdeptid=" & DeptID
                    rst = MyCommon.LRT_Select
                    If (rst.Rows.Count > 0) Then
                        infoMessage = Copient.PhraseLib.Lookup("departments.inuse", LanguageID)
                    Else
                        MyCommon.QueryStr = "DELETE FROM ChargeBackDepts with (RowLock) WHERE ChargeBackDeptID = " & DeptID
                        MyCommon.LRT_Execute()
                        MyCommon.Activity_Log(17, DeptID, AdminUserID, Copient.PhraseLib.Lookup("history.department-delete", LanguageID))
                        Response.Status = "301 Moved Permanently"
                        Response.AddHeader("Location", "department-list.aspx")
                    End If
                End If
            End If
        End If

        LastUpdate = ""

        If Not bCreate Then
            ' no one clicked anything
            If DeptID <> -1 OrElse (DeptID = -1 AndAlso bSave) Then
                MyCommon.QueryStr = "select ChargebackDeptID, ExternalID, CD.Name, CD.LastUpdate, PhraseID, CD.BannerID, BAN.Name as BannerName " & _
                                    "from ChargebackDepts CD with (NoLock) " & _
                                    "left join Banners BAN with (NoLock) on CD.BannerID = BAN.BannerID and BAN.Deleted=0  " & _
                                        "where CD.Deleted=0 and CD.ChargebackDeptID =@DeptID"
                MyCommon.DBParameters.Add("@DeptID", SqlDbType.BigInt).Value = DeptID
            Else
                MyCommon.QueryStr = "select ChargebackDeptID, ExternalID, CD.Name, CD.LastUpdate, PhraseID, CD.BannerID, BAN.Name as BannerName " & _
                         "from ChargebackDepts CD with (NoLock) " & _
                         "left join Banners BAN with (NoLock) on CD.BannerID = BAN.BannerID and BAN.Deleted=0  " & _
                         "where CD.Deleted=0 and CD.ExternalID=@ExternalID"
                MyCommon.DBParameters.Add("@ExternalID", SqlDbType.NVarChar, 120).Value = ExternalID
            End If
            rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
            If (rst.Rows.Count > 0) Then
                DeptID = MyCommon.NZ(rst.Rows(0).Item("ChargebackDeptID"), -1)
                DeptName = MyCommon.NZ(rst.Rows(0).Item("Name"), "")
                ExternalID = MyCommon.NZ(rst.Rows(0).Item("ExternalID"), "")
                BannerID = MyCommon.NZ(rst.Rows(0).Item("BannerID"), 0)
                BannerName = MyCommon.NZ(rst.Rows(0).Item("BannerName"), Copient.PhraseLib.Lookup("term.multiple-banners", LanguageID))
                If (IsDBNull(rst.Rows(0).Item("LastUpdate"))) Then
                    LastUpdate = Copient.PhraseLib.Lookup("term.unknown", LanguageID)
                Else
                    LastUpdate = Logix.ToLongDateTimeString(rst.Rows(0).Item("LastUpdate"), MyCommon)
                End If
            ElseIf (DeptID > 0) Then
                Send("")
                Send("<div id=""intro"">")
                Send("    <h1 id=""title"">" & Copient.PhraseLib.Lookup("term.department", LanguageID) & " #" & DeptID & "</h1>")
                Send("</div>")
                Send("<div id=""main"">")
                Send("    <div id=""infobar"" class=""red-background"">")
                Send("        " & Copient.PhraseLib.Lookup("error.deleted", LanguageID))
                Send("    </div>")
                Send("</div>")
                GoTo done
            End If
        End If

        MyCommon.QueryStr = "select distinct I.IncentiveID as OfferID, I.IncentiveName as OfferName, I.EndDate as ProdEndDate,buy.ExternalBuyerId as BuyerID from CPE_Incentives I " & _
                        "inner join CPE_RewardOptions RO with (NoLock) on RO.IncentiveID = I.IncentiveID and I.Deleted=0 and RO.Deleted=0 " & _
                        "inner join CPE_Deliverables DEL with (NoLock) on DEL.RewardOptionID = RO.RewardOptionID and DEL.Deleted = 0 and DEL.DeliverableTypeId = 2 " & _
                        "inner join CPE_Discounts DISC with (NoLock) on DEL.OutputID = DISC.DiscountID and DISC.Deleted=0 " & _
                         "left outer join Buyers as buy with (nolock) on buy.BuyerId= I.BuyerId " & _
                        "where I.IsTemplate = 0 and DISC.ChargebackDeptID = " & DeptID & _
                        "union " & _
                        "select distinct O.OfferId as OfferID, O.Name as OfferName, O.ProdEndDate as ProdEndDate,NULL as BuyerID from Offers O with (NoLock) " & _
                        "inner join OfferRewards OREW with (NoLock) on OREW.OfferID = O.OfferID and OREW.Deleted=0 and O.Deleted=0 " & _
                        "inner join Discounts DISC with (NoLock) on DISC.DiscountID = OREW.LinkID " & _
                        "where O.IsTemplate=0 and DISC.ChargebackDeptID = " & DeptID & _
                        "order by OfferName;"
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
  <input type="hidden" id="DeptID" name="DeptID" value="<% Sendb(DeptID) %>" />
  <div id="intro">
    <%
      Sendb("<h1 id=""title"">")
      If DeptID = -1 Then
        Sendb(Copient.PhraseLib.Lookup("term.newdepartment", LanguageID))
      Else
        Sendb(Copient.PhraseLib.Lookup("term.department", LanguageID) & " #" & DeptID & ": ")
        MyCommon.QueryStr = "SELECT Name FROM ChargebackDepts with (NoLock) WHERE ChargebackDeptID = " & DeptID & ";"
        rst2 = MyCommon.LRT_Select
        If (rst2.Rows.Count > 0) Then
          DeptNameTitle = MyCommon.NZ(rst2.Rows(0).Item("Name"), "")
        End If
        Send(MyCommon.TruncateString(DeptNameTitle, 40))
      End If
      Sendb("</h1>")
    %>
    <div id="controls">
      <%
        If (DeptID = -1) Then
          If (Logix.UserRoles.EditDepartments) Then
            Send_Save()
          End If
        Else
          ShowActionButton = (Logix.UserRoles.EditDepartments)
          If (ShowActionButton) Then
            Send("<input type=""button"" class=""regular"" id=""actions"" name=""actions"" value=""" & Copient.PhraseLib.Lookup("term.actions", LanguageID) & " ▼"" onclick=""toggleDropdown();"" />")
            Send("<div class=""actionsmenu"" id=""actionsmenu"">")
            If (Logix.UserRoles.EditDepartments) Then
              Send_Save()
            End If
            If (Logix.UserRoles.EditDepartments) Then
              If HasAssociatedOffers Then
                Send_Delete("disabled=true")
              Else
                Send_Delete()
              End If
            End If
            If (Logix.UserRoles.EditDepartments) Then
              Send_New()
            End If
            Send("</div>")
          End If
          If MyCommon.Fetch_SystemOption(75) Then
            If (Logix.UserRoles.AccessNotes) Then
              Send_NotesButton(22, DeptID, AdminUserID)
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
        <label id="lblExtID" for="ExtID"><% Sendb(Copient.PhraseLib.Lookup("term.code", LanguageID))%>:</label><br style="line-height: 0.1;" />
        <% Sendb("<input type=""text"" class=""longest"" id=""ExtID"" name=""ExtID"" maxlength=""120"" value=""" & ExternalID & """" & IIf(HasAssociatedOffers Or DeptID > 0, " readonly style=""color:gray;"" ", "") & " />")%>
        <br class="half" />
        <label for="DeptName"><% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>:</label><br />
        <%If (DeptName Is Nothing) Then DeptName = ""
          Sendb("<input type=""text"" class=""longest"" id=""DeptName"" name=""DeptName"" maxlength=""100"" value=""" & DeptName.Replace("""", "&quot;") & """ />")%>
        <br />
        <br class="half" />
        <%
          If (BannersEnabled) Then
            If (DeptID = -1) Then
              MyCommon.QueryStr = "select distinct BAN.BannerID, BAN.Name from Banners BAN with (NoLock) " & _
                                   "inner join AdminUserBanners AUB with (NoLock) on AUB.BannerID = BAN.BannerID " & _
                                   "where BAN.Deleted=0 and BAN.AllBanners=0 and AdminUserID = " & AdminUserID & " order by BAN.Name;"
              rst = MyCommon.LRT_Select
              Send("<label for=""BannerID"">" & Copient.PhraseLib.Lookup("term.banner", LanguageID) & ":</label><br />")
              Send("<select class=""longest"" name=""BannerID"" id=""BannerID"">")
              For Each row In rst.Rows
                Send("  <option value=""" & MyCommon.NZ(row.Item("BannerID"), -1) & """>" & MyCommon.NZ(row.Item("Name"), Copient.PhraseLib.Lookup("term.multiple-banners", LanguageID)) & "</option>")
              Next
              Send("  <option value=""0"">[" & Copient.PhraseLib.Lookup("term.multiple-banners", LanguageID) & "]</option>")
              Send("</select>")
            Else
              Send(Copient.PhraseLib.Lookup("term.banner", LanguageID) & ": " & MyCommon.SplitNonSpacedString(BannerName, 25))
            End If
            Send("<br /><br class=""half"" />")
          End If

          If (DeptID > -1) Then
            Sendb(Copient.PhraseLib.Lookup("term.edited", LanguageID) & " " & LastUpdate)
          End If
        %>
        <hr class="hidden" />
      </div>
    </div>
    
    <div id="gutter">
    </div>
    
    <div id="column2">
      <div class="box" id="offers"<%if(DeptID = -1)then sendb(" style=""visibility: hidden;""") %>>
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.associatedoffers", LanguageID))%>
          </span>
        </h2>
        <div class="boxscroll">
          <% 
              Dim assocName As String=""
            If (DeptID > -1) Then
              If rstAssociated.Rows.Count > 0 Then
                For Each row In rstAssociated.Rows
                    If (MyCommon.Fetch_UE_SystemOption(168) = "1" AndAlso MyCommon.NZ(row.Item("BuyerID"), "") <> "") Then
                    assocName = "Buyer " + row.Item("BuyerID").ToString() + " - " + MyCommon.NZ(row.Item("OfferName"), "").ToString()
                    Else
                    assocName = MyCommon.NZ(row.Item("OfferName"),Copient.PhraseLib.Lookup("term.unknown", LanguageID))
                    End If
                  If (Logix.IsAccessibleOffer(AdminUserID, row.Item("OfferID"))) Then
                    Sendb(" <a href=""offer-redirect.aspx?OfferID=" & row.Item("OfferID") & """>" & assocName & "</a>")
                  Else
                    Sendb(assocName)
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
  } else {
    document.onclick=handlePageClick;
  }
</script>
<%
  If MyCommon.Fetch_SystemOption(75) Then
    If (Logix.UserRoles.AccessNotes) Then
      Send_Notes(22, DeptID, AdminUserID)
    End If
  End If
done:
Finally
  MyCommon.Close_LogixRT()
End Try
Send_BodyEnd("mainform", "ExtID")
MyCommon = Nothing
Logix = Nothing
%>