<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.CommonInc" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%
    ' *****************************************************************************
    ' * FILENAME: store-detail.aspx 
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
  
    Dim AdminUserID As Long = 0
    Dim MyCommon As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
    Dim LocationID As Long = 0
    Dim LocationName As String = ""
    Dim ExtLocationCode As String = ""
    Dim LocationTypeID As Integer = 1
    Dim Description As String = ""
    Dim ContactName As String = ""
    Dim PhoneNumber As String = ""
    Dim Address1 As String = ""
    Dim Address2 As String = ""
    Dim City As String = ""
    Dim EngineType As Integer
    Dim EngineName As String = ""
    Dim EnginePhraseID As Integer = 0
    Dim TestingLocation As Boolean = False
    Dim State As String = ""
    Dim Zip As String = ""
    Dim CountryID As Integer
    Dim rst As DataTable
    Dim row As DataRow
    Dim rst2 As DataTable
    Dim row2 As DataRow
    Dim dst2 As DataTable
    Dim dst3 As DataTable
    Dim rows() As DataRow
    Dim bSave As Boolean
    Dim bDelete As Boolean
    Dim bCreate As Boolean
    Dim dtGroups As DataTable = Nothing
    Dim dtOffers As DataTable = Nothing
    Dim sQuery As String
    Dim longDate As New DateTime
    Dim lastHeard As String
    Dim lastIPL As String
    Dim generateIPL As String = "&nbsp;"
    Dim bGenerateIPL As Boolean
    Dim deployDate As String
    Dim OptionID As Integer
    Dim LocationOptionValue As String = ""
    Dim tempstr As String = ""
    Dim SendAlert As Boolean
    Dim SanityCheckPassed As Boolean = False
    Dim bWaitingForIPL As Boolean = False
    Dim Cookie As HttpCookie = Nothing
    Dim BoxesValue As String = ""
    Dim ValidateOfferColor As String = "green"
    Dim ValidateProdGroupColor As String = "green"
    Dim ValidateCustGroupColor As String = "green"
    Dim infoMessage As String = ""
    Dim statusMessage As String = ""
    Dim Handheld As Boolean = False
    Dim alertStatus As String = ""
    Dim reportStatus As String = ""
    Dim ReportEnabled As Boolean
    Dim MinutesInError As Long
    Dim CentralHighValue As Integer = 180
    Dim CentralMediumValue As Integer = 180
    Dim CentralLowValue As Integer = 90
    Dim ErrorText As String = ""
    Dim QryStr As String = ""
    Dim SeverityTypes As New Hashtable(5)
    Dim Severity As SeverityEntry
    Dim SeverityDesc As String = ""
    Dim CommsFilter As String = ""
    Dim SearchClause As String = ""
    Dim RowCt, Counter As Integer
    Dim LogFileType As String = ""
    Dim lastFailoverTime As String = String.Empty
    Dim IncentiveFetchURL As String = ""
    Dim ImageFetchURL As String = ""
    Dim PhoneHomeIPOverride As String = String.Empty
    Dim OfflineFTPUser As String = String.Empty
    Dim OfflineFTPPass As String = String.Empty
    Dim OfflineFTPPath As String = String.Empty
    Dim OfflineFTPIP As String = String.Empty
  
    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If
  
    Response.Expires = 0
    MyCommon.AppName = "store-detail.aspx"
    ' Open database connections
    MyCommon.Open_LogixRT()
    MyCommon.Open_LogixWH()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
    Try
        'Sendb("RequestType = " & Request.RequestType)
        ' fill in if it was a get method
        If Request.RequestType = "GET" Then
            LocationID = MyCommon.Extract_Val(Request.QueryString("LocationID"))
            If Request.Form("save") = "" Then
                bSave = False
            Else
                bSave = True
            End If
            If Request.QueryString("mode") = "" Then
                bCreate = False
            Else
                bCreate = True
            End If
            bGenerateIPL = (Request.QueryString("generateIPL") <> "")
        Else
            LocationID = Request.Form("LocationID")
            If LocationID = 0 AndAlso Request.QueryString("LocationID") <> "" Then LocationID = MyCommon.Extract_Val(Request.QueryString("LocationID"))
            If Request.Form("save") = "" Then
                bSave = False
            Else
                bSave = True
            End If
            If Request.Form("mode") = "" Then
                bCreate = False
            Else
                bCreate = True
            End If
            bGenerateIPL = (Request.Form("generateIPL") <> "")
        End If
        
        MyCommon.QueryStr = "select GenIPL from Locations with (NoLock) " & _
               "where LocationID =" & LocationID & " ;"
        rst = MyCommon.LRT_Select()
        row = rst.Rows(0)
        If (row.Item("GenIPL")) Then
            statusMessage = Copient.PhraseLib.Lookup("confirm.redeployed", LanguageID)
        End If
    
    
        ' find the location type
        If Request.QueryString("LocationTypeID") <> "" Then
            LocationTypeID = MyCommon.Extract_Val(Request.QueryString("LocationTypeID"))
        Else
            MyCommon.QueryStr = "select LocationTypeID from Locations with (NoLock) where LocationID=" & LocationID & ";"
            rst = MyCommon.LRT_Select
            If rst.Rows.Count > 0 Then
                LocationTypeID = MyCommon.NZ(rst.Rows(0).Item("LocationTypeID"), 1)
            End If
        End If
    
        ' load up the severity types
        MyCommon.QueryStr = "select HealthSeverityID, Description, PhraseID from LS_HealthSeverityTypes with (NoLock)"
        rst = MyCommon.LWH_Select
        For Each row In rst.Rows
            Severity = New SeverityEntry(MyCommon.NZ(row.Item("Description"), ""), MyCommon.NZ(row.Item("PhraseID"), 0))
            SeverityTypes.Add("Sev" & MyCommon.NZ(row.Item("HealthSeverityID"), "-1").ToString, Severity)
        Next
    
    
        If LocationTypeID = 2 Then
            Send_HeadBegin("term.serverhealth", , LocationID)
        Else
            Send_HeadBegin("term.storehealth", , LocationID)
        End If
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
        Send_Subtabs(Logix, 8, 7)
       
        ' Get the user's preference from the cookie for collapsing/showing the boxes
        Cookie = Request.Cookies("BoxesCollapsed")
        If Not (Cookie Is Nothing) Then
            BoxesValue = Cookie.Value
            If (BoxesValue Is Nothing OrElse BoxesValue.Trim = "") Then
                BoxesValue = "0"
            End If
        Else
            BoxesValue = "0"
        End If
    
        If (Logix.UserRoles.AccessStoreHealth = False) Then
            Send("<script type=""text/javascript"" language=""javascript"">")
            Send("  function updateCookie() { return true; } ")
            Send("</script>")
            Send_Denied(1, "perm.admin-store-health")
            GoTo done
        End If
    
        ' no one clicked anything
        MyCommon.QueryStr = "select L.LocationName, L.TestingLocation, L.Description, L.EngineID, L.ExtLocationCode, L.SendAlert, L.HealthReported, " & _
                            "L.Address1, L.Address2, L.City, L.State, L.CountryID, PE.Description as EngineName, PE.PhraseID as EnginePhraseID " & _
                            "from Locations L with (NoLock) " & _
                            "left join PromoEngines PE with (NoLock) on PE.EngineID=L.EngineID " & _
                            "where Deleted=0 and LocationID=" & LocationID
        rst = MyCommon.LRT_Select()
        If (rst.Rows.Count > 0) Then
            For Each row In rst.Rows
                LocationName = MyCommon.NZ(row.Item("LocationName"), Copient.PhraseLib.Lookup("term.unknown", LanguageID))
                ExtLocationCode = MyCommon.NZ(row.Item("ExtLocationCode"), "=")
                Description = MyCommon.NZ(row.Item("Description"), "")
                TestingLocation = MyCommon.NZ(row.Item("TestingLocation"), False)
                EngineType = MyCommon.NZ(row.Item("EngineID"), 0)
                EngineName = MyCommon.NZ(row.Item("EngineName"), Copient.PhraseLib.Lookup("term.unknown", LanguageID))
                EnginePhraseID = MyCommon.NZ(row.Item("EnginePhraseID"), 0)
                SendAlert = MyCommon.NZ(row.Item("SendAlert"), False)
                ReportEnabled = MyCommon.NZ(row.Item("HealthReported"), False)
                Address1 = MyCommon.NZ(row.Item("Address1"), "")
                Address2 = MyCommon.NZ(row.Item("Address2"), "")
                City = MyCommon.NZ(row.Item("City"), "")
                State = MyCommon.NZ(row.Item("State"), "")
                CountryID = MyCommon.NZ(row.Item("CountryID"), 0)
            Next
        ElseIf (Request.QueryString("new") = "") And (LocationID > 0) Then
            Send("")
            Send("<div id=""intro"">")
            Send("    <h1 id=""title"">" & Copient.PhraseLib.Lookup("term.store", LanguageID) & " #" & LocationID & "</h1>")
            Send("</div>")
            Send("<div id=""main"">")
            Send("    <div id=""infobar"" class=""red-background"">")
            Send("        " & Copient.PhraseLib.Lookup("error.deleted", LanguageID))
            Send("    </div>")
            Send("</div>")
            GoTo done
        End If
    
        ' load up the system options for the severity threshold levels
        If EngineType = 2 Then 'CPE Engine
            If (Not Integer.TryParse(MyCommon.Fetch_CPE_SystemOption(56), CentralHighValue)) Then CentralHighValue = 270
            If (Not Integer.TryParse(MyCommon.Fetch_CPE_SystemOption(57), CentralMediumValue)) Then CentralMediumValue = 180
            If (Not Integer.TryParse(MyCommon.Fetch_CPE_SystemOption(58), CentralLowValue)) Then CentralLowValue = 90
        ElseIf EngineType = 9 Then
            If (Not Integer.TryParse(MyCommon.Fetch_UE_SystemOption(56), CentralHighValue)) Then CentralHighValue = 270
            If (Not Integer.TryParse(MyCommon.Fetch_UE_SystemOption(57), CentralMediumValue)) Then CentralMediumValue = 180
            If (Not Integer.TryParse(MyCommon.Fetch_UE_SystemOption(58), CentralLowValue)) Then CentralLowValue = 90
        End If

        ' Lets see if they clicked save
        If bSave AndAlso infoMessage = "" Then
            If (EngineType = 2) Then
                ImageFetchURL = Trim(Request.Form("imagefetchurl"))
                If Not (Right(ImageFetchURL, 1) = "/") Then
                    ImageFetchURL = ImageFetchURL & "/"
                End If
                ImageFetchURL = MyCommon.Parse_Quotes(ImageFetchURL)
        
                IncentiveFetchURL = Trim(Request.Form("incentivefetchurl"))
                If Not (Right(IncentiveFetchURL, 1) = "/") Then
                    IncentiveFetchURL = IncentiveFetchURL & "/"
                End If
                IncentiveFetchURL = MyCommon.Parse_Quotes(IncentiveFetchURL)
        
                PhoneHomeIPOverride = Trim(Request.Form("PhoneHomeIPOverride"))
                PhoneHomeIPOverride = MyCommon.Parse_Quotes(PhoneHomeIPOverride)
        
                OfflineFTPUser = Trim(Request.Form("FTPUser"))
                OfflineFTPPass = Trim(Request.Form("FTPPass"))
                OfflineFTPPath = Trim(Request.Form("FTPPath"))
                OfflineFTPIP = Trim(Request.Form("FTPIP"))
        
                MyCommon.QueryStr = "Update LocalServers with (RowLock) set ImageFetchURL=@ImageFetchURL, IncentiveFetchURL=@IncentiveFetchURL, PhoneHomeIPOverride=@PhoneHomeIPOverride, " & _
                                    "OfflineFTPUser=@OfflineFTPUser, OfflineFTPIP=@OfflineFTPIP, OfflineFTPPass=@OfflineFTPPass, OfflineFTPPath=@OfflineFTPPath " & _
                                    "where LocationID=@LocationID;"
                MyCommon.DBParameters.Add("@ImageFetchURL", SqlDbType.VarChar).Value = ImageFetchURL
                MyCommon.DBParameters.Add("@IncentiveFetchURL", SqlDbType.VarChar).Value = IncentiveFetchURL
                MyCommon.DBParameters.Add("@PhoneHomeIPOverride", SqlDbType.NVarChar).Value = PhoneHomeIPOverride
                MyCommon.DBParameters.Add("@OfflineFTPUser", SqlDbType.NVarChar).Value = OfflineFTPUser
                MyCommon.DBParameters.Add("@OfflineFTPIP", SqlDbType.NVarChar).Value = OfflineFTPIP
                MyCommon.DBParameters.Add("@OfflineFTPPass", SqlDbType.NVarChar).Value = OfflineFTPPass
                MyCommon.DBParameters.Add("@OfflineFTPPath", SqlDbType.NVarChar).Value = OfflineFTPPath
                MyCommon.DBParameters.Add("@LocationID", SqlDbType.BigInt).Value = LocationID
                MyCommon.ExecuteNonQuery(Copient.DataBases.LogixRT)
            End If
            
            If ((EngineType = 2) Or (EngineType = 9)) Then
               
                If Request.Form("sendalert") = "" Then
                    MyCommon.QueryStr = "Update Locations with (RowLock) set SendAlert=0 where LocationID=" & LocationID & ";"
                Else
                    MyCommon.QueryStr = "Update Locations with (RowLock) set SendAlert=1 where LocationID=" & LocationID & ";"
                End If
                MyCommon.LRT_Execute()
        
                If Request.Form("reportenabled") = "" Then
                    MyCommon.QueryStr = "Update Locations with (RowLock) set HealthReported=0 where LocationID=" & LocationID & ";"
                Else
                    MyCommon.QueryStr = "Update Locations with (RowLock) set HealthReported=1 where LocationID=" & LocationID & ";"
                End If
                MyCommon.LRT_Execute()
        
                MyCommon.Activity_Log(10, LocationID, AdminUserID, Copient.PhraseLib.Lookup("history.store-edit", LanguageID))
                Response.Status = "301 Moved Permanently"
                Response.AddHeader("Location", "store-detail.aspx?LocationID=" & LocationID)
                GoTo done
            End If
        ElseIf bGenerateIPL Then
            MyCommon.QueryStr = "update Locations with (RowLock) set GenIpl=1 where LocationID=" & LocationID
            MyCommon.LRT_Execute()
            statusMessage = Copient.PhraseLib.Lookup("confirm.redeployed", LanguageID)
            If EngineType = 0 Then
                MyCommon.QueryStr = "select LocalServerID from LocalServers with (NoLock) where LocationID=" & LocationID
                rst = MyCommon.LRT_Select
                If (rst.Rows.Count = 0) Then
                    MyCommon.QueryStr = "insert into LocalServers with (RowLock) (LocationID,MustIPl) values (" & LocationID & ",0);"
                    MyCommon.LRT_Execute()
                End If
            End If
            Response.Redirect("store-detail.aspx?LocationID=" & LocationID)
        End If
    
        If Not bCreate Then
            sQuery = "select distinct b.LocationGroupId,b.Name,b.AllLocations from LocGroupItems a with (NoLock),LocationGroups b with (NoLock)"
            sQuery += " where b.LocationGroupId = a.LocationGroupId and a.Deleted = 0 and a.LocationID = " & LocationID
            MyCommon.QueryStr = sQuery
            dtGroups = MyCommon.LRT_Select
            
            sQuery = "select distinct d.OfferId,d.Name from LocGroupItems a with (NoLock),LocationGroups b with (NoLock),OfferLocations c with (NoLock),Offers d with (NoLock)"
            sQuery += " where d.deleted = 0 and d.OfferId = c.OfferId and c.Deleted = 0 and c.Excluded = 0 and c.LocationGroupId = b.LocationGroupId"
            sQuery += " and b.LocationGroupId = a.LocationGroupId and a.Deleted = 0 and a.LocationID = " & LocationID
            MyCommon.QueryStr = sQuery
            dtOffers = MyCommon.LRT_Select
        End If
    
        'Reset IncentiveFetchOffline
        If (Request.QueryString("resetIncentiveFetchOffline") <> "" AndAlso Request.Form("IFOffLineResetDone") Is Nothing) Then
            MyCommon.QueryStr = "Update LocalServers SET IncentiveFetchOffline = 0, IncentiveFetchNAKCount = 0 WHERE LocationID=" & LocationID & ";"
            MyCommon.LRT_Execute()
            If (Request.Form("IFOffLineResetDone") Is Nothing) Then
                Send("<input type=""hidden"" id=""IFOffLineResetDone"" name=""IFOffLineResetDone"" value=""1"" />")
            Else
                Send("<input type=""hidden"" id=""tester1"" name=""tester1"" />")
            End If
        End If
%>
<script type="text/javascript">
  function LoadDocument(url) { 
    location = url; 
  }
  
  var divElems = new Array("validationbody");
  var divVals  = new Array(256);
  var divImages = new Array("imgValidation");
  var boxesValue = <% Sendb(BoxesValue) %>;
  
  function updateCookie() {
    updateBoxesCookie(divElems, divVals);
  }
  
  function collapseBoxes() {
    updatePageBoxes(divElems, divVals, divImages, boxesValue);
  }
  
  function showDiv(elemName) {
    var elem = document.getElementById(elemName);
    
    if (elem != null) {
      elem.style.display = (elem.style.display == "none") ? "block" : "none";
    }
  }
  
  function setComponentsColor(color) {
    var elem = document.getElementById("linkComponent");
    
    if (elem != null) {
      elem.style.color = color;
    }
  }
  
  function launchScReport(locID) {
    openPopup('sanity-check-rpt.aspx?loc=' + locID);
  }

</script>
<form action="store-detail.aspx" id="mainform" name="mainform" method="post">
<input type="hidden" id="EngineID" name="EngineID" value="<%Sendb(EngineType)%>" />
<div id="intro">
    <h1 id="title">
        <%
            If LocationID = 0 Then
                If LocationTypeID = 2 Then
                    Sendb(Copient.PhraseLib.Lookup("term.newserver", LanguageID))
                Else
                    Sendb(Copient.PhraseLib.Lookup("term.newstore", LanguageID))
                End If
            Else
                If LocationTypeID = 2 Then
                    Sendb(Copient.PhraseLib.Lookup("term.server", LanguageID) & " " & ExtLocationCode & ": ")
                Else
                    Sendb(Copient.PhraseLib.Lookup("term.store", LanguageID) & " " & ExtLocationCode & ": ")
                End If
                Sendb(MyCommon.TruncateString(LocationName, 40))
            End If
        %>
    </h1>
    <div id="controls">
        <%
            If (Logix.UserRoles.CRUDStoresAndTerminals = True) Then
                If ((EngineType = 2) Or (EngineType = 9)) Then
                    Send_Save()
                End If
            End If
            Send(" <input type=""hidden"" id=""LocationID"" name=""LocationID"" value=""" & LocationID & """ />")
            If MyCommon.Fetch_SystemOption(75) Then
                If (LocationID > 0 And Logix.UserRoles.AccessNotes) Then
                    Send_NotesButton(13, LocationID, AdminUserID)
                End If
            End If
        %>
    </div>
</div>
<div id="main">
    <%
        bWaitingForIPL = False
        MyCommon.QueryStr = "select LS.MustIPL, LS.IncentiveFetchOffline, L.CMOADeployStatus " & _
                            "from Locations L with (NoLock) left join LocalServers LS with (NoLock) on LS.LocationID = L.LocationID " & _
                            "where L.Deleted=0 and L.LocationID=" & LocationID
        rst = MyCommon.LRT_Select()
        If (rst.Rows.Count > 0) Then
            If MyCommon.NZ(rst.Rows(0).Item("MustIPL"), False) Then
                If LocationTypeID = 2 Then
                    infoMessage = Copient.PhraseLib.Lookup("server-edit.mustipl", LanguageID)
                Else
                    infoMessage = Copient.PhraseLib.Lookup("store-edit.mustipl", LanguageID)
                End If
                bWaitingForIPL = True
            ElseIf MyCommon.NZ(rst.Rows(0).Item("IncentiveFetchOffline"), True) Then
                infoMessage = Copient.PhraseLib.Lookup("store-edit.incentivefetchoffline", LanguageID)
            ElseIf MyCommon.NZ(rst.Rows(0).Item("CMOADeployStatus"), 0) = -1 And EngineType <= 1 Then
                infoMessage = Copient.PhraseLib.Lookup("status.warning", LanguageID)
            End If
        End If
      
        If (infoMessage <> "") Then
            Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
        End If
        If (statusMessage <> "") Then
            Send("<div id=""statusbar"" class=""green-background"">" & statusMessage & "</div>")
        End If
      
        If LocationID = 0 Then
            Send("</div>")
            Send("</form>")
            GoTo done
        End If
    %>
    <div id="column1">
        <div class="box" id="identification">
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.identification", LanguageID))%>
                </span>
            </h2>
            <%
                If LocationTypeID = 2 Then
                    Sendb("<b>" & Copient.PhraseLib.Lookup("term.server", LanguageID) & " ")
                Else
                    Sendb("<b>" & Copient.PhraseLib.Lookup("term.store", LanguageID) & " ")
                End If
                Send(ExtLocationCode & "</b><br />")
                If (ExtLocationCode = LocationName) Or ((Copient.PhraseLib.Lookup("term.store", LanguageID) & " " & ExtLocationCode) = LocationName) Then
                Else
                    Send(MyCommon.SplitNonSpacedString(LocationName, 25) & "<br />")
                End If
                If Address1 <> "" Then
                    Sendb(MyCommon.SplitNonSpacedString(Address1, 25))
                End If
                If (Address1 <> "") And (Address2 <> "") Then
                    Sendb(", ")
                End If
                If Address2 <> "" Then
                    Sendb(MyCommon.SplitNonSpacedString(Address2, 25) & "<br />")
                End If
                If (City = "") And (State = "") Then
                Else
                    Send(MyCommon.SplitNonSpacedString(City, 25) & ", " & MyCommon.SplitNonSpacedString(State, 25) & "<br />")
                End If
                MyCommon.QueryStr = "select CountryID, PhraseID from Countries with (NoLock) where CountryID=" & CountryID
                rst = MyCommon.LRT_Select()
                For Each row In rst.Rows
                    Sendb(Copient.PhraseLib.Lookup(row.Item("PhraseID"), LanguageID) & "<br />")
                Next
                Send("<br class=""half"" />")
                If EnginePhraseID > 0 Then
                    Send(Copient.PhraseLib.Lookup(EnginePhraseID, LanguageID))
                Else
                    Send(EngineName)
                End If
                Send(" " & Copient.PhraseLib.Lookup("term.promotionengine", LanguageID) & "<br />")
                Send("<br class=""half"" />")
                If (TestingLocation) Then
                    Send(Copient.PhraseLib.Lookup("term.testinglocation", LanguageID) & "<br />")
                    Send("<br class=""half"" />")
                End If
                Sendb("<b><a href=""store-edit.aspx?LocationID=" & LocationID & """>")
                If LocationTypeID = 2 Then
                    Sendb(Copient.PhraseLib.Lookup("term.serverconfiguration", LanguageID))
                Else
                    Sendb(Copient.PhraseLib.Lookup("term.storeconfiguration", LanguageID))
                End If
                Send("</a></b><br />")
            %>
            <span style="line-height: 0.1;">&nbsp;</span><br />
            <hr class="hidden" />
        </div>
        <% MyCommon.QueryStr = "select LocalServerID,IncentiveFetchURL from LocalServers with (NoLock) where LocationID=" & LocationID & ";"
            rst = MyCommon.LRT_Select
            Dim LocalServerID As String = "&nbsp;"
            If (rst.Rows.Count > 0) Then
                LocalServerID = MyCommon.NZ(rst.Rows(0).Item("LocalServerID"), "&nbsp;")
                IncentiveFetchURL = MyCommon.NZ(rst.Rows(0).Item("IncentiveFetchURL"), "")
            End If%>
        <div class="box" id="store" <% if ((enginetype = 2) or (enginetype = 9)) then sendb(" style=""display:none;""")%>>
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.storestatus", LanguageID))%>
                </span>
            </h2>
            <%
                lastHeard = Copient.PhraseLib.Lookup("term.notfound", LanguageID)
                lastIPL = Copient.PhraseLib.Lookup("term.notfound", LanguageID)
                alertStatus = Copient.PhraseLib.Lookup("term.notfound", LanguageID)
                reportStatus = Copient.PhraseLib.Lookup("term.notfound", LanguageID)
                Dim iplTable As String
                MyCommon.QueryStr = "select LOC.GenIPL, LOC.LastIPL, LS.CMLastHeard, LS.LocalServerID, LOC.SendAlert, LOC.HealthReported " & _
                                    "from Locations as LOC with (NoLock) " & _
                                    "left join LocalServers as LS with (NoLock) " & _
                                    "on LOC.LocationID=LS.LocationID " & _
                                    "where (LOC.LocationID=" & LocationID & ");"
                rst = MyCommon.LRT_Select()
                If (rst.Rows.Count > 0) Then
                    row = rst.Rows(0)
                    If (Not IsDBNull(row.Item("CMLastHeard"))) Then
                        lastHeard = Logix.ToShortDateTimeString(row.Item("CMLastHeard"), MyCommon)
                    Else
                        lastHeard = Copient.PhraseLib.Lookup("term.never", LanguageID)
                    End If
                    If (Not IsDBNull(row.Item("LastIPL"))) Then
                        iplTable = Logix.ToShortDateTimeString(row.Item("LastIPL"), MyCommon)
                    Else
                        iplTable = Copient.PhraseLib.Lookup("term.never", LanguageID)
                    End If
                    If (row.Item("GenIPL")) Then
                        iplTable = iplTable & " (" & Copient.PhraseLib.Lookup("term.inprogress", LanguageID) & ")"
                    End If
                    If (MyCommon.NZ(row.Item("SendAlert"), False)) Then
                        alertStatus = Copient.PhraseLib.Lookup("term.on", LanguageID)
                    Else
                        alertStatus = Copient.PhraseLib.Lookup("term.off", LanguageID)
                    End If
                    If (MyCommon.NZ(row.Item("HealthReported"), False)) Then
                        reportStatus = Copient.PhraseLib.Lookup("term.on", LanguageID)
                    Else
                        reportStatus = Copient.PhraseLib.Lookup("term.off", LanguageID)
                    End If
                End If
            %>
            <table cellpadding="0" cellspacing="0" summary="<% Sendb(Copient.PhraseLib.Lookup("term.status", LanguageID))%>">
                <tr>
                    <td style="width: 33%;">
                        <% Sendb(Copient.PhraseLib.Lookup("term.lastheard", LanguageID) & ":")%>
                    </td>
                    <td>
                        <% Sendb(lastHeard)%>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                        <% Sendb(Copient.PhraseLib.Lookup("term.lastipl", LanguageID) & ":")%>
                    </td>
                    <td>
                        <% Sendb(iplTable)%>
                    </td>
                    <td class="green">
                    </td>
                </tr>
                <tr>
                    <td>
                        <% Sendb(Copient.PhraseLib.Lookup("term.alert", LanguageID) & ":")%>
                    </td>
                    <td>
                        <% Sendb(alertStatus)%>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                        <% Sendb(Copient.PhraseLib.Lookup("term.report", LanguageID) & ":")%>
                    </td>
                    <td>
                        <% Sendb(reportStatus)%>
                    </td>
                    <td>
                    </td>
                </tr>
                <tr>
                    <td>
                        <% Sendb(Copient.PhraseLib.Lookup("term.cmconnector", LanguageID) & ":")%>
                    </td>
                    <td>
                        <% Sendb("<a href=""log-view.aspx?filetype=2&localserverid=" & rst.Rows(0).Item("LocalServerID") & """ target=""_blank"">" & Copient.PhraseLib.Lookup("term.log", LanguageID) & "</a>")%>
                    </td>
                    <td>
                    </td>
                </tr>
            </table>
            <hr class="hidden" />
        </div>

        <div class="box" id="localserver"<% if (enginetype <> 2) then sendb(" style=""display:none;""")%>>
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.localserver", LanguageID))%>
          </span>
        </h2>
        <%
          MyCommon.QueryStr = "select LocalServerID, ImageFetchURL, IncentiveFetchURL, PhoneHomeIPOverride, OfflineFTPUser, OfflineFTPPass, OfflineFTPPath, OfflineFTPIP " & _
                              "from LocalServers with (NoLock) where LocationID=@LocationID;"
          MyCommon.DBParameters.Add("@LocationID", SqlDbType.BigInt).Value = LocationID
          rst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT) 'MyCommon.LRT_Select
          If (rst.Rows.Count > 0) Then
            ImageFetchURL = MyCommon.NZ(rst.Rows(0).Item("ImageFetchURL"), "")
            IncentiveFetchURL = MyCommon.NZ(rst.Rows(0).Item("IncentiveFetchURL"), "")
            PhoneHomeIPOverride = MyCommon.NZ(rst.Rows(0).Item("PhoneHomeIPOverride"), "")
            OfflineFTPUser = MyCommon.NZ(rst.Rows(0).Item("OfflineFTPUser"), "")
            OfflineFTPPass = MyCommon.NZ(rst.Rows(0).Item("OfflineFTPPass"), "")
            OfflineFTPPath = MyCommon.NZ(rst.Rows(0).Item("OfflineFTPPath"), "")
            OfflineFTPIP = MyCommon.NZ(rst.Rows(0).Item("OfflineFTPIP"), "")
            Send("<label for=""imagefetchurl"">" & Copient.PhraseLib.Lookup("term.imagefetchurl", LanguageID) & ":</label><br />")
            Send("<input type=""text"" id=""imagefetchurl"" name=""imagefetchurl"" class=""longer"" maxlength=""255"" value=""" & ImageFetchURL & """ /><br />")
            Send("<label for=""incentivefetchurl"">" & Copient.PhraseLib.Lookup("term.incentivefetchurl", LanguageID) & ":</label><br />")
            Send("<input type=""text"" id=""incentivefetchurl"" name=""incentivefetchurl"" class=""longer"" maxlength=""255"" value=""" & IncentiveFetchURL & """ /><br />")
            Send("<label for=""PhoneHomeIPOverride"">" & Copient.PhraseLib.Lookup("term.phonehomeipoverride", LanguageID) & ":</label><br />")
            Send("<input type=""text"" id=""PhoneHomeIPOverride"" name=""PhoneHomeIPOverride"" class=""longer"" maxlength=""255"" value=""" & PhoneHomeIPOverride & """ /><br />")
            'OfflineFTP
            Send("<label for=""FTPUser"">" & Copient.PhraseLib.Lookup("term.OfflineFTPUsername", LanguageID) & ":</label><br />")
            Send("<input type=""text"" id=""FTPUser"" name=""FTPUser"" class=""longer"" maxlength=""255"" value=""" & OfflineFTPUser & """ /><br />")
            Send("<label for=""FTPPass"">" & Copient.PhraseLib.Lookup("term.OfflineFTPPassword", LanguageID) & ":</label><br />")
            Send("<input type=""password"" id=""FTPPass"" name=""FTPPass"" class=""longer"" maxlength=""255"" value=""" & OfflineFTPPass & """ /><br />")
            Send("<label for=""FTPPath"">" & Copient.PhraseLib.Lookup("term.OfflineFTPPath", LanguageID) & ":</label><br />")
            Send("<input type=""text"" id=""FTPPath"" name=""FTPPath"" class=""longer"" maxlength=""255"" value=""" & OfflineFTPPath & """ /><br />")
            Send("<label for=""FTPIP"">" & Copient.PhraseLib.Lookup("term.OfflineFTPIP", LanguageID) & ":</label><br />")
            Send("<input type=""text"" id=""FTPIP"" name=""FTPIP"" class=""longer"" maxlength=""255"" value=""" & OfflineFTPIP & """ /><br />")
          End If
          Send("<br class=""half"" />")
          Sendb(Copient.PhraseLib.Lookup("sanitycheck.status", LanguageID) & ": ")
          MyCommon.QueryStr = "select DBOK from SanityCheckStatus with (NoLock) where LocationID=" & LocationID
          rst = MyCommon.LRT_Select
          If (rst.Rows.Count > 0) Then
            SanityCheckPassed = rst.Rows(0).Item("DBOK")
            If (SanityCheckPassed) Then
              Send(Copient.PhraseLib.Lookup("term.passed", LanguageID))
            Else
              Sendb(Copient.PhraseLib.Lookup("term.failed", LanguageID))
              Sendb("<a href=""javascript:launchScReport(" & LocationID & ");"" alt=""" & Copient.PhraseLib.Lookup("store-detail.click-to-view", LanguageID) & _
                   """ title=""" & Copient.PhraseLib.Lookup("store-detail.click-to-view", LanguageID) & """ style=""margin-left:7px;"">")
              Send("<img src=""../images/info.png"" border=""0"" style=""vertical-align: bottom;"" /></a>")
            End If
          Else
            Send(Copient.PhraseLib.Lookup("term.noresults", LanguageID))
          End If
          
        %>
        <hr class="hidden" />
      </div>


        <%
            MyCommon.QueryStr = "select LocalServerID, LocalServers.LocationID, " & _
                                "IncentiveLastHeard, IncentiveFetchOffline, TransactionLastHeard, TransDownloadLastHeard, TransactionWaitingACK, " & _
                                "MovementLastHeard, InfoNowLastHeard, ImageFetchURL, IncentiveFetchURL, PhoneHomeIPOverride, LastHeard, " & _
                                "LastWarningTime, LastMUWarning, RecordCreated, MustIPL, MacAddress, FailoverServer, LastIP, LastLocationID, " & _
                                "SanityCheckLastHeard, TUMD5, CMLastHeard,FailOverStart from LocalServers with (NoLock) " & _
                                "left outer join ( Select top 1 FailOverStart,LocationID from FailOverHistory where    LocationID =" & LocationID & "  order by FailOverHistoryID desc) as  b  " & _
                                "on LocalServers.LocationID =b.LocationID where LocalServers.LocationID=" & LocationID & ";"
            rst = MyCommon.LRT_Select
        
            If ((EngineType <> 2) And (EngineType <> 9)) Or (rst.Rows.Count = 0) Then
                GoTo skippast
            End If
        %>
        <div class="box" id="server">
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.server", LanguageID))%>
                </span>
            </h2>
            <table cellpadding="0" cellspacing="0" summary="<% Sendb(Copient.PhraseLib.Lookup("term.server", LanguageID))%>">
                <tr>
                    <td style="width: 33%;">
                        <b>
                            <% Sendb(Copient.PhraseLib.Lookup("term.localserver", LanguageID) & Copient.PhraseLib.Lookup("term.id", LanguageID) & ":")%>
                        </b>
                    </td>
                    <td>
                        <b>
                            <%
                               
                                Sendb(LocalServerID)%>
                        </b>
                    </td>
                </tr>
                <tr>
                    <td style="width: 33%;">
                        <b>
                            <% Sendb(Copient.PhraseLib.Lookup("term.location", LanguageID) & Copient.PhraseLib.Lookup("term.id", LanguageID) & ":")%>
                        </b>
                    </td>
                    <td>
                        <b>
                            <% Sendb(LocationID)%>
                        </b>
                    </td>
                </tr>
                <tr>
                    <td style="width: 33%;">
                        <% Sendb(Copient.PhraseLib.Lookup("term.failoverserver", LanguageID) & ":")%>
                    </td>
                    <td>
                        <%
                            If MyCommon.NZ(rst.Rows(0).Item("FailoverServer"), False) Then
                                Sendb(Copient.PhraseLib.Lookup("term.true", LanguageID))
                            Else
                                Sendb(Copient.PhraseLib.Lookup("term.false", LanguageID))
                            End If
                        %>
                    </td>
                </tr>
                <tr>
                    <td>
                        <% Sendb(Copient.PhraseLib.Lookup("term.ipaddress", LanguageID) & ":")%>
                    </td>
                    <td>
                        <% Sendb(MyCommon.NZ(rst.Rows(0).Item("LastIP"), Copient.PhraseLib.Lookup("term.na", LanguageID)))%>
                    </td>
                </tr>
                <tr>
                    <td>
                        <% Sendb(Copient.PhraseLib.Lookup("term.macaddress", LanguageID) & ":")%>
                    </td>
                    <td>
                        <% Sendb(MyCommon.NZ(rst.Rows(0).Item("MacAddress"), Copient.PhraseLib.Lookup("term.na", LanguageID)))%>
                    </td>
                </tr>
            </table>
            <hr class="hidden" />
        </div>
        <div class="box" id="communications">
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.communication", LanguageID))%>
                </span>
            </h2>
            <table cellpadding="0" cellspacing="0" summary="<% Sendb(Copient.PhraseLib.Lookup("term.communication", LanguageID))%>">
                <tr>
                    <td>
                        <% Sendb(Copient.PhraseLib.Lookup("store-edit.lastcommunication", LanguageID) & ":")%>
                    </td>
                    <td>
                        <% 
                            If IsDBNull(rst.Rows(0).Item("LastHeard")) Then
                                Sendb("&nbsp;")
                            Else
                                Sendb(Logix.ToShortDateTimeString(rst.Rows(0).Item("LastHeard"), MyCommon))
                            End If
                        %>
                    </td>
                </tr>
                <tr>
                    <td>
                        <% Sendb(Copient.PhraseLib.Lookup("store-edit.lastlookup", LanguageID) & ":")%>
                    </td>
                    <td>
                        <%
                            If IsDBNull(rst.Rows(0).Item("InfoNowLastHeard")) Then
                                Sendb("&nbsp;")
                            Else
                                Sendb(Logix.ToShortDateTimeString(rst.Rows(0).Item("InfoNowLastHeard"), MyCommon))
                            End If
                            If ((MyCommon.NZ(rst.Rows(0).Item("InfoNowLastHeard"), Now)) = Now) Then
                                Sendb("-")
                            Else
                                LogFileType = -5 'CPE-PhoneHomeLog
                                If EngineType = 9 Then LogFileType = 209 'UE-GetCusomterInfoLog
                                Sendb("&nbsp;(<a href=""log-view.aspx?filetype=" & LogFileType & "&localserverid=" & rst.Rows(0).Item("LocalServerID") & """ target=""_blank"">" & Copient.PhraseLib.Lookup("term.log", LanguageID) & "</a>)")
                            End If
                        %>
                    </td>
                </tr>
                <tr>
                    <td>
                        <% Sendb(Copient.PhraseLib.Lookup("store-edit.lasttu", LanguageID) & ":")%>
                    </td>
                    <td>
                        <%
                            If IsDBNull(rst.Rows(0).Item("TransactionLastHeard")) Then
                                Sendb("&nbsp;")
                            Else
                                Sendb(Logix.ToShortDateTimeString(rst.Rows(0).Item("TransactionLastHeard"), MyCommon))
                            End If
                            If ((MyCommon.NZ(rst.Rows(0).Item("TransactionLastHeard"), Now)) = Now) Then
                                Sendb("-")
                            Else
                                LogFileType = -4 'CPE-TransUpdateLog
                                If EngineType = 9 Then
                                    If LocationTypeID = 2 And EngineType = 9 Then
                                        LogFileType = 108 'UE running at enterprise. MessageReceiverAgent log should be shown.
                                    Else
                                        LogFileType = 201 'UE-TransUpdateLog
                                    End If
                                End If
                                Sendb("&nbsp;(<a href=""log-view.aspx?filetype=" & LogFileType & "&localserverid=" & rst.Rows(0).Item("LocalServerID") & """ target=""_blank"">" & Copient.PhraseLib.Lookup("term.log", LanguageID) & "</a>)")
                            End If
                        %>
                    </td>
                </tr>
                <tr>
                    <td>
                        <% Sendb(Copient.PhraseLib.Lookup("store-edit.lasttd", LanguageID) & ":")%>
                    </td>
                    <td>
                        <%
                            If IsDBNull(rst.Rows(0).Item("TransDownloadLastHeard")) Then
                                Sendb("&nbsp;")
                            Else
                                Sendb(Logix.ToShortDateTimeString(rst.Rows(0).Item("TransDownloadLastHeard"), MyCommon))
                            End If
                            If ((MyCommon.NZ(rst.Rows(0).Item("TransDownloadLastHeard"), Now)) = Now) Then
                                Sendb("-")
                            Else
                                LogFileType = -4 'CPE-TransUpdateLog
                                If EngineType = 9 Then
                                    If LocationTypeID = 2 And EngineType = 9 Then
                                        LogFileType = 106 'UE running at enterprise. MessageSenderAgent log should be shown.
                                    Else
                                        LogFileType = 201 'UE-TransUpdateLog
                                    End If
                                End If
                                Sendb("&nbsp;(<a href=""log-view.aspx?filetype=" & LogFileType & "&localserverid=" & rst.Rows(0).Item("LocalServerID") & """ target=""_blank"">" & Copient.PhraseLib.Lookup("term.log", LanguageID) & "</a>)")
                            End If
                        %>
                    </td>
                </tr>
                <tr>
                    <td>
                        <% Sendb(Copient.PhraseLib.Lookup("store-edit.lastofferupdate", LanguageID) & ":")%>
                    </td>
                    <td>
                        <%
                            If (MyCommon.NZ(rst.Rows(0).Item("IncentiveFetchOffline"), 0) > 0) Then
                                Sendb("<span class=""red"">" & Copient.PhraseLib.Lookup("store-health.IncentiveFetchOffline", LanguageID) & "</span>")
                                Sendb("&nbsp;(<a href=""?resetIncentiveFetchOffline=1&LocationID=" & LocationID & """>" & Copient.PhraseLib.Lookup("term.reset", LanguageID) & "</a>)")
                            Else
                                If IsDBNull(rst.Rows(0).Item("IncentiveLastHeard")) Then
                                    Sendb("&nbsp;")
                                Else
                                    Sendb(Logix.ToShortDateTimeString(rst.Rows(0).Item("IncentiveLastHeard"), MyCommon))
                                End If
                                If ((MyCommon.NZ(rst.Rows(0).Item("IncentiveLastHeard"), Now)) = Now) Then
                                    Sendb("-")
                                Else
                                    LogFileType = -3 'CPE-IncentiveFetchLog
                                    If EngineType = 9 Then LogFileType = 200 'UE-IncentiveFetchLog
                                    Sendb("&nbsp;(<a href=""log-view.aspx?filetype=" & LogFileType & "&localserverid=" & rst.Rows(0).Item("LocalServerID") & """ target=""_blank"">" & Copient.PhraseLib.Lookup("term.log", LanguageID) & "</a>)")
                                End If
                            End If
                        %>
                    </td>
                </tr>
                <%If (LocationTypeID = 1 And EngineType = 9) Then%>
                <tr>
                    <td>
                        <% Sendb(Copient.PhraseLib.Lookup("store-edit.lastsanitycheck", LanguageID) & ":")%>
                    </td>
                    <td>
                        <%
                            If IsDBNull(rst.Rows(0).Item("SanityCheckLastHeard")) Then
                                Sendb("&nbsp;")
                            Else
                                Sendb(Logix.ToShortDateTimeString(rst.Rows(0).Item("SanityCheckLastHeard"), MyCommon))
                            End If
                            If ((MyCommon.NZ(rst.Rows(0).Item("SanityCheckLastHeard"), Now)) = Now) Then
                                Sendb("-")
                            Else
                                LogFileType = -7 'CPE-SanityCheckLog
                                If EngineType = 9 Then LogFileType = 205 'UE-SanityCheckLog
                                Sendb("&nbsp;(<a href=""log-view.aspx?filetype=" & LogFileType & "&localserverid=" & rst.Rows(0).Item("LocalServerID") & """ target=""_blank"">" & Copient.PhraseLib.Lookup("term.log", LanguageID) & "</a>)")
                            End If
                        %>
                    </td>
                </tr>
                <tr>
                    <td>
                        <% Sendb(Copient.PhraseLib.Lookup("store-edit.lastfailover", LanguageID) & ":")%>
                    </td>
                    <td>
                        <%
                            If IsDBNull(rst.Rows(0).Item("FailOverStart")) Then
                                Sendb("-")
                            Else
                                Sendb(Logix.ToShortDateTimeString(rst.Rows(0).Item("FailOverStart"), MyCommon))
                            End If
                        %>
                    </td>
                </tr>
                <%End If%>
                <tr>
                    <td>
                        <% Sendb(Copient.PhraseLib.Lookup("store-edit.lastipl", LanguageID) & ":")%>
                    </td>
                    <td>
                        <%
                            MyCommon.QueryStr = "select CMOADeployRpt, CMOADeployStatus, CMOARptDate, CMOADeploySuccessDate, LastIPL " & _
                                                "from Locations with (NoLock) where LocationID=" & LocationID & ";"
                            rst = MyCommon.LRT_Select
                            If rst.Rows.Count > 0 Then
                                If ((EngineType = 2) Or (EngineType = 9)) Then
                                    Sendb(MyCommon.NZ(rst.Rows(0).Item("LastIPL"), Copient.PhraseLib.Lookup("term.never", LanguageID)))
                                Else
                                    If (Not IsDBNull(rst.Rows(0).Item("CMOARptDate"))) Then
                                        deployDate = Logix.ToShortDateTimeString(rst.Rows(0).Item("CMOARptDate"), MyCommon)
                                    Else
                                        deployDate = Copient.PhraseLib.Lookup("term.never", LanguageID)
                                    End If
                                    Sendb(deployDate)
                                    Send("<br class=""half"" />")
                                    If LocationID = 0 Then
                                    Else
                                        If ((EngineType <> 2) And (EngineType <> 9)) Then
                                            Send_GenerateIPL()
                                        End If
                                    End If
                                End If
                            End If
                        %>
                    </td>
                </tr>
                <tr>
                    <td colspan="2">
                        <input type="checkbox" id="sendalert" name="sendalert" value="true" <% if SendAlert then sendb(" checked=""checked""") %> />
                        <label for="sendalert">
                            <% Sendb(Copient.PhraseLib.Lookup("store-edit.alert", LanguageID))%>
                        </label>
                    </td>
                </tr>
                <tr>
                    <td colspan="2">
                        <input type="checkbox" id="reportenabled" name="reportenabled" value="true" <% if ReportEnabled then sendb(" checked=""checked""") %> />
                        <label for="reportenabled">
                            <% Sendb(Copient.PhraseLib.Lookup("store-edit.report", LanguageID))%>
                        </label>
                    </td>
                </tr>
            </table>
            <hr class="hidden" />
        </div>
        <% skippast:%>
        <% If ((EngineType <> 2) And (EngineType <> 9)) Then%>
        <div class="box" id="deployment">
            <h2>
                <span>
                    <%Sendb(Copient.PhraseLib.Lookup("store-edit.iplstatus", LanguageID))%>
                </span>
            </h2>
            <h3>
                <%Sendb(Copient.PhraseLib.Lookup("term.lastattempted", LanguageID) & ":")%>
            </h3>
            <%
                MyCommon.QueryStr = "select CMOADeployRpt, CMOADeployStatus, CMOARptDate, CMOADeploySuccessDate from Locations with (NoLock) " & _
                                    "where LocationID=" & LocationID & ";"
                rst = MyCommon.LRT_Select
                For Each row In rst.Rows
                    If (Not IsDBNull(row.Item("CMOARptDate"))) Then
                        deployDate = Logix.ToShortDateTimeString(row.Item("CMOARptDate"), MyCommon)
                    Else
                        deployDate = Copient.PhraseLib.Lookup("term.never", LanguageID)
                    End If
                    Send(deployDate & "<br />")
                Next
            %>
            <br class="half" />
            <h3>
                <%Sendb(Copient.PhraseLib.Lookup("term.lastsuccessful", LanguageID) & ":")%>
            </h3>
            <%
                For Each row In rst.Rows
                    If (Not IsDBNull(row.Item("CMOADeploySuccessDate"))) Then
                        deployDate = Logix.ToShortDateTimeString(row.Item("CMOADeploySuccessDate"), MyCommon)
                    Else
                        deployDate = Copient.PhraseLib.Lookup("term.never", LanguageID)
                    End If
                    Send(deployDate & "<br />")
                Next
            %>
            <br class="half" />
            <h3>
                <%Sendb(Copient.PhraseLib.Lookup("term.laststatus", LanguageID) & ":")%>
            </h3>
            <%
                For Each row In rst.Rows
                    Send(MyCommon.NZ(row.Item("CMOADeployRpt"), ""))
                Next
            %>
            <br />
            <span style="line-height: 0.1;">&nbsp;</span><br />
            <hr class="hidden" />
        </div>
        <div class="box" id="ipl">
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.generateipl", LanguageID))%>
                </span>
            </h2>
            <% 
                MyCommon.QueryStr = "select * from PromoEngines with (NoLock) where Installed=1 and EngineID=" & EngineType
                rst = MyCommon.LRT_Select
                If (rst.Rows.Count > 0) Then
                    If (Logix.UserRoles.CRUDStoresAndTerminals = True) Then
                        If (LocationID <> 0) Then
                            Send("<br class=""half"" />")
                            Send_GenerateIPL()
                        End If
                    End If
                End If
            %>
            <hr class="hidden" />
        </div>
        <% End If%>
    </div>
    <div id="gutter">
    </div>
    <div id="column2">
        <%
            If ((EngineType = 2) Or (EngineType = 9)) Then
                CommsFilter = " (CASE WHEN DATEADD(n, " & CentralLowValue & ", LS.LastHeard) >= getDate() and DATEADD(n," & CentralLowValue & ", LS.IncentiveLastHeard) >=getDate() and DATEADD(n," & CentralLowValue & ",LS.TransactionLastHeard) >=getDate() and DATEADD(n," & CentralLowValue & ", LS.TransDownLoadLastHeard) >=getDate() and DBOK=1 THEN  1 ELSE 0 END) "
                SearchClause &= " and (DBOK = 0 or isnull(" & CommsFilter & ",0) = 0 or (ls.Sev1Errors>0 or ls.Sev10Errors>0)) and HealthReported = 1 and ls.LocalServerID is not null "

                MyCommon.QueryStr = "select loc.LocationID, ls.LocalServerID, loc.LocationName, loc.EngineID, loc.ExtLocationCode, loc.HealthReported, PE.Description as EngineName, ls.CMLastHeard, ls.LastHeard, ls.FailoverServer, ls.MustIPL, ls.SanityCheckLastHeard, " & _
                                    "IsNull(loc.SendAlert,0) SendAlert, scs.DBOK as SanityCheckResult, " & CommsFilter & " as Comms, " & _
                                    "ls.LastRunID, ls.Sev1Errors, ls.Sev10Errors, ls.LastHeard, ls.IncentiveLastHeard, ls.TransactionLastHeard, ls.TransDownloadLastHeard, DBOK, scs.LastReportDate, " & _
                                    "case when ls.Sev1Errors > 0 or DATEADD(n, " & CentralHighValue & ", LS.LastHeard) <= getDate() or DATEADD(n," & CentralHighValue & ", LS.IncentiveLastHeard) <=getDate() or DATEADD(n," & CentralHighValue & ",LS.TransactionLastHeard) <=getDate() " & _
                                    "  or DATEADD(n," & CentralHighValue & ", LS.TransDownLoadLastHeard) <=getDate() then 1 " & _
                                    "     when DATEADD(n," & CentralMediumValue & ", LS.LastHeard) <= getDate() or DATEADD(n," & CentralMediumValue & ", LS.IncentiveLastHeard) <=getDate() or DATEADD(n," & CentralMediumValue & ",LS.TransactionLastHeard) <=getDate() " & _
                                    "  or DATEADD(n," & CentralMediumValue & ", LS.TransDownLoadLastHeard) <=getDate() or (DBOK=0 and DATEADD(n," & CentralMediumValue & ", scs.LastReportDate) <= getDate()) then 5 " & _
                                    "     when ls.Sev10Errors > 0 or DATEADD(n," & CentralLowValue & ", LS.LastHeard) <= getDate() or DATEADD(n," & CentralLowValue & ", LS.IncentiveLastHeard) <=getDate() or DATEADD(n," & CentralLowValue & ",LS.TransactionLastHeard) <=getDate() " & _
                                    "  or DATEADD(n," & CentralLowValue & ", LS.TransDownLoadLastHeard) <=getDate() or (DBOK=0 and DATEADD(n," & CentralLowValue & ", scs.LastReportDate) <= getDate()) then 10 else 0 end as Severity " & _
                                    "from Locations as loc with (nolock) " & _
                                    "left join PromoEngines PE with (NoLock) on PE.EngineID=loc.EngineID " & _
                                    "left join LocalServers as ls with (NoLock) on loc.LocationID = ls.LocationID " & _
                                    "left join SanityCheckStatus scs with (NoLock) on loc.LocationID = scs.LocationID " & _
                                    "where loc.EngineID in (2,9) and loc.Deleted = 0 and ls.LocationID=" & LocationID & SearchClause
                rst2 = MyCommon.LRT_Select
                If (rst2.Rows.Count > 0) Then
                    ' does this location wish to have errors reported and if so, are there any errors to report 
                    If (MyCommon.NZ(rst2.Rows(0).Item("HealthReported"), False) AndAlso MyCommon.NZ(rst2.Rows(0).Item("Severity"), 0) > 0) Then
                        Send("<div class=""box"" id=""warnings"">")
                        Send("  <h2>")
                        Send("    <span class=""white"">")
                        Send("     " & Copient.PhraseLib.Lookup("term.warnings", LanguageID))
                        Send("    </span>")
                        Send("  </h2>")
              
                        ' list all errors for this location
                        MyCommon.QueryStr = "select 'term.local' as ServerType, 1 as ServerTypeID, HE.HealthSeverityID, 'LS' + CONVERT(nvarchar(6),HE.ErrorID) as ErrorCode, " & _
                                            "HE.ErrorID, 0 as MinutesInError, HE.ErrorText, HT.TagName, HS.SectionName from LS_HealthErrors HE with (NoLock) " & _
                                            "left join HealthTags HT with (NoLock) on HT.TagID = HE.TagID " & _
                                            "left join HealthSections HS with (NoLock) on HS.SectionID = HE.SectionID " & _
                                            "where LocalServerID=" & MyCommon.NZ(rst2.Rows(0).Item("LocalServerID"), -1) & _
                                            " and RunID=" & MyCommon.NZ(rst2.Rows(0).Item("LastRunID"), -1) & " order by HE.HealthSeverityID;"
                        dst2 = MyCommon.LWH_Select
              
                        MyCommon.QueryStr = "dbo.pa_StoreHealth_CentralCommErrors"
                        MyCommon.Open_LRTsp()
                        MyCommon.LRTsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = MyCommon.NZ(rst2.Rows(0).Item("LocalServerID"), -1)
                        MyCommon.LRTsp.Parameters.Add("@High", SqlDbType.Int).Value = CentralHighValue
                        MyCommon.LRTsp.Parameters.Add("@Medium", SqlDbType.Int).Value = CentralMediumValue
                        MyCommon.LRTsp.Parameters.Add("@Low", SqlDbType.Int).Value = CentralLowValue
                        dst3 = MyCommon.LRTsp_select
                        MyCommon.Close_LRTsp()
                        dst2.Merge(dst3)
              
                        Send("  <table style=""width:95%"" summary=""" & Copient.PhraseLib.Lookup("term.errors", LanguageID) & """>")
                        Send("    <tr>")
                        Send("      <th>" & Copient.PhraseLib.Lookup("term.severity", LanguageID) & "</th>")
                        Send("      <th>" & Copient.PhraseLib.Lookup("term.from", LanguageID) & "</th>")
                        Send("      <th>" & Copient.PhraseLib.Lookup("term.id", LanguageID) & "</th>")
                        Send("      <th style=""width:180px;"">" & Copient.PhraseLib.Lookup("term.duration", LanguageID) & "</th>")
                        Send("    </tr>")
                        For Each row2 In dst2.Rows
                            If (MyCommon.NZ(row2.Item("ServerType"), "term.central") = "term.local") Then
                                ' this is a local server error so we need to find out the duration of the error
                                MyCommon.QueryStr = "dbo.pa_StoreHealth_ErrorDuration"
                                MyCommon.Open_LWHsp()
                                MyCommon.LWHsp.Parameters.Add("@LocalServerID", SqlDbType.Int).Value = MyCommon.NZ(rst2.Rows(0).Item("LocalServerID"), -1)
                                MyCommon.LWHsp.Parameters.Add("@ErrorID", SqlDbType.Int).Value = MyCommon.NZ(row2.Item("ErrorID"), -1)
                                MyCommon.LWHsp.Parameters.Add("@MinutesInError", SqlDbType.Int).Direction = ParameterDirection.Output
                                MyCommon.LWHsp.ExecuteNonQuery()
                                MinutesInError = MyCommon.LWHsp.Parameters("@MinutesInError").Value
                                MyCommon.Close_LWHsp()
                                row2.Item("MinutesInError") = MinutesInError
                            End If
                        Next
              
                        rows = dst2.Select("", "HealthSeverityID asc, MinutesInError desc")
                        RowCt = rows.Length
                        Counter = 1
                        For Each row2 In rows
                            MinutesInError = MyCommon.NZ(row2.Item("MinutesInError"), 0)
                            Send("    <tr>")
                            Send("      <td>" & GetSeverityText(MyCommon.NZ(row2.Item("HealthSeverityID"), -1), SeverityTypes) & "</td>")
                            Send("      <td>" & Copient.PhraseLib.Lookup(MyCommon.NZ(row2.Item("ServerType"), ""), LanguageID) & "</td>")
                            QryStr = "?SrvType=" & MyCommon.NZ(row2.Item("ServerTypeID"), "2") & "&Err=" & MyCommon.NZ(row2.Item("ErrorID"), 0)
                            Sendb("      <td><a href=""javascript:openPopup('health-resolutions.aspx" & QryStr & "');"" ")
                            Sendb(" title=""" & Copient.PhraseLib.Lookup("store-health.resolution-note", LanguageID) & """")
                            Send(">" & MyCommon.NZ(row2.Item("ErrorCode"), "&nbsp;") & "</a></td>")
                            Send("      <td>" & GetDurationText(MinutesInError) & "</td>")
                            Send("    </tr>")
                            If (Counter < RowCt) Then
                                Send("    <tr>")
                                Send("      <td colspan=""4"" style=""background-color: #cccccc;height: 1px;padding: 0;margin: 0;""></td>")
                                Send("    </tr>")
                            End If
                            Counter += 1
                        Next
                        Send("  </table>")
                        Send("  &nbsp;<br />")
                        Send("</div>")
                    End If
                End If
            End If
        %>
        <div class="box" id="files">
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.files", LanguageID))%>
                </span>
            </h2>
            <% Sendb(Copient.PhraseLib.Lookup("store-edit.files", LanguageID) & ":")%>
            <br />
            <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.files", LanguageID))%>" style="table-layout: fixed;">
                <thead>
                    <tr>
                        <th class="th-name" scope="col">
                            <% Sendb(Copient.PhraseLib.Lookup("term.file", LanguageID))%>
                        </th>
                        <th class="th-age" scope="col">
                            <% Sendb(Copient.PhraseLib.Lookup("term.age", LanguageID))%>
                        </th>
                        <th class="th-datetime" scope="col" style="width: 118px;">
                            <% Sendb(Copient.PhraseLib.Lookup("term.created", LanguageID))%>
                        </th>
                    </tr>
                </thead>
                <tbody>
                    <%
                        Dim dtFiles As DataTable
                        Dim rowFile As DataRow
                        Dim hours, minutes As Integer
                        Dim sName As String
                        If (EngineType = 2) Then
                            MyCommon.QueryStr = "select FileName, datediff(mi, CreationDate, getDate()) as Minutes, convert(varchar, creationdate, 101)+' '+convert(varchar, creationdate, 108) as CreationDate " & _
                                                "from CPE_IncentiveDLBuffer dlb with (NoLock) " & _
                                                "inner join LocalServers ls with (NoLock)on ls.LocalServerID =  dlb.LocalServerID " & _
                                                "where ls.LocationID = " & LocationID & " and WaitingACK<2 order by Minutes desc;"
                        ElseIf (EngineType = 9) Then
                            MyCommon.QueryStr = "select FileName, datediff(mi, CreationDate, getDate()) as Minutes, convert(varchar, creationdate, 101)+' '+convert(varchar, creationdate, 108) as CreationDate " & _
                                                "from UE_IncentiveDLBuffer dlb with (NoLock) " & _
                                                "inner join LocalServers ls with (NoLock)on ls.LocalServerID =  dlb.LocalServerID " & _
                                                "where ls.LocationID = " & LocationID & " and WaitingACK<2 order by Minutes desc;"
                        Else
                            MyCommon.QueryStr = "select FileTypeId, LinkId, datediff(mi, CreationDate, getDate()) as Minutes, convert(varchar, creationdate, 101)+' '+convert(varchar, creationdate, 108) as CreationDate " & _
                                                "from OfferDLBuffer dlb with (NoLock) " & _
                                                "where dlb.LocationID = " & LocationID & " and dlb.StatusFlag in (0,1) order by Minutes desc;"
                        End If
                        dtFiles = MyCommon.LRT_Select
                        If dtFiles.Rows.Count = 0 Then
                            Send("<tr>")
                            Send("    <td colspan=""3""><i>" & Copient.PhraseLib.Lookup("term.none", LanguageID) & "</i></td>")
                            Send("</tr>")
                        Else
                            If Not (Right(IncentiveFetchURL, 1) = "/") Then
                                IncentiveFetchURL = IncentiveFetchURL & "/"
                            End If
                            For Each rowFile In dtFiles.Rows
                                minutes = MyCommon.NZ(rowFile.Item("Minutes"), 0)
                                hours = minutes \ 60
                                minutes = minutes Mod 60
                                Send("<tr>")
                                If ((EngineType = 2) Or (EngineType = 9)) Then
                                    Send("    <td style=""overflow:hidden;""><a href=""" & IncentiveFetchURL & MyCommon.NZ(rowFile.Item("FileName"), "") & """ target=""_blank"">" & MyCommon.NZ(rowFile.Item("FileName"), "") & "</a></td>")
                                Else
                                    Select Case MyCommon.NZ(rowFile.Item("FileTypeId"), 0)
                                        Case 10, 11, 12
                                            sName = Copient.PhraseLib.Lookup("term.Offer", LanguageID) & ": " & MyCommon.NZ(rowFile.Item("LinkId"), 0)
                                        Case 20, 22, 23
                                            sName = Copient.PhraseLib.Lookup("term.ProductGroup", LanguageID) & ": " & MyCommon.NZ(rowFile.Item("LinkId"), 0)
                                        Case 30, 32, 33
                                            sName = Copient.PhraseLib.Lookup("term.CustomerGroup", LanguageID) & ": " & MyCommon.NZ(rowFile.Item("LinkId"), 0)
                                        Case 40, 41
                                            sName = Copient.PhraseLib.Lookup("term.Ipl", LanguageID)
                                        Case Else
                                            sName = "Unknown file type!"
                                    End Select
                                    Select Case MyCommon.NZ(rowFile.Item("FileTypeId"), 0)
                                        Case 11, 12, 22, 32
                                            sName += " (" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & ")"
                                        Case 23, 33
                                            sName += " (" & Copient.PhraseLib.Lookup("term.manual", LanguageID) & ")"
                                        Case 40
                                            sName += " (" & Copient.PhraseLib.Lookup("term.offers", LanguageID) & ")"
                                        Case 41
                                            sName += " (" & Copient.PhraseLib.Lookup("term.ProductGroups", LanguageID) & ")"
                                    End Select
                                    Send("    <td style=""overflow:hidden;""><a target=""_blank"">" & sName & "</a></td>")
                                End If
                                Send("    <td>" & hours.ToString().PadLeft(2, "0") & ":" & minutes.ToString().PadLeft(2, "0") & "</td>")
                                Send("    <td>" & MyCommon.NZ(rowFile.Item("CreationDate"), "") & "</td>")
                                Send("</tr>")
                            Next
                        End If
                    %>
                </tbody>
            </table>
            <hr class="hidden" />
        </div>
        <% If ((EngineType <> 2) And (EngineType <> 9)) Then%>
        <div class="box" id="validation">
            <h2 style="float: left;">
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.validationreport", LanguageID))%>
                </span>
            </h2>
            <% Send_BoxResizer("validationbody", "imgValidation", Copient.PhraseLib.Lookup("term.validationreport", LanguageID), True)%>
            <div id="validationbody">
                <%
                    Dim dtValid As DataTable
                    Dim rowOK(), rowWaiting(), rowWatches(), rowWarnings() As DataRow
                    Dim objTemp1, objTemp2 As Object
                    Dim GraceHours As Integer
                    Dim GraceHoursWarn As Integer
                    Dim iTotalCount As Integer
                    Dim EnginePrefix As String = "CM"
            
                    EnginePrefix = "CM"
                    objTemp1 = MyCommon.Fetch_CM_SystemOption(10)
                    objTemp2 = MyCommon.Fetch_CM_SystemOption(11)
            
                    objTemp1 = MyCommon.Fetch_CM_SystemOption(10)
                    If Not (Integer.TryParse(objTemp1.ToString, GraceHours)) Then
                        GraceHours = 4
                    End If
            
                    objTemp2 = MyCommon.Fetch_CM_SystemOption(11)
                    If Not (Integer.TryParse(objTemp2.ToString, GraceHoursWarn)) Then
                        GraceHoursWarn = 24
                    End If
            
                    ' Display Offer validation
                    MyCommon.QueryStr = "dbo.pa_" & EnginePrefix & "_ValidationReport_LocOffers"
            
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
                    MyCommon.LRTsp.Parameters.Add("@GraceHours", SqlDbType.Int).Value = GraceHours
                    MyCommon.LRTsp.Parameters.Add("@GraceHoursWarn", SqlDbType.Int).Value = GraceHoursWarn
            
                    dtValid = MyCommon.LRTsp_select()
                    iTotalCount = dtValid.Rows.Count
                    rowOK = dtValid.Select("Status=0", "Name")
                    rowWaiting = dtValid.Select("Status=1", "Name")
                    rowWatches = dtValid.Select("Status=2", "Name")
                    rowWarnings = dtValid.Select("Status=3", "Name")
                    MyCommon.Close_LRTsp()
            
                    ValidateOfferColor = IIf(rowWarnings.Length > 0, "red", "green")
            
                    Send("<a href=""javascript:showDiv('divOffer');"" style=""color:" & ValidateOfferColor & ";""><b>+ Offers</b><br /></a>")
                    Send("<div id=""divOffer"" style=""margin-left:10px;display:none;"">")
                    Send("<a id=""validLinkOffer" & LocationID & """ href=""javascript:openPopup('" & EnginePrefix & "-validation-report-loc.aspx?type=in&id=" & LocationID & "&level=0&gh=" & GraceHours & "&ghw=" & GraceHoursWarn & "');"">")
                    Send(Copient.PhraseLib.Lookup("term.valid", LanguageID) & " (" & rowOK.Length & " of " & iTotalCount & ")</a><br />")
                    Send("<a id=""waitingLinkOffer" & LocationID & """ href=""javascript:openPopup('" & EnginePrefix & "-validation-report-loc.aspx?type=in&id=" & LocationID & "&level=1&gh=" & GraceHours & "&ghw=" & GraceHoursWarn & "');"">")
                    Send(Copient.PhraseLib.Lookup("term.waiting", LanguageID) & " (" & rowWaiting.Length & " of " & iTotalCount & ")</a><br />")
                    Send("<a id=""watchLinkOffer" & LocationID & """ href=""javascript:openPopup('" & EnginePrefix & "-validation-report-loc.aspx?type=in&id=" & LocationID & "&level=2&gh=" & GraceHours & "&ghw=" & GraceHoursWarn & "');"">")
                    Send(Copient.PhraseLib.Lookup("term.watches", LanguageID) & " (" & rowWatches.Length & " of " & iTotalCount & ")</a><br />")
                    Send("<a id=""warningLinkOffer" & LocationID & """ href=""javascript:openPopup('" & EnginePrefix & "-validation-report-loc.aspx?type=in&id=" & LocationID & "&level=3&gh=" & GraceHours & "&ghw=" & GraceHoursWarn & "');"">")
                    Send(Copient.PhraseLib.Lookup("term.warnings", LanguageID) & " (" & rowWarnings.Length & " of " & iTotalCount & ")</a><br /></div>")
            
                    ' Display Product Group validation
                    MyCommon.QueryStr = "dbo.pa_CM_ValidationReport_LocProdGroups"
                    MyCommon.Open_LRTsp()
                    MyCommon.LRTsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
                    MyCommon.LRTsp.Parameters.Add("@GraceHours", SqlDbType.Int).Value = GraceHours
                    MyCommon.LRTsp.Parameters.Add("@GraceHoursWarn", SqlDbType.Int).Value = GraceHoursWarn
            
                    dtValid = MyCommon.LRTsp_select()
                    iTotalCount = dtValid.Rows.Count
            
                    rowOK = dtValid.Select("Status=0", "Name")
                    rowWaiting = dtValid.Select("Status=1", "Name")
                    rowWatches = dtValid.Select("Status=2", "Name")
                    rowWarnings = dtValid.Select("Status=3", "Name")
                    MyCommon.Close_LRTsp()
            
                    ValidateProdGroupColor = IIf(rowWarnings.Length > 0, "red", "green")
            
                    Send("<a href=""javascript:showDiv('divProdGroups');"" style=""color:" & ValidateProdGroupColor & ";""><b>+ Product Groups</b><br /></a>")
                    Send("<div id=""divProdGroups"" style=""margin-left:10px;display:none;"">")
                    Send("<a id=""validLinkPG" & LocationID & """ href=""javascript:openPopup('" & EnginePrefix & "-validation-report-loc.aspx?type=pg&id=" & LocationID & "&level=0&gh=" & GraceHours & "&ghw=" & GraceHoursWarn & "');"">")
                    Send(Copient.PhraseLib.Lookup("term.valid", LanguageID) & " (" & rowOK.Length & " of " & iTotalCount & ")</a><br />")
                    Send("<a id=""waitingLinkPG" & LocationID & """ href=""javascript:openPopup('" & EnginePrefix & "-validation-report-loc.aspx?type=pg&id=" & LocationID & "&level=1&gh=" & GraceHours & "&ghw=" & GraceHoursWarn & "');"">")
                    Send(Copient.PhraseLib.Lookup("term.waiting", LanguageID) & " (" & rowWaiting.Length & " of " & iTotalCount & ")</a><br />")
                    Send("<a id=""watchLinkPG" & LocationID & """ href=""javascript:openPopup('" & EnginePrefix & "-validation-report-loc.aspx?type=pg&id=" & LocationID & "&level=2&gh=" & GraceHours & "&ghw=" & GraceHoursWarn & "');"">")
                    Send(Copient.PhraseLib.Lookup("term.watches", LanguageID) & " (" & rowWatches.Length & " of " & iTotalCount & ")</a><br />")
                    Send("<a id=""warningLinkPG" & LocationID & """ href=""javascript:openPopup('" & EnginePrefix & "-validation-report-loc.aspx?type=pg&id=" & LocationID & "&level=3&gh=" & GraceHours & "&ghw=" & GraceHoursWarn & "');"">")
                    Send(Copient.PhraseLib.Lookup("term.warnings", LanguageID) & " (" & rowWarnings.Length & " of " & iTotalCount & ")</a><br /></div>")
            
                    If EngineType = 1 Then
                        ' Display Customer Group validation
                        MyCommon.QueryStr = "dbo.pa_CM_ValidationReport_LocCustGroups"
                        MyCommon.Open_LRTsp()
                        MyCommon.LRTsp.Parameters.Add("@LocationID", SqlDbType.Int).Value = LocationID
                        MyCommon.LRTsp.Parameters.Add("@GraceHours", SqlDbType.Int).Value = GraceHours
                        MyCommon.LRTsp.Parameters.Add("@GraceHoursWarn", SqlDbType.Int).Value = GraceHoursWarn
              
                        dtValid = MyCommon.LRTsp_select()
                        iTotalCount = dtValid.Rows.Count
              
                        rowOK = dtValid.Select("Status=0", "Name")
                        rowWaiting = dtValid.Select("Status=1", "Name")
                        rowWatches = dtValid.Select("Status=2", "Name")
                        rowWarnings = dtValid.Select("Status=3", "Name")
                        MyCommon.Close_LRTsp()
              
                        ValidateCustGroupColor = IIf(rowWarnings.Length > 0, "red", "green")
              
                        Send("<a href=""javascript:showDiv('divCustGroups');"" style=""color:" & ValidateCustGroupColor & ";""><b>+ " & Copient.PhraseLib.Lookup("term.customergroups", LanguageID) & "</b><br /></a>")
                        Send("<div id=""divCustGroups"" style=""margin-left:10px;display:none;"">")
                        Send("<a id=""validLinkCG" & LocationID & """ href=""javascript:openPopup('CM-validation-report-loc.aspx?type=cg&id=" & LocationID & "&level=0&gh=" & GraceHours & "&ghw=" & GraceHoursWarn & "');"">")
                        Send(Copient.PhraseLib.Lookup("term.valid", LanguageID) & " (" & rowOK.Length & " of " & iTotalCount & ")</a><br />")
                        Send("<a id=""waitingLinkCG" & LocationID & """ href=""javascript:openPopup('CM-validation-report-loc.aspx?type=cg&id=" & LocationID & "&level=1&gh=" & GraceHours & "&ghw=" & GraceHoursWarn & "');"">")
                        Send(Copient.PhraseLib.Lookup("term.waiting", LanguageID) & " (" & rowWaiting.Length & " of " & iTotalCount & ")</a><br />")
                        Send("<a id=""watchLinkCG" & LocationID & """ href=""javascript:openPopup('CM-validation-report-loc.aspx?type=cg&id=" & LocationID & "&level=2&gh=" & GraceHours & "&ghw=" & GraceHoursWarn & "');"">")
                        Send(Copient.PhraseLib.Lookup("term.watches", LanguageID) & " (" & rowWatches.Length & " of " & iTotalCount & ")</a><br />")
                        Send("<a id=""warningLinkCG" & LocationID & """ href=""javascript:openPopup('CM-validation-report-loc.aspx?type=cg&id=" & LocationID & "&level=3&gh=" & GraceHours & "&ghw=" & GraceHoursWarn & "');"">")
                        Send(Copient.PhraseLib.Lookup("term.warnings", LanguageID) & " (" & rowWarnings.Length & " of " & iTotalCount & ")</a><br /></div>")
                    End If
                %>
            </div>
            <hr class="hidden" />
        </div>
        <% End If%>
    </div>
    <br clear="all" />
</div>
</form>
<script runat="server">
    Public Class SeverityEntry
        Public Description As String
        Public PhraseID As Integer
        Sub New(ByVal SevDesc As String, ByVal SevPhrase As Integer)
            Description = SevDesc
            PhraseID = SevPhrase
        End Sub
    End Class
  
    Function GetDurationText(ByVal MinutesInError As Integer) As String
        Dim DurationText As New StringBuilder()
        Dim DurSpan As New TimeSpan(MinutesInError * TimeSpan.TicksPerMinute)
        Dim ErrDays, ErrHours, ErrMinutes As Integer
    
        ErrDays = DurSpan.Days
        ErrHours = DurSpan.Hours
        ErrMinutes = DurSpan.Minutes
    
        If (ErrDays > 0) Then DurationText.Append(ErrDays & " " & Copient.PhraseLib.Lookup("term.days", LanguageID) & ", ")
        If (ErrHours > 0 OrElse ErrDays > 0) Then DurationText.Append(ErrHours & " " & Copient.PhraseLib.Lookup("term.hours", LanguageID) & ", ")
        If (ErrMinutes > 0 OrElse ErrHours > 0 OrElse ErrDays > 0) Then DurationText.Append(ErrMinutes & " " & Copient.PhraseLib.Lookup("term.minutes", LanguageID))
    
        Return DurationText.ToString
    End Function
  
    Function GetSeverityText(ByVal SeverityID As Integer, ByVal SeverityTypes As Hashtable) As String
        Dim SeverityText As String = ""
        Dim Severity As SeverityEntry
    
        Severity = SeverityTypes.Item("Sev" & SeverityID)
        If (Severity IsNot Nothing) Then
            SeverityText = Copient.PhraseLib.Lookup(Severity.PhraseID, LanguageID, Severity.Description)
        Else
            SeverityText = Copient.PhraseLib.Lookup("term.unknown", LanguageID)
        End If
    
        Return SeverityText
    End Function
</script>
<%
    If MyCommon.Fetch_SystemOption(75) Then
        If (LocationID > 0 And Logix.UserRoles.AccessNotes) Then
            Send_Notes(13, LocationID, AdminUserID)
        End If
    End If
done:
Finally
    MyCommon.Close_LogixRT()
    MyCommon.Close_LogixWH()
End Try
Send_BodyEnd("mainform")
MyCommon = Nothing
Logix = Nothing
%>
