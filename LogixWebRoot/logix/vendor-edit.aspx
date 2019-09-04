<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.CommonInc" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%
  ' *****************************************************************************
  ' * FILENAME: vendor-edit.aspx 
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
  Dim VendorID As Long = -1
  Dim VendorName As String
  Dim VendorDesc As String
  Dim ExternalID As String
  Dim LastUpdate As String
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  Dim dst As DataTable
  Dim rst As DataTable
  Dim rst2 As DataTable
  Dim rstAssociated As DataTable = Nothing
  Dim row As DataRow
  Dim GName As String = ""
  Dim bSave As Boolean
  Dim bDelete As Boolean
  Dim bCreate As Boolean
  Dim bClose As Boolean
  Dim VendorNameTitle As String = ""
  Dim ShowActionButton As Boolean = False
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim HasAssociatedOffers As Boolean = False
  Dim Chargeable As Boolean = False
  Dim OfferID As Integer = 0
  Dim EngineID As Integer = -1
  Dim CreatedFromOffer As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "vendor-edit.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  Try
    ' fill in if it was a get method
    If Request.RequestType = "GET" Then
      VendorID = IIf(Request.QueryString("VendorID") = "", -1, MyCommon.Extract_Val(Request.QueryString("VendorID")))
      VendorName = Logix.TrimAll(Request.QueryString("VendorName"))
      VendorDesc = Logix.TrimAll(Request.QueryString("VendorDesc"))
            ExternalID = Logix.TrimAll(Request.QueryString("ExtID"))
      Chargeable = IIf(Request.QueryString("chargeable") = "1", True, False)
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
      If Request.QueryString("close") = "" Then
        bClose = False
      Else
        bClose = True
      End If
    Else
      VendorID = IIf(Request.Form("VendorID") = "", -1, MyCommon.Extract_Val(Request.Form("VendorID")))
      If VendorID <= 0 Then
        VendorID = IIf(Request.QueryString("VendorID") = "", -1, MyCommon.Extract_Val(Request.QueryString("VendorID")))
      End If
      VendorName = Logix.TrimAll(Request.Form("VendorName"))
      VendorDesc = Logix.TrimAll(Request.Form("VendorDesc"))
      ExternalID = Logix.TrimAll(Request.Form("ExtID"))
      Chargeable = IIf(Request.Form("chargeable") = "1", True, False)
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
      If Request.Form("close") = "" Then
        bClose = False
      Else
        bClose = True
      End If
    End If
    
    OfferID = MyCommon.Extract_Val(Request.QueryString("OfferID"))
    EngineID = IIf(Request.QueryString("EngineID") = "", -1, MyCommon.Extract_Val(Request.QueryString("EngineID")))
    CreatedFromOffer = OfferID > 0 And EngineID > 0
    If CreatedFromOffer And VendorID = 0 Then
      GName = Copient.PhraseLib.Lookup("term.offer", LanguageID) & " " & OfferID & " " & StrConv(Copient.PhraseLib.Lookup("term.group", LanguageID), VbStrConv.Lowercase)
      MyCommon.QueryStr = "select count(*) as GroupCount from ProductGroups where Name like '" & GName & "%';"
      rst = MyCommon.LRT_Select
      If rst.Rows(0).Item("GroupCount") > 0 Then
        GName = GName & " (" & rst.Rows(0).Item("GroupCount") & ")"
      End If
    End If
    
    
    Send_HeadBegin("term.vendor", , VendorID)
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send_Scripts()
    Send_HeadEnd()
    If CreatedFromOffer Then
      Send_BodyBegin(3)
    Else
      Send_BodyBegin(1)
      Send_Bar(Handheld)
      Send_Help(CopientFileName)
      Send_Logos()
      Send_Tabs(Logix, 8)
      Send_Subtabs(Logix, 8, 4, , VendorID)
    End If
    
    If (Logix.UserRoles.EditSystemConfiguration = False) Then
      Send_Denied(1, "perm.admin-configuration")
      GoTo done
    End If
    
    If (Request.QueryString("new") <> "") Then
      Response.Redirect("vendor-edit.aspx")
    End If
    
    If (VendorID > 0) Then
      MyCommon.QueryStr = "select I.IncentiveID as OfferID, I.IncentiveName as OfferName, I.EndDate as ProdEndDate,buy.ExternalBuyerId from CPE_Incentives I " &
                            "Left Outer Join Buyers buy on buy.BuyerID = I.BuyerID " & 
                            "where I.ChargebackVendorID=" & VendorID & " and I.Deleted=0;"
      rstAssociated = MyCommon.LRT_Select
      HasAssociatedOffers = (rstAssociated.Rows.Count > 0)
    End If
    
    If bSave OrElse (CreatedFromOffer AndAlso VendorID = 0) Then
      If (VendorName = "") Or (ExternalID = "") Then
        infoMessage = Copient.PhraseLib.Lookup("vendors.noname", LanguageID)
      Else
        If (VendorID = -1) Then
          MyCommon.QueryStr = "SELECT VendorID FROM Vendors with (NoLock) WHERE Name = '" & MyCommon.Parse_Quotes(VendorName) & "'"
          dst = MyCommon.LRT_Select
          If (dst.Rows.Count > 0) Then
            infoMessage = Copient.PhraseLib.Lookup("vendors.nameused", LanguageID)
          Else
                        MyCommon.QueryStr = "SELECT VendorID FROM Vendors with (NoLock) WHERE ExtVendorID = '" & MyCommon.Parse_Quotes(ExternalID) & "'"
            dst = MyCommon.LRT_Select
            If (dst.Rows.Count > 0) Then
              infoMessage = Copient.PhraseLib.Lookup("vendors.codeused", LanguageID)
            Else
              MyCommon.QueryStr = "dbo.pt_Vendor_Insert"
              MyCommon.Open_LRTsp()
              MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 100).Value = VendorName
                            MyCommon.LRTsp.Parameters.Add("@ExtVendorID", SqlDbType.NVarChar, 50).Value = ExternalID
                            MyCommon.LRTsp.Parameters.Add("@Description", SqlDbType.NVarChar, 1000).Value = VendorDesc
                            MyCommon.LRTsp.Parameters.Add("@Chargeable", SqlDbType.Bit).Value = IIf(Chargeable, 1, 0)
                            MyCommon.LRTsp.Parameters.Add("@VendorID", SqlDbType.Int).Direction = ParameterDirection.Output
                            MyCommon.LRTsp.ExecuteNonQuery()
                            VendorID = MyCommon.LRTsp.Parameters("@VendorID").Value
                            MyCommon.Activity_Log(37, VendorID, AdminUserID, Copient.PhraseLib.Lookup("history.vendor-create", LanguageID))
                        End If
                    End If
                Else
                    ' update the existing department
                    MyCommon.QueryStr = "SELECT VendorID FROM Vendors with (NoLock) WHERE Name = '" & MyCommon.Parse_Quotes(VendorName) & "' and VendorID <> " & VendorID
                    dst = MyCommon.LRT_Select
                    If (dst.Rows.Count > 0) Then
                        infoMessage = Copient.PhraseLib.Lookup("vendors.nameused", LanguageID)
                    Else
                        MyCommon.QueryStr = "SELECT VendorID FROM Vendors with (NoLock) WHERE ExtVendorID = '" & MyCommon.Parse_Quotes(ExternalID) & "' and VendorID <> " & VendorID
                        dst = MyCommon.LRT_Select
                        If (dst.Rows.Count > 0) Then
                            infoMessage = Copient.PhraseLib.Lookup("vendors.codeused", LanguageID)
                        Else
                          
                            MyCommon.QueryStr = "update Vendors with (RowLock) set Name='" & MyCommon.Parse_Quotes(VendorName) & "', ExtVendorID='" & MyCommon.Parse_Quotes(ExternalID) & "', " & _
                                                "Description='" & MyCommon.Parse_Quotes(VendorDesc) & "', Chargeable=" & IIf(Chargeable, 1, 0) & ", LastUpdate=getdate() " & _
                                                "where VendorID=" & VendorID
                            MyCommon.LRT_Execute()
                            MyCommon.Activity_Log(37, VendorID, AdminUserID, Copient.PhraseLib.Lookup("history.vendor-edit", LanguageID))
                        End If
                    End If
        End If
      End If
      
    ElseIf bDelete OrElse (CreatedFromOffer AndAlso VendorID = 0) Then
      If (VendorID = 0) Then
        infoMessage = Copient.PhraseLib.Lookup("vendors.nodelete", LanguageID)
      Else
        If (HasAssociatedOffers) Then
          infoMessage = Copient.PhraseLib.Lookup("vendors.inuse", LanguageID)
        Else
          MyCommon.QueryStr = "DELETE FROM Vendors with (RowLock) WHERE VendorID = " & VendorID
          MyCommon.LRT_Execute()
          MyCommon.Activity_Log(37, VendorID, AdminUserID, Copient.PhraseLib.Lookup("history.vendor-delete", LanguageID))
          Response.Status = "301 Moved Permanently"
          Response.AddHeader("Location", "vendor-list.aspx")
        End If
      End If
    ElseIf (Request.QueryString("close") <> "") Then
      Response.Status = "301 Moved Permanently"
    End If
    
    LastUpdate = ""
    
    If Not bCreate Then
      ' no one clicked anything
      MyCommon.QueryStr = "select VendorID, ExtVendorID, Name, Description, Chargeable, LastUpdate " & _
                          "from Vendors CV with (NoLock) " & _
                          "where AnyVendor=0 and Deleted=0 and VendorID=" & VendorID
      rst = MyCommon.LRT_Select()
      If (rst.Rows.Count > 0) Then
        VendorID = MyCommon.NZ(rst.Rows(0).Item("VendorID"), -1)
        VendorName = MyCommon.NZ(rst.Rows(0).Item("Name"), "")
        VendorDesc = MyCommon.NZ(rst.Rows(0).Item("Description"), "")
        ExternalID = MyCommon.NZ(rst.Rows(0).Item("ExtVendorID"), "")
        Chargeable = MyCommon.NZ(rst.Rows(0).Item("Chargeable"), False)
        If (IsDBNull(rst.Rows(0).Item("LastUpdate"))) Then
          LastUpdate = Copient.PhraseLib.Lookup("term.unknown", LanguageID)
        Else
          LastUpdate = Logix.ToLongDateTimeString(rst.Rows(0).Item("LastUpdate"), MyCommon)
        End If
      ElseIf (VendorID > 0) Then
        Send("")
        Send("<div id=""intro"">")
        Send("    <h1 id=""title"">" & Copient.PhraseLib.Lookup("term.vendor", LanguageID) & " #" & VendorID & "</h1>")
        Send("</div>")
        Send("<div id=""main"">")
        Send("    <div id=""infobar"" class=""red-background"">")
        Send("        " & Copient.PhraseLib.Lookup("error.deleted", LanguageID))
        Send("    </div>")
        Send("</div>")
        GoTo done
      End If
    End If
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
  <%
    If CreatedFromOffer Then
      Send("<input type=""hidden"" id=""OfferID"" name=""OfferID"" value=""" & OfferID & """ />")
      Send("<input type=""hidden"" id=""EngineID"" name=""EngineID"" value=""" & EngineID & """ />")
      Send("<input type=""hidden"" id=""slct"" name=""slct"" value=""" & Request.QueryString("slct") & """ />")
      Send("<input type=""hidden"" id=""ex"" name=""ex"" value=""" & Request.QueryString("ex") & """ />")
      Send("<input type=""hidden"" id=""condChanged"" name=""condChanged"" value=""" & Request.QueryString("condChanged") & """ />")
    End If
  %>
  <input type="hidden" id="VendorID" name="VendorID" value="<% Sendb(VendorID) %>" />
  <div id="intro">
    <%Sendb("<h1 id=""title"">")
      If VendorID = -1 Then
        Sendb(Copient.PhraseLib.Lookup("term.newvendor", LanguageID))
      Else
        Sendb(Copient.PhraseLib.Lookup("term.vendor", LanguageID) & " #" & VendorID & ": ")
        MyCommon.QueryStr = "SELECT Name FROM Vendors with (NoLock) WHERE VendorID = " & VendorID & ";"
        rst2 = MyCommon.LRT_Select
        If (rst2.Rows.Count > 0) Then
          VendorNameTitle = MyCommon.NZ(rst2.Rows(0).Item("Name"), "")
        End If
        Send(MyCommon.TruncateString(VendorNameTitle, 40))
      End If
      Sendb("</h1>")
    %>
    <div id="controls">
      <%
        If (VendorID = -1) Then
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
              Send_Delete()
            End If
            If (Logix.UserRoles.EditDepartments And Not CreatedFromOffer) Then
              Send_New()
            End If
            If CreatedFromOffer Then
              Send_Close()
            End If
            Send("</div>")
          End If
          If MyCommon.Fetch_SystemOption(75) And Not CreatedFromOffer Then
            If (Logix.UserRoles.AccessNotes) Then
              Send_NotesButton(35, VendorID, AdminUserID)
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
        <%
          Send("<label id=""lblExtID"" for=""ExtID"">" & Copient.PhraseLib.Lookup("term.code", LanguageID) & ":</label><br style=""line-height: 0.1;"" />")
            Send("<input type=""text"" class=""" & IIf(CreatedFromOffer, "long", "longest") & """ id=""ExtID"" name=""ExtID"" maxlength=""50"" value=""" & ExternalID.Replace("""", "&quot;") & """" & IIf(HasAssociatedOffers, " readonly style=""color:gray;"" ", "") & " />")
          Send("<br class=""half"" />")
          Send("<label for=""VendorName"">" & Copient.PhraseLib.Lookup("term.name", LanguageID) & ":</label><br />")
          If (VendorName Is Nothing) Then
            VendorName = ""
          End If
          Send("<input type=""text"" class=""" & IIf(CreatedFromOffer, "long", "longest") & """ id=""VendorName"" name=""VendorName"" maxlength=""100"" value=""" & VendorName.Replace("""", "&quot;") & """ />")
          Send("<br />")
          Send("<label for=""VendorDesc"">" & Copient.PhraseLib.Lookup("term.description", LanguageID) & "</label><br />")
          If (VendorDesc Is Nothing) Then
            VendorDesc = ""
          End If
            Send("<textarea class""" & IIf(CreatedFromOffer, "long", "longest") & """ id=""VendorDesc"" name=""VendorDesc"" cols=""36"" rows=""3"" oninput=""limitText(this,1000);"">" & VendorDesc & "</textarea><br />")
          Send("<br />")
          Send("<input type=""checkbox"" id=""chargeable"" name=""chargeable"" value=""1""" & IIf(Chargeable, "checked=""checked""", "") & " />")
          Send("<label for=""chargeable"">" & Copient.PhraseLib.Lookup("term.chargeable", LanguageID) & "</label>")
          Send("<br />")
          Send("<br class=""half"" />")
          If (VendorID > -1) Then
            Sendb(Copient.PhraseLib.Lookup("term.edited", LanguageID) & " " & LastUpdate)
          End If
        %>
        <hr class="hidden" />
      </div>
    </div>
    
    <div id="gutter">
    </div>
    
    <div id="column2">
      <% If Not CreatedFromOffer Then %>
      <div class="box" id="offers"<%if(VendorID = -1)then sendb(" style=""visibility: hidden;""") %>>
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.associatedoffers", LanguageID))%>
          </span>
        </h2>
        <div class="boxscroll">
          <% 
              Dim assocName As String=""
            If (VendorID > -1 AndAlso rstAssociated IsNot Nothing) Then
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
      <% End If %>
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
    If (VendorID > 0 And Logix.UserRoles.AccessNotes And Not CreatedFromOffer) Then
      Send_Notes(35, VendorID, AdminUserID)
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
