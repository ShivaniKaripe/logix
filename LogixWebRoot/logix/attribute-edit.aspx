<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.CommonInc" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%
    ' *****************************************************************************
    ' * FILENAME: attribute-edit.aspx 
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
    Dim AttributeTypeID As Long = -1
    Dim ExtID As String = ""
    Dim Description As String = ""
    Dim ValueCount As Integer = 0
    Dim LastUpdate As String
    Dim AttributeValueID As Long = 0
    Dim MyCommon As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
    Dim MyCryptLib As New Copient.CryptLib
    Dim dst As DataTable
    Dim rst As DataTable
    Dim rst2 As DataTable
    Dim rstAssociatedOffers As DataTable = Nothing
    Dim rstAssociatedCustomers As DataTable = Nothing
    Dim row As DataRow
    Dim bSave As Boolean
    Dim bDelete As Boolean
    Dim bCreate As Boolean
    Dim iReadOnlyAttribute As Integer = 0
    Dim infoMessage As String = ""
    Dim Handheld As Boolean = False
    Dim BannerID As Integer = 0
    Dim BannerName As String = ""
    Dim BannersEnabled As Boolean = False
    Dim AllowEditing As Boolean = False
    Dim AllowDelete As Boolean = True
    Dim Shaded As String = ""
    Dim StrTemp As String = ""
  
    Dim TypeInUse As Boolean = False
    Dim ValueInUse As Boolean = False
    Dim TypeInUseOfferCount As Integer = 0
    Dim TypeInUseCustomerCount As Integer = 0
    Dim ValueInUseOfferCount As Integer = 0
    Dim ValueInUseCustomerCount As Integer = 0
  
    Dim NewValue As String = ""
    Dim NewValueExtID As String = ""
    Dim NewValueDesc As String = ""
  
    Dim DeleteValue As String = ""
    Dim SaveValue As String = ""
    Dim SaveValueExtID As String = ""
    Dim SaveValueDesc As String = ""
  
    Dim i As Integer = 0
    Dim MaxEngineSubTypeID As Integer = 0
    Dim SelectedEngineCount As Integer = 0
  
    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If
  
    Response.Expires = 0
    MyCommon.AppName = "attribute-edit.aspx"
    MyCommon.Open_LogixRT()
    MyCommon.Open_LogixXS()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)
    StrTemp = Request.QueryString("ReadOnlyAttribute")
  
    BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
    If (MyCommon.Fetch_CPE_SystemOption(108) = "1") And (Logix.UserRoles.EditAttributes) Then
        AllowEditing = True
    End If
    MyCommon.QueryStr = "select top 1 SubTypeID from PromoEngineSubTypes order by SubTypeID DESC;"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
        MaxEngineSubTypeID = MyCommon.NZ(rst.Rows(0).Item("SubTypeID"), 0)
    End If
  
    If Request.RequestType = "GET" Then
        AttributeTypeID = IIf(Request.QueryString("AttributeTypeID") = "", -1, MyCommon.Extract_Val(Request.QueryString("AttributeTypeID")))
        ExtID = Logix.TrimAll(Request.QueryString("ExtID"))
        Description = Logix.TrimAll(Request.QueryString("Description"))
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
        AttributeTypeID = IIf(Request.Form("AttributeTypeID") = "", -1, MyCommon.Extract_Val(Request.Form("AttributeTypeID")))
        If AttributeTypeID < 0 Then
            AttributeTypeID = IIf(Request.QueryString("AttributeTypeID") = "", -1, MyCommon.Extract_Val(Request.QueryString("AttributeTypeID")))
        End If
        ExtID = Logix.TrimAll(Request.Form("ExtID"))
        Description = Logix.TrimAll(Request.Form("Description"))
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
  
    Send_HeadBegin("term.attribute", , AttributeTypeID)
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
%>
<style type="text/css">
    td
    {
        vertical-align: top;
    }
    .editvalue, .savevalue, .cancelvalue
    {
        font-size: 10px;
        width: 55px;
    }
    .descinput
    {
        font-size: 12px;
        width: 290px;
    }
    * html .descinput
    {
        width: 275px;
    }
    #NewValueExtID
    {
        color: #aaaaaa;
        font-size: 12px;
        width: 75px;
    }
    #NewValueDesc
    {
        color: #aaaaaa;
        font-size: 12px;
        width: 140px;
    }
    #NewValue
    {
        font-size: 10px;
    }
</style>
<%
    Send_Scripts()
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

    function toggleNewValueExtID(action) {
        if (action == "clear") {
            if (document.getElementById("NewValueExtID").value == '<% Sendb(Copient.PhraseLib.Lookup("term.NewExtID", LanguageID)) %>...') {
                document.getElementById("NewValueExtID").value = '';
                document.getElementById("NewValueExtID").style.color = '#000000';
            }
        } else {
            if (document.getElementById("NewValueExtID").value == '') {
                document.getElementById("NewValueExtID").value = '<% Sendb(Copient.PhraseLib.Lookup("term.NewExtID", LanguageID)) %>...';
                document.getElementById("NewValueExtID").style.color = '#aaaaaa';
            }
        }
    }
    function toggleNewValueDesc(action) {
        if (action == "clear") {
            if (document.getElementById("NewValueDesc").value == '<% Sendb(Copient.PhraseLib.Lookup("term.NewDescription", LanguageID)) %>...') {
                document.getElementById("NewValueDesc").value = '';
                document.getElementById("NewValueDesc").style.color = '#000000';
            }
        } else {
            if (document.getElementById("NewValueDesc").value == '') {
                document.getElementById("NewValueDesc").value = '<% Sendb(Copient.PhraseLib.Lookup("term.NewDescription", LanguageID)) %>...';
                document.getElementById("NewValueDesc").style.color = '#aaaaaa';
            }
        }
    }

    function deleteValue(AttributeValueID) {
        if (confirm('<% Sendb(Copient.PhraseLib.Lookup("confirm.deletevalue", LanguageID)) %>')) {
            document.getElementById("DeleteValue").value = AttributeValueID;
            document.mainform.submit();
        } else {
            return false;
        }
    }

    function saveValue(AttributeValueID) {
        document.getElementById("SaveValue").value = guid;
        document.getElementById("SaveValueDesc").value = document.getElementById("descinput-" + AttributeValueID).value;
        document.mainform.submit();
    }

    function newValue() {
        document.getElementById('NewValue').value = '1';

        if (document.getElementById("NewValueExtID").value == '<% Sendb(Copient.PhraseLib.Lookup("term.NewExtID", LanguageID)) %>...') {
            document.getElementById("NewValueExtID").value = '';
        }
        if (document.getElementById("NewValueDesc").value == '<% Sendb(Copient.PhraseLib.Lookup("term.NewDescription", LanguageID)) %>...') {
            document.getElementById("NewValueDesc").value = '';
        }

        document.mainform.submit();
    }

    function clearDefaultButton(buttonGroup) {
        for (i = 0; i < buttonGroup.length; i++) {
            if (buttonGroup[i].checked == true) { // if a button in group is checked,
                buttonGroup[i].checked = false;  // uncheck it
            }
        }

        document.getElementById('selectedRadioID').value = 'clear';
    }

    function toggleDescEdit(AttributeValueID) {
        if (document.getElementById('descedit-' + AttributeValueID).style.display == 'none') {
            document.getElementById('desc-' + AttributeValueID).style.display = 'none';
            document.getElementById('descedit-' + AttributeValueID).style.display = 'block';
        } else {
            document.getElementById('desc-' + AttributeValueID).style.display = 'block';
            document.getElementById('descedit-' + AttributeValueID).style.display = 'none';
        }
    }

    function selectedRadio(newSelectedRadioID) {
        document.getElementById('selectedRadioID').value = newSelectedRadioID;
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
  
    If (Logix.UserRoles.EditSystemConfiguration = False) Then
        Send_Denied(1, "perm.admin-configuration")
        GoTo done
    End If
  
    If (Request.QueryString("new") <> "") Then
        Response.Redirect("attribute-edit.aspx")
    End If
  
    'Determine if there are any associated offers and/or customers, and set the counts and "in-use" booleans accordingly.
    'Offer usage
    MyCommon.QueryStr = "select distinct I.IncentiveID as OfferID, I.IncentiveName as OfferName, I.EndDate as ProdEndDate,buy.ExternalBuyerId as BuyerID from CPE_Incentives as I " & _
                        "inner join CPE_RewardOptions as RO with (NoLock) on RO.IncentiveID=I.IncentiveID and I.Deleted=0 and RO.Deleted=0 " & _
                        "inner join CPE_IncentiveAttributes as IA with (NoLock) on IA.RewardOptionID=RO.RewardOptionID " & _
                        "inner join CPE_IncentiveAttributeTiers as IAT with (NoLock) on IAT.IncentiveAttributeID=IA.IncentiveAttributeID " & _
                          "left outer join Buyers as buy with (nolock) on buy.BuyerId= I.BuyerId " & _
                        "where I.IsTemplate=0 And IAT.AttributeTypeID=" & AttributeTypeID & " " & _
                        "order by OfferName;"
    rstAssociatedOffers = MyCommon.LRT_Select
    TypeInUseOfferCount = rstAssociatedOffers.Rows.Count
    'Customer usage
    MyCommon.QueryStr = "select count(CustomerPK) as Customers from CustomerAttributes as CA with (NoLock) " & _
                        "where AttributeTypeID=" & AttributeTypeID & " and Deleted=0;"
    rst = MyCommon.LXS_Select
    TypeInUseCustomerCount = rst.Rows(0).Item("Customers")
    'Details for the top 100 associated customers
    If TypeInUseCustomerCount > 0 Then
        MyCommon.QueryStr = "select top 100 CA.CustomerPK , CID.ExtCardID, CID.CardTypeID from CustomerAttributes as CA with (NoLock) " & _
                                "left join CardIDs as CID with (NoLock) on CID.CustomerPK = CA.CustomerPK " & _
                                "where AttributeTypeID=" & AttributeTypeID & " and Deleted=0 order by CustomerPK;"
        rstAssociatedCustomers = MyCommon.LXS_Select
        
            Dim rowCount As Integer = 0
            For Each row In rstAssociatedCustomers.Rows
            Dim e_ExtCardID As String = ""
            e_ExtCardID = MyCryptLib.SQL_StringDecrypt(rstAssociatedCustomers.Rows(rowCount)("ExtCardID"))
            If MyCommon.Fetch_SystemOption(144) Then
                'Mask the AltID last four digits
                'CASE WHEN CID.CardTypeID=3 THEN LEFT(CID.ExtCardID,LEN(CID.ExtCardID)-4) ELSE CID.ExtCardID END AS ExtCardID             
                If (CStr(rstAssociatedCustomers.Rows(rowCount)("CardTypeID")) = "3") Then
                    e_ExtCardID = e_ExtCardID.Substring(0, e_ExtCardID.Length - 4)
                End If
            End If
            rstAssociatedCustomers.Rows(rowCount)("ExtCardID") = e_ExtCardID
            rowCount = rowCount+1
            Next
    End If
    If (TypeInUseOfferCount = 0) And (TypeInUseCustomerCount = 0) Then
        TypeInUse = False
    Else
        TypeInUse = True
    End If
  
    If AllowEditing Then
        If bSave Then
            'Save routine for attribute type
            For i = 0 To MaxEngineSubTypeID
                If (Request.QueryString("EngineSubTypeID-" & i) = "1") Then
                    SelectedEngineCount += 1
                End If
            Next
            If (Request.QueryString("ReadOnlyAttribute") = "on") Then
                iReadOnlyAttribute = 1
            End If
            'Currently we allow the default attribute value to change, see BZ 2411
            Dim AllowDefaultChange As Boolean = False
            'Handle a default value change
            If (Request.QueryString("selectedRadioID") <> "" AndAlso Request.QueryString("selectedRadioID") <> "clear") Then
                MyCommon.QueryStr = "update AttributeValues set DefaultValue=N'False'" & _
                                    "where AttributeTypeID=" & AttributeTypeID & ";"
                MyCommon.LRT_Execute()
                MyCommon.QueryStr = "update AttributeValues set DefaultValue=N'True'" & _
                            "where AttributeValueID=" & Request.QueryString("selectedRadioID") & ";"
                MyCommon.LRT_Execute()
                AllowDefaultChange = True
            ElseIf (Request.QueryString("selectedRadioID") = "clear") Then 'Default Value has been cleared
                MyCommon.QueryStr = "update AttributeValues set DefaultValue=N'False'" & _
                                    "where AttributeTypeID=" & AttributeTypeID & ";"
                MyCommon.LRT_Execute()
                AllowDefaultChange = True
            End If
            Description = Logix.TrimAll(Request.QueryString("Description"))
      ExtID = Logix.TrimAll(Request.QueryString("ExtID"))
      
            If (TypeInUse AndAlso (Not AllowDefaultChange)) Then
                infoMessage = Copient.PhraseLib.Lookup("attributes.inuse", LanguageID)
            ElseIf (TypeInUse AndAlso (AllowDefaultChange)) Then
                infoMessage = ""
            ElseIf (Description = "") Then
                infoMessage = Copient.PhraseLib.Lookup("attributes.noname", LanguageID)
            ElseIf (ExtID = "") Then
                infoMessage = Copient.PhraseLib.Lookup("attributes.nocode", LanguageID)
            ElseIf (SelectedEngineCount = 0) Then
                infoMessage = Copient.PhraseLib.Lookup("attributes.noengine", LanguageID)
            ElseIf (Description.IndexOf("<") > -1) OrElse (Description.IndexOf(">") > -1) OrElse (ExtID.IndexOf("<") > -1) OrElse (ExtID.IndexOf(">") > -1) Then
                infoMessage = Copient.PhraseLib.Lookup("attributes.invalidchars", LanguageID)
            Else
                If (AttributeTypeID = -1) Then
                    'New attribute type
                    MyCommon.QueryStr = "select AttributeTypeID from AttributeTypes with (NoLock) where Description=@Description and Deleted=0;"
                    MyCommon.DBParameters.Add("@Description", SqlDbType.NVarChar).Value = Description
          
                    dst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                    If (dst.Rows.Count > 0) Then
                        infoMessage = Copient.PhraseLib.Lookup("attributes.nameused", LanguageID)
                    Else
                        MyCommon.QueryStr = "select AttributeTypeID from AttributeTypes with (NoLock) where ExtID=@ExtID and Deleted=0;"
                        MyCommon.DBParameters.Add("@ExtID",SqlDbType.NVarChar).Value=ExtID
                        dst = MyCommon.ExecuteQuery(Copient.DataBases.LogixRT)
                        If (dst.Rows.Count > 0) Then
                            infoMessage = Copient.PhraseLib.Lookup("attributes.codeused", LanguageID)
                        Else
                            'Insert new attribute
                            MyCommon.QueryStr = "dbo.pt_AttributeType_Insert"
                            MyCommon.Open_LRTsp()
                            MyCommon.LRTsp.Parameters.Add("@Description", SqlDbType.NVarChar, 100).Value = Description
                            MyCommon.LRTsp.Parameters.Add("@ExtID", SqlDbType.NVarChar, 50).Value = ExtID
                            MyCommon.LRTsp.Parameters.Add("@ReadOnlyAttribute", SqlDbType.Bit).Value = iReadOnlyAttribute
                            MyCommon.LRTsp.Parameters.Add("@AttributeTypeID", SqlDbType.Int).Direction = ParameterDirection.Output
                            MyCommon.LRTsp.ExecuteNonQuery()
                            AttributeTypeID = MyCommon.LRTsp.Parameters("@AttributeTypeID").Value
                            'Populate attribute engines table
                            For i = 0 To MaxEngineSubTypeID
                                If (Request.QueryString("EngineSubTypeID-" & i) = "1") Then
                                    MyCommon.QueryStr = "insert into AttributeTypeEngines (AttributeTypeID, EngineID, EngineSubTypeID) " & _
                                                        "values (" & AttributeTypeID & ", 2, " & i & ");"
                                    MyCommon.LRT_Execute()
                                End If
                            Next
                            MyCommon.Activity_Log(43, AttributeTypeID, AdminUserID, Copient.PhraseLib.Lookup("history.attribute-create", LanguageID))
                        End If
                    End If
                Else
                    'Existing attribute type
                    MyCommon.QueryStr = "select AttributeTypeID from AttributeTypes with (NoLock) where Description='" & Description & "' " & _
                                        "and AttributeTypeID<>" & AttributeTypeID & " and Deleted=0;"
                    dst = MyCommon.LRT_Select
                    If (dst.Rows.Count > 0) Then
                        infoMessage = Copient.PhraseLib.Lookup("attributes.nameused", LanguageID)
                    Else
                        MyCommon.QueryStr = "select AttributeTypeID from AttributeTypes with (NoLock) where ExtID='" & ExtID & "' " & _
                                            "and AttributeTypeID<>" & AttributeTypeID & " and Deleted=0;"
                        dst = MyCommon.LRT_Select
                        If (dst.Rows.Count > 0) Then
                            infoMessage = Copient.PhraseLib.Lookup("attributes.codeused", LanguageID)
                        Else
                            'Update attribute
                            MyCommon.QueryStr = "update AttributeTypes set Description=N'" & Description & "', ExtID=N'" & ExtID & "', ReadOnlyAttribute=" & iReadOnlyAttribute & _
                                                "where AttributeTypeID=" & AttributeTypeID & ";"
                            MyCommon.LRT_Execute()
                            'Update attribute engines table
                            MyCommon.QueryStr = "delete from AttributeTypeEngines where AttributeTypeID=" & AttributeTypeID & ";"
                            MyCommon.LRT_Execute()
                            For i = 0 To MaxEngineSubTypeID
                                If (Request.QueryString("EngineSubTypeID-" & i) = "1") Then
                                    MyCommon.QueryStr = "insert into AttributeTypeEngines (AttributeTypeID, EngineID, EngineSubTypeID) " & _
                                                        "values (" & AttributeTypeID & ", 2, " & i & ");"
                                    MyCommon.LRT_Execute()
                                End If
                            Next
                            MyCommon.Activity_Log(43, AttributeTypeID, AdminUserID, Copient.PhraseLib.Lookup("history.attribute-edit", LanguageID))
                        End If
                    End If
                End If
                If infoMessage = "" Then
                    Response.Status = "301 Moved Permanently"
                    Response.AddHeader("Location", "attribute-edit.aspx?AttributeTypeID=" & AttributeTypeID)
                End If
            End If
        ElseIf bDelete Then
            'Delete routine
            If (AttributeTypeID > -1) Then
                If TypeInUse Then
                    infoMessage = Copient.PhraseLib.Lookup("attributes.inuse", LanguageID)
                Else
                    MyCommon.QueryStr = "update AttributeTypes set Deleted=1, LastUpdate=getdate(), CPESendToStore=1 where AttributeTypeID=" & AttributeTypeID & ";"
                    MyCommon.LRT_Execute()
                    MyCommon.QueryStr = "update AttributeValues set Deleted=1, LastUpdate=getdate(), CPESendToStore=1 where AttributeTypeID=" & AttributeTypeID & ";"
                    MyCommon.LRT_Execute()
                    MyCommon.Activity_Log(43, AttributeTypeID, AdminUserID, Copient.PhraseLib.Lookup("history.attribute-delete", LanguageID))
                    Response.Status = "301 Moved Permanently"
                    Response.AddHeader("Location", "attribute-list.aspx")
                End If
            End If
        End If
    
        LastUpdate = ""
    
        'Value editing routines
        If AttributeTypeID > -1 Then
            If (Request.QueryString("NewValue") <> "") Then
                'Adding a new value
                NewValueDesc = MyCommon.Parse_Quotes(Request.QueryString("NewValueDesc")).ToString().Trim()
                NewValueExtID = MyCommon.Parse_Quotes(Request.QueryString("NewValueExtID")).ToString().Trim()
                If (NewValueExtID.Length = 0) Then
                    infoMessage = Copient.PhraseLib.Lookup("inputValidator.blankextidvalue", LanguageID)
                ElseIf (NewValueDesc.Length = 0) Then
                    infoMessage = Copient.PhraseLib.Lookup("inputValidator.blankdescriptionvalue", LanguageID)
                Else
                    MyCommon.QueryStr = "select AttributeValueID from AttributeValues with (NoLock) where AttributeTypeID=" & AttributeTypeID & " " & _
                                    "and (ExtID='" & NewValueExtID & "' or (Description='" & NewValueDesc & "' and Description<>'New description...' )) and Deleted=0;"
                    dst = MyCommon.LRT_Select
                    If dst.Rows.Count > 0 Then
                        infoMessage = "The external ID and description of each value must be unique."
                    ElseIf (NewValueDesc.IndexOf("<") > -1) OrElse (NewValueDesc.IndexOf(">") > -1) OrElse (NewValueExtID.IndexOf("<") > -1) OrElse (NewValueExtID.IndexOf(">") > -1) Then
                        infoMessage = Copient.PhraseLib.Lookup("attributes.invalidchars", LanguageID)
                    Else
                        MyCommon.QueryStr = "insert into AttributeValues (AttributeTypeID, ExtID, Description) " & _
                                            "values (" & AttributeTypeID & ", N'" & NewValueExtID & "', N'" & NewValueDesc & "');"
                        MyCommon.LRT_Execute()
                        MyCommon.Activity_Log(43, AttributeTypeID, AdminUserID, Copient.PhraseLib.Lookup("history.attribute-valueadd", LanguageID))
                        Response.Redirect("/logix/attribute-edit.aspx?AttributeTypeID=" & AttributeTypeID)
                    End If
                End If
            ElseIf (Request.QueryString("SaveValue") <> "") Then
                'Updating an existing value
                SaveValue = Request.QueryString("SaveValue")
                SaveValueDesc = MyCommon.Parse_Quotes(Request.QueryString("SaveValueDesc"))
                SaveValueExtID = MyCommon.Parse_Quotes(Request.QueryString("SaveValueExtID"))
                MyCommon.QueryStr = "select AttributeValueID from AttributeValues with (NoLock) where AttributeValueID=" & SaveValue & ";"
                dst = MyCommon.LRT_Select
                If dst.Rows.Count > 0 Then
                    MyCommon.QueryStr = "select AttributeValueID from AttributeValues with (NoLock) where AttributeTypeID=" & AttributeTypeID & " " & _
                                        "and (Description='" & SaveValueDesc & "' or ExtID='" & SaveValueExtID & "');"
                    dst = MyCommon.LRT_Select
                    If dst.Rows.Count > 0 Then
                        infoMessage = "The external ID and description of each value must be unique."
                    ElseIf (SaveValueDesc.IndexOf("<") > -1) OrElse (SaveValueDesc.IndexOf(">") > -1) OrElse (SaveValueExtID.IndexOf("<") > -1) OrElse (SaveValueExtID.IndexOf(">") > -1) Then
                        infoMessage = Copient.PhraseLib.Lookup("attributes.invalidchars", LanguageID)
                    Else
                        MyCommon.QueryStr = "update AttributeValues set ExtID=N'" & SaveValueExtID & ", Description=N'" & SaveValueDesc & "' where AttributeValueID=" & SaveValue & ";"
                        MyCommon.LRT_Execute()
                        MyCommon.Activity_Log(43, AttributeTypeID, AdminUserID, Copient.PhraseLib.Lookup("history.attribute-valueedit", LanguageID))
                        Response.Redirect("/logix/attribute-edit.aspx?AttributeTypeID=" & AttributeTypeID)
                    End If
                End If
            ElseIf (Request.QueryString("DeleteValue") <> "") Then
                'Deleting a value
                DeleteValue = Request.QueryString("DeleteValue")
                MyCommon.QueryStr = "select AttributeValueID from AttributeValues where AttributeTypeID=" & AttributeTypeID & " and AttributeValueID=" & DeleteValue & " and Deleted=0;"
                dst = MyCommon.LRT_Select
                If (dst.Rows.Count > 0) Then
                    MyCommon.QueryStr = "update AttributeValues set Deleted=1, LastUpdate=getdate(), CPESendToStore=1 " & _
                                        "where AttributeTypeID=" & AttributeTypeID & " and AttributeValueID='" & DeleteValue & "';"
                    MyCommon.LRT_Execute()
                    MyCommon.Activity_Log(43, AttributeTypeID, AdminUserID, Copient.PhraseLib.Lookup("history.attribute-valuedel", LanguageID))
                    Response.Redirect("/logix/attribute-edit.aspx?AttributeTypeID=" & AttributeTypeID)
                End If
            End If
        End If
    End If
  
    If Not bCreate Then
        'No one clicked anything, so load up the attribute type's details
        MyCommon.QueryStr = "select AttributeTypeID, ExtID, Description, Deletable, ReadOnlyAttribute, (select COUNT(*) from AttributeValues where AttributeTypeID=1) as ValueCount " & _
                            "from AttributeTypes as AT with (NoLock) where AttributeTypeID=" & AttributeTypeID & ";"
        rst = MyCommon.LRT_Select()
        If (rst.Rows.Count > 0) Then
            AttributeTypeID = MyCommon.NZ(rst.Rows(0).Item("AttributeTypeID"), -1)
            Description = MyCommon.NZ(rst.Rows(0).Item("Description"), "")
            ExtID = MyCommon.NZ(rst.Rows(0).Item("ExtID"), "")
            ValueCount = MyCommon.NZ(rst.Rows(0).Item("ValueCount"), 0)
            AllowDelete = MyCommon.NZ(rst.Rows(0).Item("Deletable"), True)
            iReadOnlyAttribute = MyCommon.NZ(rst.Rows(0).Item("ReadOnlyAttribute"), 0)
        ElseIf (AttributeTypeID > -1) Then
            'No rows returned, so assume it's deleted
            Send("")
            Send("<div id=""intro"">")
            Send("  <h1 id=""title"">" & Copient.PhraseLib.Lookup("term.attribute", LanguageID) & " #" & AttributeTypeID & "</h1>")
            Send("</div>")
            Send("<div id=""main"">")
            Send("  <div id=""infobar"" class=""red-background"">")
            Send("    " & Copient.PhraseLib.Lookup("error.deleted", LanguageID))
            Send("  </div>")
            Send("</div>")
            GoTo done
        End If
    End If
%>
<form action="#" id="mainform" name="mainform">
<%
    Send("<input type=""hidden"" id=""AttributeTypeID"" name=""AttributeTypeID"" value=""" & AttributeTypeID & """ />")
    Send("<input type=""hidden"" id=""NewValue"" name=""NewValue"" value="""" />")
    Send("<input type=""hidden"" id=""DeleteValue"" name=""DeleteValue"" value="""" />")
    Send("<input type=""hidden"" id=""SaveValue"" name=""SaveValue"" value="""" />")
    Send("<input type=""hidden"" id=""SaveValueExtID"" name=""SaveValueExtID"" value="""" />")
    Send("<input type=""hidden"" id=""SaveValueDesc"" name=""SaveValueDesc"" value="""" />")
    Send("<input type=""hidden"" id=""selectedRadioID"" name=""selectedRadioID"" value="""" />")
%>
<div id="intro">
    <h1 id="title">
        <%
            If AttributeTypeID = -1 Then
                Sendb(Copient.PhraseLib.Lookup("term.newattribute", LanguageID))
            Else
                Sendb(Copient.PhraseLib.Lookup("term.attribute", LanguageID) & " #" & AttributeTypeID & ": " & MyCommon.TruncateString(Description, 40))
            End If
        %>
    </h1>
    <div id="controls">
        <%
            If (AttributeTypeID = -1) Then
                If (AllowEditing) Then
                    Send_Save()
                End If
            Else
                If (AllowEditing) Then
                    Send("<input type=""button"" class=""regular"" id=""actions"" name=""actions"" value=""" & Copient.PhraseLib.Lookup("term.actions", LanguageID) & " ▼"" onclick=""toggleDropdown();"" />")
                    Send("<div class=""actionsmenu"" id=""actionsmenu"">")
                    Send_Save()
                    If AllowDelete Then Send_Delete()
                    Send_New()
                    Send("</div>")
                End If
                If MyCommon.Fetch_SystemOption(75) Then
                    If (Logix.UserRoles.AccessNotes) Then
                        Send_NotesButton(39, AttributeTypeID, AdminUserID)
                    End If
                End If
            End If
        %>
    </div>
</div>
<div id="main">
    <%
        If (infoMessage <> "") Then
            Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
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
                Send("<table summary=""" & Copient.PhraseLib.Lookup("term.identification", LanguageID) & """>")
                Send("  <tr>")
                Send("    <td><label for=""ExtID"">" & Copient.PhraseLib.Lookup("term.externalid", LanguageID) & ":</label></td>")
                Send("    <td><input type=""text"" class=""long"" id=""ExtID"" name=""ExtID"" maxlength=""50"" value=""" & ExtID.Replace("""", "&quot;") & """" & IIf(AllowEditing AndAlso TypeInUseOfferCount = 0, "", " disabled=""disabled""") & " /></td>")
                Send("  </tr>")
                Send("  <tr>")
                Send("    <td><label for=""Description"">" & Copient.PhraseLib.Lookup("term.description", LanguageID) & ":</label></td>")
                Send("    <td><input type=""text"" class=""long"" id=""Description"" name=""Description"" maxlength=""100"" value=""" & Description.Replace("""", "&quot;") & """" & IIf(AllowEditing AndAlso TypeInUseOfferCount = 0, "", " disabled=""disabled""") & " /></td>")
                Send("  </tr>")
                Send("  <tr>")
                Send("    <td>" & Copient.PhraseLib.Lookup("term.engines", LanguageID) & ":</td>")
                Send("    <td>")
                MyCommon.QueryStr = "select PE.EngineID, PE.Description as EngineName, PE.PhraseID as EnginePhraseID, " & _
                                    "  PEST.SubTypeID, PEST.SubTypeName as EngineSubTypeName, PEST.PhraseID as SubTypePhraseID " & _
                                    "from PromoEngines as PE with (NoLock) " & _
                                    "left join PromoEngineSubTypes as PEST on PEST.PromoEngineID=PE.EngineID " & _
                                    "where PE.EngineID=2 and PE.Installed=1 and PEST.Installed=1;"
                rst = MyCommon.LRT_Select
                If rst.Rows.Count > 0 Then
                    For Each row In rst.Rows
                        MyCommon.QueryStr = "select AttributeTypeID from AttributeTypeEngines with (NoLock) where AttributeTypeID=" & AttributeTypeID & " and EngineID=2 and EngineSubTypeID=" & MyCommon.NZ(row.Item("SubTypeID"), 0) & ";"
                        rst2 = MyCommon.LRT_Select
                        If (AllowEditing AndAlso TypeInUseOfferCount = 0) Then
                            Sendb("      <input type=""checkbox"" id=""EngineSubTypeID-" & MyCommon.NZ(row.Item("SubTypeID"), 0) & """ name=""EngineSubTypeID-" & MyCommon.NZ(row.Item("SubTypeID"), 0) & """")
                            If rst2.Rows.Count > 0 Then
                                Send(" checked=""checked""")
                            End If
                            Sendb(" value=""1"" /><label for=""EngineSubTypeID-" & MyCommon.NZ(row.Item("SubTypeID"), 0) & """>")
                            Sendb(Copient.PhraseLib.Lookup(CInt(MyCommon.NZ(row.Item("EnginePhraseID"), 0)), LanguageID, MyCommon.NZ(row.Item("EngineName"), "")) & " ")
                            Sendb(Copient.PhraseLib.Lookup(CInt(MyCommon.NZ(row.Item("SubTypePhraseID"), 0)), LanguageID, MyCommon.NZ(row.Item("EngineSubTypeName"), "")))
                            Send("</label><br />")
                        Else
                            If rst2.Rows.Count > 0 Then
                                Sendb("      ")
                                Sendb(Copient.PhraseLib.Lookup(CInt(MyCommon.NZ(row.Item("EnginePhraseID"), 0)), LanguageID, MyCommon.NZ(row.Item("EngineName"), "")) & " ")
                                Sendb(Copient.PhraseLib.Lookup(CInt(MyCommon.NZ(row.Item("SubTypePhraseID"), 0)), LanguageID, MyCommon.NZ(row.Item("EngineSubTypeName"), "")))
                                Send("<br />")
                            End If
                        End If
                    Next
                End If
                Send("    </td>")
                Send("  </tr>")
                If MyCommon.Fetch_SystemOption(119) Then
                    Send("  <tr>")
                    Send("    <td>" & Copient.PhraseLib.Lookup("term.miscellaneous", LanguageID) & ":</td>")
                    Send("    <td>")
                    MyCommon.QueryStr = "select ReadOnlyAttribute from AttributeTypes with (NoLock) where AttributeTypeID=" & AttributeTypeID & ";"
                    rst = MyCommon.LRT_Select
                    If (AllowEditing) Then
                        Sendb("      <input type=""checkbox"" id=""ReadOnlyAttribute"" name=""ReadOnlyAttribute"" ")
                        If rst.Rows.Count > 0 AndAlso MyCommon.NZ(rst.Rows(0).Item("ReadOnlyAttribute"), False) Then
                            Send(" checked=""checked""")
                        End If
                        Send(" /><label for=""ReadOnlyAttribute"">" & Copient.PhraseLib.Lookup("term.readonly", LanguageID) & "</label><br />")
                    Else
                        Sendb("      <input type=""checkbox"" id=""ReadOnlyAttribute"" name=""ReadOnlyAttribute"" ")
                        If rst.Rows.Count > 0 AndAlso MyCommon.NZ(rst.Rows(0).Item("ReadOnlyAttribute"), False) Then
                            Send(" checked=""checked""")
                        End If
                        Send(" disabled=""disabled"" /><label for=""ReadOnlyAttribute"">" & Copient.PhraseLib.Lookup("term.readonly", LanguageID) & "</label><br />")
                    End If
                    Send("    </td>")
                    Send("  </tr>")
                End If

                Send("</table>")
            %>
            <hr class="hidden" />
        </div>
        <div class="box" id="values" <% Sendb(IIf(Attributetypeid = -1, " style=""display: none;""", "")) %>>
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.values", LanguageID))%>
                </span>
            </h2>
            <%
                If AttributeTypeID > -1 Then
                    If AllowEditing Then
                        Send("<input type=""text"" id=""NewValueExtID"" name=""NewValueExtID"" maxlength=""50"" value=""" & Copient.PhraseLib.Lookup("term.NewExtID", LanguageID) & "..."" onfocus=""javascript:toggleNewValueExtID('clear');"" onblur=""javascript:toggleNewValueExtID('restore');"" />")
                        Send("<input type=""text"" id=""NewValueDesc"" name=""NewValueDesc"" maxlength=""100"" value=""" & Copient.PhraseLib.Lookup("term.NewDescription", LanguageID) & "..."" onfocus=""javascript:toggleNewValueDesc('clear');"" onblur=""javascript:toggleNewValueDesc('restore');"" />")
                        Send("<input type=""button"" value=""" & Copient.PhraseLib.Lookup("term.SaveNewValue", LanguageID) & """ onclick=""javascript:newValue();"" />")
                        Send("<br />")
                        Send("<br />")
                        Send("<br />")
                    End If
                    Send("<table summary=""" & Copient.PhraseLib.Lookup("term.values", LanguageID) & """>")
                    Send("  <tr>")
                    If AllowEditing Then
                        Send("    <th style=""width:32px;"">" & Left(Copient.PhraseLib.Lookup("term.delete", LanguageID), 3) & "</th>")
                    End If
                    Send("    <th>" & Copient.PhraseLib.Lookup("term.externalid", LanguageID) & "</th>")
                    Send("    <th colspan=""2"">" & Copient.PhraseLib.Lookup("term.description", LanguageID) & "</th>")
                    Send("    <th style=""width:32px;"">" & Copient.PhraseLib.Lookup("term.default", LanguageID) & "</th>")
                    Send("  </tr>")
                    MyCommon.QueryStr = "select AttributeValueID, ExtID, Description, DefaultValue from AttributeValues with (NoLock) " & _
                                        "where Deleted=0 and AttributeTypeID=" & AttributeTypeID & ";"
                    rst = MyCommon.LRT_Select
                    If rst.Rows.Count > 0 Then
                        For Each row In rst.Rows
                            AttributeValueID = MyCommon.NZ(row.Item("AttributeValueID"), 0)
                            'Determine if the value is in use (by checking for any associated offers or associated customers).
                            If AttributeTypeID > -1 Then
                                MyCommon.QueryStr = "select distinct I.IncentiveID as OfferID, I.IncentiveName as OfferName, I.EndDate as ProdEndDate,buy.ExternalBuyerId as BuyerID from CPE_Incentives as I " & _
                                                    "inner join CPE_RewardOptions as RO with (NoLock) on RO.IncentiveID=I.IncentiveID and I.Deleted=0 and RO.Deleted=0 " & _
                                                    "inner join CPE_IncentiveAttributes as IA with (NoLock) on IA.RewardOptionID=RO.RewardOptionID " & _
                                                    "inner join CPE_IncentiveAttributeTiers as IAT with (NoLock) on IAT.IncentiveAttributeID=IA.IncentiveAttributeID " & _
                                                      "left outer join Buyers as buy with (nolock) on buy.BuyerId= I.BuyerId " & _
                                                    "where I.IsTemplate=0 and IAT.AttributeTypeID=" & AttributeTypeID & " " & _
                                                    "and " & AttributeValueID & " in (select items from dbo.Split(IAT.AttributeValues, ',')) " & _
                                                    "order by OfferName;"
                                rst = MyCommon.LRT_Select
                                ValueInUseOfferCount = rst.Rows.Count
                                MyCommon.QueryStr = "select count(CustomerPK) as Customers from CustomerAttributes as CA with (NoLock) " & _
                                                    "where AttributeTypeID=" & AttributeTypeID & " and AttributeValueID=" & AttributeValueID & " and Deleted=0;"
                                rst = MyCommon.LXS_Select
                                ValueInUseCustomerCount = rst.Rows(0).Item("Customers")
                                If (ValueInUseOfferCount = 0) And (ValueInUseCustomerCount = 0) Then
                                    ValueInUse = False
                                Else
                                    ValueInUse = True
                                End If
                            End If
                            Send("  <tr>")
                            If AllowEditing Then
                                Send("    <td><input type=""button"" value=""X"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ name=""ex"" id=""ex-" & MyCommon.NZ(row.Item("AttributeValueID"), 0) & """ class=""ex""" & IIf(Not ValueInUse, "", " disabled=""disabled""") & " onclick=""javascript:deleteValue('" & MyCommon.NZ(row.Item("AttributeValueID"), 0) & "')"" /></td>")
                            End If
                            Send("    <td>" & MyCommon.NZ(row.Item("ExtID"), "") & "</td>")
                            Send("    <td>" & MyCommon.NZ(row.Item("Description"), "") & "</td>")
                            Send("    <td></td><td><input type=""radio"" id=""defaultradio" & MyCommon.NZ(row.Item("AttributeValueID"), 0) & """ name=""defaultradio"" value=""" & MyCommon.NZ(row.Item("AttributeValueID"), 0) & """" & IIf(MyCommon.NZ(row.Item("DefaultValue"), False), " checked=""checked"" ", "") & "onclick=""selectedRadio(" & MyCommon.NZ(row.Item("AttributeValueID"), 0) & ");"" " & "  />" & "</td>")
                            Send("    <td>")
                            Send("    </td>")
                            Send("  </tr>")
                        Next
                        Send("    <td></td><td><input type=""radio"" id=""defaultradioGroupSetup"" name=""defaultradio"" value=""GroupSetup"" style=""display: none;"" />" & "</td>")
                    End If
                    Send("</table>")
                    Send("<input type=""button"" value=""" & Copient.PhraseLib.Lookup("term.cleardefault", LanguageID) & """ onclick=""javascript:clearDefaultButton(document.mainform.defaultradio);"" />")
                End If
            %>
            <hr class="hidden" />
        </div>
    </div>
    <div id="gutter">
    </div>
    <div id="column2">
        <div class="box" id="offers" <% Sendb(IIf(Attributetypeid = -1, " style=""display:none;""", "")) %>>
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.associatedoffers", LanguageID))%>
                </span>
            </h2>
            <div class="boxscroll">
                <%
                    Dim assocName As String = ""
                    If (AttributeTypeID > -1) Then
                        If rstAssociatedOffers.Rows.Count > 0 Then
                            For Each row In rstAssociatedOffers.Rows
                                If (MyCommon.Fetch_UE_SystemOption(168) = "1" AndAlso MyCommon.NZ(row.Item("BuyerID"), "") <> "") Then
                                    assocName = "Buyer " + row.Item("BuyerID").ToString() + " - " + MyCommon.NZ(row.Item("Name"), "").ToString()
                                Else
                                    assocName = MyCommon.NZ(row.Item("Name"), Copient.PhraseLib.Lookup("term.unknown", LanguageID))
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
                            Send("        " & Copient.PhraseLib.Lookup("term.none", LanguageID) & "<br />")
                        End If
                    Else
                        Send("        " & Copient.PhraseLib.Lookup("term.none", LanguageID) & "<br />")
                    End If
                %>
            </div>
            <hr class="hidden" />
        </div>
        <div class="box" id="customers" <% Sendb(IIf(Attributetypeid = -1, " style=""display:none;""", "")) %>>
            <h2>
                <span>
                    <% Sendb(Copient.PhraseLib.Lookup("term.associatedcustomers", LanguageID))%>
                </span>
            </h2>
            <div class="boxscroll">
                <%
                    If (AttributeTypeID > -1) Then
                        If TypeInUseCustomerCount > 0 Then
                            Copient.PhraseLib.Detokenize("attribute-edit.InUseBy", LanguageID, TypeInUseCustomerCount) 'Attribute is in use by {0} customer(s) with the following cards
                            If (TypeInUseCustomerCount > 100) Then
                                Send("; " & Copient.PhraseLib.Detokenize("term.firstshown", LanguageID, "100") & ":<br />")
                            Else
                                Send(":<br />")
                            End If
                            Send("<br class=""half"" />")
                            For i = 0 To (rstAssociatedCustomers.Rows.Count - 1)
                                If (i > 0) Then
                                    If (rstAssociatedCustomers.Rows(i).Item("CustomerPK") <> rstAssociatedCustomers.Rows(i - 1).Item("CustomerPK")) Then
                                        If Shaded = "shaded" Then
                                            Shaded = ""
                                        Else
                                            Shaded = "shaded"
                                        End If
                                    End If
                                Else
                                    If Shaded = "shaded" Then
                                        Shaded = ""
                                    Else
                                        Shaded = "shaded"
                                    End If
                                End If
                                Sendb(" <p style=""margin-bottom:0;"" class=""" & Shaded & """><a href=""customer-general.aspx?CustPK=" & rstAssociatedCustomers.Rows(i).Item("CustomerPK") & """>" & MyCommon.NZ(rstAssociatedCustomers.Rows(i).Item("ExtCardID"), "[" & Copient.PhraseLib.Detokenize("attribute-edit.CustomerHasNoCard", LanguageID, rstAssociatedCustomers.Rows(i).Item("CustomerPK")) & "]") & "</a></p>")
                            Next
                        Else
                            Send("        " & Copient.PhraseLib.Lookup("term.none", LanguageID) & "<br />")
                        End If
                    Else
                        Send("        " & Copient.PhraseLib.Lookup("term.none", LanguageID) & "<br />")
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
    if (window.captureEvents) {
        window.captureEvents(Event.CLICK);
        window.onclick = handlePageClick;
    }
    else {
        document.onclick = handlePageClick;
    }
</script>
<%
    If MyCommon.Fetch_SystemOption(75) Then
        If (Logix.UserRoles.AccessNotes) Then
            Send_Notes(39, AttributeTypeID, AdminUserID)
        End If
    End If
done:
    MyCommon.Close_LogixRT()
    MyCommon.Close_LogixXS()
    Send_BodyEnd("mainform", "ExtID")
    MyCommon = Nothing
    Logix = Nothing
%>
