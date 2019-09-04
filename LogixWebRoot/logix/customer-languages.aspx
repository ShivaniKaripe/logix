<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%@ Import Namespace="System.Xml" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: customer-languages.aspx 
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
%>
  
<script runat="server">


    Dim MyCommon As New Copient.CommonInc
    Dim Logix As New Copient.LogixInc
    Dim LanguagePrefID As Integer = 0


    Sub Load_Language_PrefID()
        Dim dt As DataTable

        If MyCommon.IsIntegrationInstalled(Copient.CommonIncConfigurable.Integrations.PREFERENCE_MANAGER) Then
            MyCommon.QueryStr = "select PreferenceID from MetaPrefs where PKID=1;"
            dt = MyCommon.PMRT_Select
            If dt.Rows.Count > 0 Then
                LanguagePrefID = MyCommon.NZ(dt.Rows(0).Item("PreferenceID"), 0)
            End If
        Else
            LanguagePrefID = 0
        End If
    End Sub


    Function Is_In_Use(ByVal LanguageIDs As String(), ByRef Message As String) As Boolean
        Dim InUse As Boolean = False
        Dim dt As DataTable

        If MyCommon.IsIntegrationInstalled(Copient.CommonIncConfigurable.Integrations.PREFERENCE_MANAGER) Then

            ' ensure that at least one element is in the array before doing the string.join
            If LanguageIDs Is Nothing OrElse LanguageIDs.Length = 0 Then
                ReDim LanguageIDs(0)
                LanguageIDs(0) = -1
            End If

            ' check if any of the unselected languages are currently in use by an offer with a language preference condition using that language.
            MyCommon.QueryStr = "select distinct LANG.LanguageID, " & _
                                "                case when PT.Phrase is null then LANG.Name " & _
                                "                     else CONVERT(nvarchar(1000), PT.Phrase) " & _
                                "                end as TranslatedName, " & _
                                "                INC.IncentiveID, INC.IncentiveName  " & _
                                "from CPE_IncentivePrefs AS CIP with (NoLock) " & _
                                "inner join CPE_IncentivePrefTiers as CIPT with (NoLock) " & _
                                "  on CIPT.IncentivePrefsID = CIP.IncentivePrefsID " & _
                                "inner join CPE_IncentivePrefTierValues as CIPTV with (NoLock) " & _
                                "  on CIPTV.IncentivePrefTiersID = CIPT.IncentivePrefTiersID " & _
                                "inner join CPE_RewardOptions as RO with (NoLock) " & _
                                "  on RO.RewardOptionID = CIP.RewardOptionID " & _
                                "inner join CPE_Incentives as INC with (NoLock) " & _
                                "  on INC.IncentiveID = RO.RewardOptionID " & _
                                "inner join Languages as LANG with (NoLock) " & _
                                "  on LANG.JavaLocaleCode = CIPTV.Value " & _
                                "left join UIPhrases as UIP with (NoLock) " & _
                                "  on UIP.Name = LANG.PhraseTerm " & _
                                "left join PhraseText as PT with (NoLock) " & _
                                "  on PT.PhraseID = UIP.PhraseID and PT.LanguageID = 2 " & _
                                "where INC.Deleted=0 and RO.Deleted=0 and CIP.PreferenceID=" & LanguagePrefID & _
                                "  and LANG.LanguageID not in (" & String.Join(",", LanguageIDs) & ");"
            dt = MyCommon.LRT_Select
            InUse = (dt.Rows.Count > 0)

            If InUse Then
                Message = Copient.PhraseLib.Lookup("customer-languages.LanguagesInUse", LanguageID) & "<br />"
                For Each row As DataRow In dt.Rows
                    Message &= MyCommon.NZ(row.Item("TranslatedName"), "") & " - (" & MyCommon.NZ(row.Item("IncentiveID"), 0) & ") " & _
                               MyCommon.NZ(row.Item("IncentiveName"), "") & "<br />"
                Next
            End If
        End If

        Return InUse
    End Function


    Sub Save_Languages(ByVal LanguageIDs As String())
        Dim LangCSV As String = ""

        If LanguageIDs IsNot Nothing AndAlso LanguageIDs.Length > 0 Then
            MyCommon.QueryStr = "update Languages set AvailableForCustFacing=1 where LanguageID in (" & String.Join(",", LanguageIDs) & ");"
            MyCommon.LRT_Execute()

            MyCommon.QueryStr = "update Languages set AvailableForCustFacing=0 where LanguageID not in (" & String.Join(",", LanguageIDs) & ");"
            MyCommon.LRT_Execute()
        Else
            ' wipe them out, all of them
            MyCommon.QueryStr = "update Languages set AvailableForCustFacing=0;"
            MyCommon.LRT_Execute()
        End If

        ' log this activity
        If LanguageIDs IsNot Nothing Then
            LangCSV = " - " & Left(String.Join(",", LanguageIDs), 255)
        End If
        MyCommon.Activity_Log(49, 2, 34, AdminUserID, Copient.PhraseLib.Lookup("history.editedcustomerlanguages", LanguageID) & LangCSV, _
                              Left("[history.editedcustomerlanguages]" & LangCSV, 255))
    End Sub


    Function Sync_Pref_ListItems(ByVal LanguageIDs As String(), ByRef Message As String) As Boolean
        Dim ConnInc As New Copient.ConnectorInc
        Dim IntegrationValues As New Copient.CommonInc.IntegrationValues
        Dim RootURI As String = ""
        Dim RespXML As XmlDocument = New XmlDocument()
        Dim RespNode As XmlNodeList
        Dim Success As Boolean = False
        Dim Response As String = ""
        If MyCommon.IsIntegrationInstalled(Copient.CommonIncConfigurable.Integrations.PREFERENCE_MANAGER, IntegrationValues) Then
            RootURI = IntegrationValues.HTTP_RootURI
            If RootURI IsNot Nothing AndAlso RootURI.Trim.Length > 0 Then
                If Right(RootURI, 1) <> "/" Then RootURI &= "/"
                Try

                    Response = ConnInc.Retrieve_HttpResponse(RootURI & "connectors/AMS.asmx/SyncLanguagePref", Build_Form_Data(LanguageIDs))
                    RespXML.LoadXml(Response)
                    If RespXML IsNot Nothing Then

                        RespNode = RespXML.GetElementsByTagName("boolean")
                        If RespNode IsNot Nothing AndAlso RespNode.Count >= 1 Then
                            Boolean.TryParse(RespNode(0).InnerText, Success)
                        Else
                            Throw New ApplicationException("Unexpected response returned from web service sync method. Response = " & RespXML.OuterXml)
                        End If
                    Else
                        Throw New ApplicationException("Web service synchronization method failed to return a valid response.")
                    End If
                Catch ex As Exception
                    If Message <> "" Then Message &= "<br />"
                    Message &= Copient.PhraseLib.Detokenize("customer-languages.SyncError", LanguageID, ex.ToString)
                End Try
            End If
        Else
            ' preference manager is not installed so there is no need for syncrhronization.
            Success = True
        End If

        Return Success
    End Function

    Function Build_Form_Data(ByVal LanguageIDs As String()) As String
        Dim FormDataBuf As New StringBuilder()

        FormDataBuf.Append("GUID=" & Get_Connnector_GUID() & "&LanguageIDs=")
        If LanguageIDs IsNot Nothing Then
            FormDataBuf.Append(String.Join(",", LanguageIDs))
        End If

        Return FormDataBuf.ToString
    End Function

    Function Get_Connnector_GUID() As String
        Dim GUID As String = ""
        Dim dt As DataTable
        Const AMS_WS_CONNECTOR_ID As Integer = 5

        ' need a GUID for authenication to the AMS Web service, just grab the first one and use it.
        MyCommon.QueryStr = "select top 1 [GUID] from ConnectorGUIDs with (NoLock) " & _
                            "where ConnectorID=" & AMS_WS_CONNECTOR_ID
        dt = MyCommon.PMRT_Select
        If dt.Rows.Count > 0 Then
            GUID = MyCommon.NZ(dt.Rows(0).Item("GUID"), "")
        End If

        Return GUID
    End Function

</script>
<%  
    Dim CopientFileName As String = Page.Request.Url.Segments(Page.Request.Url.Segments.GetUpperBound(0))
    Dim CopientFileVersion As String = "7.3.1.138972"
    Dim CopientProject As String = "Copient Logix"
    Dim CopientNotes As String = ""

    Dim AdminUserID As Long
    Dim rst As System.Data.DataTable
    Dim row As System.Data.DataRow
    Dim infoMessage As String = ""
    Dim Handheld As Boolean = False

    Dim LangID As Integer = 0
    Dim LangName As String = ""
    Dim LangChecked As Boolean = False
    Dim LanguageIDs() As String
    Dim i As Integer = 0

    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If

    Response.Expires = 0
    MyCommon.AppName = "customer-languages.aspx"
    MyCommon.Open_LogixRT()
    If MyCommon.IsIntegrationInstalled(Copient.CommonIncConfigurable.Integrations.PREFERENCE_MANAGER) Then
        MyCommon.Open_PrefManRT()
    End If
    AdminUserID = Verify_AdminUser(MyCommon, Logix)
    CMS.AMS.CurrentRequest.Resolver.AppName = MyCommon.AppName
    If (Request.QueryString("save") <> "") Then
        LanguageIDs = Request.QueryString.GetValues("languageid")
        Load_Language_PrefID()

        If Not Is_In_Use(LanguageIDs, infoMessage) Then
            Save_Languages(LanguageIDs)
            If Not Sync_Pref_ListItems(LanguageIDs, infoMessage) AndAlso infoMessage = "" Then
                infoMessage = Copient.PhraseLib.Lookup("customer-languages.SyncFailed", LanguageID)
            End If
        End If
    End If

    Send_HeadBegin("term.customerlanguages")
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

    If (Logix.UserRoles.AccessSystemSettings = False) Then
        Send_Denied(1, "perm.admin-settings")
        GoTo done
    End If
%>
<form action="#" id="mainform" name="mainform">
  <div id="intro">
    <h1 id="title">
      <% Sendb(Copient.PhraseLib.Lookup("term.customerlanguages", LanguageID))%>
    </h1>
    <div id="controls">
      <%
        If (Logix.UserRoles.EditSystemOptions = True) Then
          Send_Save()
        End If
        If MyCommon.Fetch_SystemOption(75) Then
          If (Logix.UserRoles.AccessNotes) Then
            Send_NotesButton(28, 0, AdminUserID)
          End If
        End If
      %>
    </div>
  </div>
  <div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <table summary="<% Sendb(Copient.PhraseLib.Lookup("term.settings", LanguageID)) %>">
      <%
        MyCommon.QueryStr = "select LANG.LanguageID, LANG.AvailableForCustFacing, " & _
                            "  case when PT.Phrase is null then LANG.Name " & _
                            "       else convert(nvarchar(1000), PT.Phrase)" & _
                            "  end as TranslatedName " & _
                            "from Languages  as LANG with (NoLock) " & _
                            "left join UIPhrases as UIP with (NoLock) on UIP.Name = LANG.PhraseTerm " & _
                            "left join PhraseText as PT with (NoLock) on PT.PhraseID = UIP.PhraseID and PT.LanguageID=" & MyCommon.GetAdminUser.LanguageID & " " & _
                            "order by AvailableForCustFacing desc, TranslatedName;"
                            
        rst = MyCommon.LRT_Select
        If (rst.Rows.Count > 0) Then
          i = 1
          For Each row In rst.Rows
            LangID = MyCommon.NZ(row.Item("LanguageID"), 0)
            LangName = MyCommon.NZ(row.Item("TranslatedName"), "")
            LangChecked = MyCommon.NZ(row.Item("AvailableForCustFacing"), False)
            
            Send("  <tr>")
            Send("    <td>")
            Send("      <input type=""checkbox"" name=""languageid"" id=""languageid" & i & """ value=""" & LangID & """" & IIf(LangChecked, "checked=""checked""", "") & ">")
            Send("    </td>")
            Send("    <td>")
            Send("      <label for=""languageid" & i & """>" & LangName & "</label>")
            Send("    </td>")
            Send("  </tr>")
            i += 1
          Next
        End If
      %>
    </table>
  </div>
</form>
<%
  If MyCommon.Fetch_SystemOption(75) Then
    If (Logix.UserRoles.AccessNotes) Then
      Send_Notes(28, 0, AdminUserID)
    End If
  End If
done:
  Send_BodyEnd()
  MyCommon.Close_LogixRT()
  If MyCommon.IsIntegrationInstalled(Copient.CommonIncConfigurable.Integrations.PREFERENCE_MANAGER) Then
    MyCommon.Close_PrefManRT()
  End If
  Logix = Nothing
  MyCommon = Nothing
%>
