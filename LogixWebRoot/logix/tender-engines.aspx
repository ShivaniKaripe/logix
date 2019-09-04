<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: tender-engines.aspx 
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
  Dim dst As System.Data.DataTable
  Dim row As DataRow
  Dim rst As System.Data.DataTable
  Dim l_code As String
  Dim l_variety As String
  Dim l_binnumber As String
  Dim l_name As String
  Dim l_OCID As Long
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim i As Integer
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "tender-engines.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  Send_HeadBegin("term.tender")
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
  
  If (Request.QueryString("add") <> "") Then
    l_name = MyCommon.NZ(Request.QueryString("name"), "")
    l_code = MyCommon.NZ(Request.QueryString("code"), "")
    l_variety = MyCommon.NZ(Request.QueryString("variety"), "")
    l_binnumber = MyCommon.NZ(Request.QueryString("binnumber"), "")
    
    l_name = MyCommon.Parse_Quotes(Logix.TrimAll(l_name))
    l_code = MyCommon.Parse_Quotes(Logix.TrimAll(l_code))
    l_variety = MyCommon.Parse_Quotes(Logix.TrimAll(l_variety))
    l_binnumber = MyCommon.Parse_Quotes(Logix.TrimAll(l_binnumber))
    If (l_name = "") OrElse (l_code = "") OrElse (l_code.Length < 2) OrElse (l_variety = "") OrElse (l_variety.Length < 2) Then
      infoMessage = Copient.PhraseLib.Lookup("tender.noname", LanguageID)
    Else
      MyCommon.QueryStr = "SELECT TenderTypeID FROM CPE_TenderTypes with (NoLock) WHERE Deleted=0 and Name = '" & l_name & "'"
      dst = MyCommon.LRT_Select
      If (dst.Rows.Count > 0) Then
        infoMessage = Copient.PhraseLib.Lookup("tender.nameused", LanguageID)
      Else
        MyCommon.QueryStr = "SELECT TenderTypeID FROM CPE_TenderTypes with (NoLock) WHERE Deleted=0 and ExtTenderType = '" & l_code & "' and ExtVariety = '" & l_variety & "' and ExtBinNum = '" & l_binnumber & "';"
        dst = MyCommon.LRT_Select
        If (dst.Rows.Count > 0) Then
          infoMessage = Copient.PhraseLib.Lookup("tender.codeused", LanguageID)
        Else
          MyCommon.QueryStr = "dbo.pt_CPE_TenderTypes_Insert"
          MyCommon.Open_LRTsp()
          MyCommon.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar).Value = Request.QueryString("name")
          MyCommon.LRTsp.Parameters.Add("@ExtTenderType", SqlDbType.NVarChar).Value = l_code
          MyCommon.LRTsp.Parameters.Add("@ExtVariety", SqlDbType.NVarChar).Value = l_variety
          MyCommon.LRTsp.Parameters.Add("@ExtBinNum", SqlDbType.NVarChar).Value = Request.QueryString("binnumber")
          MyCommon.LRTsp.Parameters.Add("@TenderTypeID", SqlDbType.Int).Direction = ParameterDirection.Output
          MyCommon.LRTsp.ExecuteNonQuery()
          MyCommon.Close_LRTsp()
          MyCommon.Activity_Log(22, l_OCID, AdminUserID, Copient.PhraseLib.Lookup("history.tender-create", LanguageID))
        End If
      End If
    End If
  ElseIf (Request.QueryString("delete") <> "") Then
    l_OCID = MyCommon.Extract_Val(Request.QueryString("tenders"))
    
    ' check if this tender type is used by any offers or templates
    MyCommon.QueryStr = "select distinct I.IncentiveID, I.IncentiveName, I.IsTemplate from CPE_IncentiveTenderTypes ITT with (NoLock) " & _
                        "inner join CPE_RewardOptions RO with (NoLock) on RO.RewardOptionID = ITT.RewardOptionID " & _
                        "inner join CPE_Incentives I with (NoLock) on I.IncentiveID = RO.IncentiveID " & _
                        "where I.Deleted = 0 and RO.Deleted=0 and ITT.Deleted=0 and TenderTypeID = " & l_OCID
    dst = MyCommon.LRT_Select
    If (dst.Rows.Count > 0) Then
      infoMessage = Copient.PhraseLib.Lookup("cpe-tender-types.inuse", LanguageID) & " (" & Copient.PhraseLib.Lookup("term.offers", LanguageID) & ": "
      i = 0
      For Each row In dst.Rows
        infoMessage &= IIf(i > 0, ",", "") & "<a style=""color:yellow;text-decoration:underline;"" href=""CPEoffer-con.aspx?OfferID=" & MyCommon.NZ(row.Item("IncentiveID"), "") & """>" & MyCommon.NZ(row.Item("IncentiveID"), "") & "</a>"
        i += 1
        If (i > 10) Then
          infoMessage &= " ..."
          Exit For
        End If
      Next
      infoMessage &= ")"
    Else
      MyCommon.QueryStr = "dbo.pt_CPE_TenderTypes_Delete"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@TenderTypeID", SqlDbType.Int).Value = l_OCID
      MyCommon.LRTsp.ExecuteNonQuery()
      MyCommon.Close_LRTsp()
      MyCommon.Activity_Log(22, l_OCID, AdminUserID, Copient.PhraseLib.Lookup("history.tender-delete", LanguageID))
    End If
    
  End If
%>
<form action="#" id="mainform" name="mainform">
  <div id="intro">
    <h1 id="title">
      <% Sendb(Copient.PhraseLib.Lookup("term.tendertypes", LanguageID))%>
    </h1>
    <div id="controls">
      <%
        If MyCommon.Fetch_SystemOption(75) Then
          If (Logix.UserRoles.AccessNotes) Then
            Send_NotesButton(27, 0, AdminUserID)
          End If
        End If
      %>
    </div>
  </div>
  <div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <div id="column1">
      <%
        If (Logix.UserRoles.EditTenderTypes = False) Then
          Send("      <select id=""tenders"" name=""tenders"" class=""long"" size=""15"">")
          MyCommon.QueryStr = "SELECT TenderTypeID, Name, ExtTenderType, ExtVariety, ExtBinNum FROM CPE_TenderTypes with (NoLock) order by ExtTenderType, ExtVariety, ExtBinNum;"
          dst = MyCommon.LRT_Select
          For Each row In dst.Rows
            Send("        <option value=""" & MyCommon.NZ(row.Item("TenderTypeID"), 0) & """>" & MyCommon.NZ(row.Item("ExtTenderType"), "") & " " & MyCommon.NZ(row.Item("ExtVariety"), "") & " " & MyCommon.NZ(row.Item("ExtBinNum"), "") & " - " & MyCommon.NZ(row.Item("Name"), "") & "</option>")
          Next
          Send("      </select>")
          Send("    </div>")
          Send("</form>")
          GoTo done
        End If
      %>
      <div class="box" id="tenderadd">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("tender.add", LanguageID))%>
          </span>
        </h2>
        <table>
          <tr>
            <td><label for="code"><% Sendb(Copient.PhraseLib.Lookup("term.code", LanguageID))%>:</label></td>
            <td><input type="text" id="code" name="code" class="shortest" value="" maxlength="2" /><span class="red">*</span></td>
          </tr>
          <tr>
            <td><label for="variety"><% Sendb(Copient.PhraseLib.Lookup("term.variety", LanguageID))%>:</label></td>
            <td><input type="text" id="variety" name="variety" class="shortest" value="" maxlength="2" /><span class="red">*</span></td>
          </tr>
          <tr>
            <td><label for="binnumber"><% Sendb(Copient.PhraseLib.Lookup("term.bin", LanguageID))%>:</label></td>
            <td><input type="text" id="binnumber" name="binnumber" class="short" value="" maxlength="8" /></td>
          </tr>
          <tr>
            <td><label for="name"><% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>:</label></td>
            <td><input type="text" id="name" name="name" class="medium" value="" maxlength="100" /><span class="red">*</span></td>
          </tr>
          <tr>
            <td></td>
            <td><input type="submit" class="regular" id="add" name="add" value="<% Sendb(Copient.PhraseLib.Lookup("term.add", LanguageID)) %>" /></td>
          </tr>
        </table>
        <hr class="hidden" />
      </div>
      <span class="red">* <small><% Sendb(Copient.PhraseLib.Lookup("term.RequiredField", LanguageID)) %></small></span>
    </div>

    <div id="gutter">
    </div>
  
    <div id="column2">
      <div class="box" id="tenderdel">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("tender.delete", LanguageID))%>
          </span>
        </h2>
        <select class="longest" id="tenders" onchange="enableDelete();" name="tenders" size="15">
          <%
            MyCommon.QueryStr = "SELECT TenderTypeID, Name, ExtTenderType, ExtVariety, ExtBinNum FROM CPE_TenderTypes with (NoLock) where Deleted = 0 order by ExtTenderType, ExtVariety, ExtBinNum;"
            dst = MyCommon.LRT_Select
            For Each row In dst.Rows
              Send("          <option value=""" & MyCommon.NZ(row.Item("TenderTypeID"), 0) & """>" & MyCommon.NZ(row.Item("ExtTenderType"), "") & " " & MyCommon.NZ(row.Item("ExtVariety"), "") & " " & MyCommon.NZ(row.Item("ExtBinNum"), "") & " - " & MyCommon.NZ(row.Item("Name"), "") & "</option>")
            Next
          %>
        </select>
        <br />
        <input type="submit" class="regular" disabled="disabled" id="delete" name="delete" onclick="if(confirm('<% Sendb(Copient.PhraseLib.Lookup("confirm.delete", LanguageID)) %>')){}else{return false}" value="<% Sendb(Copient.PhraseLib.Lookup("term.delete", LanguageID)) %>" /><br />
      </div>
    </div>
  
  </div>
  
    <br clear="all" />

</form>
<script type="text/javascript">
    function enableDelete(){
        document.getElementById("delete").disabled = false;
    }
</script>
<%
  If MyCommon.Fetch_SystemOption(75) Then
    If (Logix.UserRoles.AccessNotes) Then
      Send_Notes(27, 0, AdminUserID)
    End If
  End If
done:
  Send_BodyEnd("mainform", "code")
  MyCommon.Close_LogixRT()
  MyCommon = Nothing
  Logix = Nothing
%>
