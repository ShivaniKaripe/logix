<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>          
<%
  ' *****************************************************************************
  ' * FILENAME: tender.aspx 
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
  Dim l_name As String
  Dim l_OCID As Long
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "tender.aspx"
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
    l_name = MyCommon.Parse_Quotes(Logix.TrimAll(l_name))
    l_code = MyCommon.Parse_Quotes(Logix.TrimAll(l_code))
    If (l_name = "") Or (l_code = "") Then
      infoMessage = Copient.PhraseLib.Lookup("tender.noname", LanguageID)
    ElseIf Not IsNumeric(l_code) Then
      infoMessage = Copient.PhraseLib.Lookup("tender.badcode", LanguageID)
    ElseIf (l_code < 1) Or (Int(l_code) <> l_code) Then
      infoMessage = Copient.PhraseLib.Lookup("tender.badcode", LanguageID)
    Else
      MyCommon.QueryStr = "SELECT TenderTypeID FROM TenderTypes with (NoLock) WHERE Description = '" & l_name & "'"
      dst = MyCommon.LRT_Select
      If (dst.Rows.Count > 0) Then
        infoMessage = Copient.PhraseLib.Lookup("tender.nameused", LanguageID)
      Else
        MyCommon.QueryStr = "SELECT TenderTypeID FROM TenderTypes with (NoLock) WHERE ExtTenderCode = " & l_code
        dst = MyCommon.LRT_Select
        If (dst.Rows.Count > 0) Then
          infoMessage = Copient.PhraseLib.Lookup("tender.codeused", LanguageID)
        Else
          MyCommon.QueryStr = "INSERT INTO TenderTypes with (RowLock) (Description, ExtTenderCode) VALUES (N'" & l_name & "', '" & l_code & "')"
          MyCommon.LRT_Execute()
          MyCommon.Activity_Log(22, l_OCID, AdminUserID, Copient.PhraseLib.Lookup("history.tender-create", LanguageID))
        End If
      End If
    End If
  ElseIf (Request.QueryString("delete") <> "") Then
    l_OCID = MyCommon.Extract_Val(Request.QueryString("tenders"))
    MyCommon.QueryStr = "select distinct O.offerid,description,O.prodenddate from offers as O with (NoLock) " & _
                        "left join offerconditions as OC with (NoLock) on O.offerid=OC.offerid " & _
                        "left join conditiontendertypes as ct with (NoLock) on oc.conditionid=ct.conditionid " & _
                        "where O.deleted=0 and ct.tendertypeid=" & l_OCID
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
      infoMessage = Copient.PhraseLib.Lookup("tender.NotDeletable", LanguageID)
    Else
      MyCommon.QueryStr = "DELETE FROM TenderTypes with (RowLock) WHERE TenderTypeID = " & l_OCID
      MyCommon.LRT_Execute()
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
          MyCommon.QueryStr = "SELECT TenderTypeID, ExtTenderCode, Description FROM TenderTypes with (NoLock) order by ExtTenderCode"
          dst = MyCommon.LRT_Select
          For Each row In dst.Rows
            Send("        <option value=""" & row.Item("TenderTypeID") & """>" & row.Item("ExtTenderCode") & " - " & row.Item("Description") & "</option>")
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
        <label for="code"><% Sendb(Copient.PhraseLib.Lookup("term.code", LanguageID))%>:</label><br />
        <input type="text" id="code" name="code" class="short" value="" maxlength="4" /><br />
        <label for="name"><% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>:</label><br />
        <input type="text" id="name" name="name" class="medium" value="" maxlength="100" />
        <input type="submit" class="regular" id="add" name="add" value="<% Sendb(Copient.PhraseLib.Lookup("term.add", LanguageID)) %>" /><br />
        <hr class="hidden" />
      </div>
      <div class="box" id="tenderdel">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("tender.delete", LanguageID))%>
          </span>
        </h2>
        <select class="longest" id="tenders" name="tenders" size="15">
          <%
            MyCommon.QueryStr = "SELECT TenderTypeID, ExtTenderCode, Description FROM TenderTypes with (NoLock) order by ExtTenderCode"
            dst = MyCommon.LRT_Select
            For Each row In dst.Rows
              Send("          <option value=""" & row.Item("TenderTypeID") & """>" & row.Item("ExtTenderCode") & " - " & row.Item("Description") & "</option>")
            Next
          %>
        </select>
        <br />
        <input type="submit" class="regular" id="delete" name="delete" onclick="if(confirm('<% Sendb(Copient.PhraseLib.Lookup("confirm.delete", LanguageID)) %>')){}else{return false}" value="<% Sendb(Copient.PhraseLib.Lookup("term.delete", LanguageID)) %>" /><br />
      </div>
    </div>
    <br clear="all" />
  </div>
</form>
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
