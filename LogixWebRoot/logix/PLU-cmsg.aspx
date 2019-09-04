<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: PLU-cmsg.aspx 
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
  Dim Localization As Copient.Localization
  Dim rst As DataTable
  Dim row As DataRow
  Dim rst2 As DataTable
  Dim dt As DataTable
  Dim row2 As DataRow
  Dim Line1 As String = ""
  Dim Line2 As String = ""
  Dim TierLine1 As String = ""
  Dim TierLine2 As String = ""
  Dim TierLine2Tag As String = ""
  Dim LineLength As Integer = 0
  Dim Line2Text As String = ""
  Dim Line2Tag As String = ""
  Dim HasTag As Boolean = False
  Dim TagStart As Integer = 0
  Dim TierBeep As Integer = 0
  Dim TierBeepDuration As Integer = 0
  Dim MessageID As Long = 0
  Dim MsgAdded As Boolean = False
  Dim bIsErrorMsg As Boolean = False
  Dim DisabledAttribute As String = ""
  Dim CloseAfterSave As Boolean = False
  Dim Beep As Integer = 0
  Dim BeepDuration As Integer = 1
  Dim BeepDurDisplay As String = "none"
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim EngineID As Integer = 2
  Dim BannersEnabled As Boolean = True
  Dim TierLevels As Integer = 1
  Dim t As Integer = 1
  Dim l = 1
  Dim ValidTiers As Boolean = True
  Dim DisplayImmediate As Integer = 0
  Dim Display As Boolean = False
  Dim LanguagesDT As DataTable
  Dim MLI As New Copient.Localization.MultiLanguageRec
  Dim MultiLanguageEnabled As Boolean = False
  Dim DefaultLanguageID As Integer = 0
  Dim PKID As Integer = 0
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "PLU-cmsg.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  Localization = New Copient.Localization(MyCommon)

  MultiLanguageEnabled = IIf(MyCommon.Fetch_SystemOption(124) = "1", True, False)
  Integer.TryParse(MyCommon.Fetch_SystemOption(1), DefaultLanguageID)
  If DefaultLanguageID = 0 Then DefaultLanguageID = 1
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")
  
  EngineID = 2
  Line1 = Left(Trim(Request.QueryString("t1_line1_" & DefaultLanguageID)), 20)
  Beep = MyCommon.Extract_Val(Request.QueryString("t1_beep"))
  BeepDuration = MyCommon.Extract_Val(Request.QueryString("t1_beepDuration"))
  
  ' If there's an existing PLU message, find its ID; otherwise leave it 0
  MyCommon.QueryStr = "select MessageID from CPE_CashierMessages with (NoLock) where PLU=1;"
  rst = MyCommon.LRT_Select
  If rst.Rows.Count > 0 Then
    MessageID = MyCommon.NZ(rst.Rows(0).Item("MessageID"), 0)
  End If
  
  If (Request.QueryString("save") <> "") And (MessageID = 0) Then
    ' New PLU cashier message
    MyCommon.QueryStr = "insert into CPE_CashierMessages (LastUpdate, PLU) values (getdate(), 1);"
    MyCommon.LRT_Execute()
    MyCommon.QueryStr = "select MessageID from CPE_CashierMessages where PLU=1;"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count > 0 Then
      MessageID = MyCommon.NZ(rst.Rows(0).Item("MessageID"), 0)
    End If
    Create_MessageTiers(MessageID, DisplayImmediate, 1, DefaultLanguageID)
    MyCommon.Activity_Log(3, 0, AdminUserID, Copient.PhraseLib.Lookup("history.plucmsg-create", LanguageID))
    CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
  ElseIf (Request.QueryString("save") <> "") And (MessageID > 0) Then
    ' Update existing PLU cashier message
    MyCommon.QueryStr = "update CPE_CashierMessages with (RowLock) set LastUpdate=getDate() " & _
                        "where MessageID=" & MessageID & ";"
    MyCommon.LRT_Execute()
    MyCommon.QueryStr = "select PKID from CPE_CashierMessageTiers with (NoLock) where MessageID=" & MessageID & " and TierLevel=1;"
    rst = MyCommon.LRT_Select
    If rst.Rows.Count = 0 Then
      Create_MessageTiers(MessageID, DisplayImmediate, 1, DefaultLanguageID)
    Else
      MyCommon.QueryStr = "update CPE_CashierMessageTiers set Line1=N'" & Line1 & "', Beep=" & Beep & ", BeepDuration=" & BeepDuration & " " & _
                          "where MessageID=" & MessageID & " and TierLevel=1;"
      MyCommon.LRT_Execute()
      ' Clear and re-save the translations
      If MultiLanguageEnabled Then
        MyCommon.QueryStr = "select PKID from CPE_CashierMessageTiers where MessageID=" & MessageID & ";"
        rst = MyCommon.LRT_Select
        If rst.Rows.Count > 0 Then
          PKID = MyCommon.NZ(rst.Rows(0).Item("PKID"), 0)
        End If
        MyCommon.QueryStr = "delete from CPE_CashierMsgTranslations where CashierMsgTierID=" & PKID & ";"
        MyCommon.LRT_Execute()
        MyCommon.QueryStr = "SELECT LanguageID FROM Languages WHERE (InstalledForUI = 1 Or AvailableForCustFacing = 1) " & _
                            "ORDER BY CASE WHEN LanguageID=" & DefaultLanguageID & " THEN 1 ELSE 2 END, LanguageID;"
        LanguagesDT = MyCommon.LRT_Select
        For Each row In LanguagesDT.Rows
          Line1 = Left(Trim(Request.QueryString("t1_line1_" & row.Item("LanguageID"))), 20)
          MyCommon.QueryStr = "insert into CPE_CashierMsgTranslations (CashierMsgTierID, LanguageID, Line1, Line2) " & _
                              "values (" & PKID & ", " & row.Item("LanguageID") & ", '" & Line1 & "', '');"
          MyCommon.LRT_Execute()
        Next
      End If
    End If
    MyCommon.Activity_Log(3, 0, AdminUserID, Copient.PhraseLib.Lookup("history.plucmsg-edit", 1))
    CloseAfterSave = (MyCommon.Fetch_SystemOption(48) = "1")
  End If
  
  MyCommon.QueryStr = "select CM.MessageID, CMT.Line1, CMT.Line2, CMT.Beep, CMT.BeepDuration " & _
                      "from CPE_CashierMessages CM with (NoLock) " & _
                      "inner join CPE_CashierMessageTiers CMT with (NoLock) on CMT.MessageID=CM.MessageID " & _
                      "where CM.MessageID=" & MessageID & ";"
  rst = MyCommon.LRT_Select
  If (rst.Rows.Count > 0) Then
    Line1 = MyCommon.NZ(rst.Rows(0).Item("Line1"), "")
    Line2 = MyCommon.NZ(rst.Rows(0).Item("Line2"), "")
    MessageID = MyCommon.NZ(rst.Rows(0).Item("MessageID"), 0)
    Beep = MyCommon.NZ(rst.Rows(0).Item("Beep"), 0)
    BeepDuration = MyCommon.NZ(rst.Rows(0).Item("BeepDuration"), 0)
  End If
  
  DisabledAttribute = IIf(Logix.UserRoles.EditSystemConfiguration, "", " disabled=""disabled""")
  
  Send_HeadBegin("term.triggercode")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
%>
<script type="text/javascript">
function cleanMessage() {
  var elem = document.getElementById("t1_line1_<%Sendb(DefaultLanguageID)%>");
  
  if (elem != null) {
    elem.value = elem.value.replace("<", "\1");
  }
  
  return true;
}

function beepTypeChanged(elem, t) {
  var BEEP_DURATION_VALUE = 3;
  var elemDurationRow = document.getElementById("t" + t + "_trDuration");
  var elemDurText = document.getElementById("t" + t + "_beepDuration");
  
  if (elem != null && elemDurationRow != null) {
    if (elem.options[elem.selectedIndex].value == BEEP_DURATION_VALUE) {
      elemDurationRow.style.display = "";
      elemDurText.focus();
      elemDurText.select();
    } else {
      elemDurationRow.style.display = "none";
      elemDurText.value = "";
    }
  }
}

function isDangerousString(obj)  {
    var stringToCheck = obj.value;
	var savebutton = document.getElementById("save")
         if ((stringToCheck.indexOf("<") > -1) || (stringToCheck.indexOf(">") > -1))  {
          alert('<% Sendb(Copient.PhraseLib.Lookup("categories.invalidname", LanguageID))%>');
		  obj.focus();
		  savebutton.disabled = true;
		  return false;
		} else {
		   savebutton.disabled = false;
           return true;
        }		 
  }
</script>
<%
  Send_HeadEnd()
  Send_BodyBegin(2)
  If (Logix.UserRoles.EditSystemConfiguration = False) Then
    Send_Denied(2, "perm.admin-configuration")
    GoTo done
  End If
%>
<form action="PLU-cmsg.aspx" id="mainform" name="mainform">
  <div id="intro">
    <%
      Send("<h1 id=""title"">" & Copient.PhraseLib.Lookup("term.triggercode", LanguageID) & " " & StrConv(Copient.PhraseLib.Lookup("term.cashiermessage", LanguageID), VbStrConv.Lowercase) & "</h1>")
      Send("<div id=""controls"">")
      If Logix.UserRoles.EditSystemConfiguration Then
        Send_Save()
      End If
      Send("</div>")
    %>
  </div>
  <div id="main">
    <%
      If MessageID = 0 Then
        infoMessage = Copient.PhraseLib.Lookup("plu.CashierMessageNotDefined", LanguageID)
      End If
      If (infoMessage <> "" And bIsErrorMsg) Then
        Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
      ElseIf (infoMessage <> "") Then
        Send("<div id=""infobar"" class=""green-background"">" & infoMessage & "</div>")
      End If
    %>
    <div id="column1">
      <div class="box" id="message">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.message", LanguageID))%>
          </span>
        </h2>
        <%
          MyCommon.QueryStr = "select TierLevel, Line1, Line2, Beep, BeepDuration from CPE_CashierMessageTiers with (NoLock) " & _
                              "where MessageID=" & MessageID & " order by TierLevel;"
          rst = MyCommon.LRT_Select
          
          Dim TierRecordDT As DataTable
          For t = 1 To TierLevels
            MyCommon.QueryStr = "select PKID, TierLevel, Line1, Line2, Beep, BeepDuration from CPE_CashierMessageTiers with (NoLock) " & _
                                "where MessageID=" & MessageID & " and TierLevel=" & t & ";"
            TierRecordDT = MyCommon.LRT_Select
            If TierRecordDT.Rows.Count > 0 Then
              PKID = MyCommon.NZ(TierRecordDT.Rows(0).Item("PKID"), 0)
            Else
              PKID = 0
            End If
            
            If TierLevels > 1 Then
              Send("<label for=""t" & t & "_line1"" style=""position:relative;""><b>" & Copient.PhraseLib.Lookup("term.tier", LanguageID) & " " & t & ":</b></label>")
              Send("<br />")
            End If
            
            ' Beep
            Send("<select id=""t" & t & "_beep"" name=""t" & t & "_beep"" style=""float:left;position:relative;"" onchange=""beepTypeChanged(this, " & t & ");""" & DisabledAttribute & ">")
            MyCommon.QueryStr = "select BeepTypeID, PhraseID from BeepTypes BT with (NoLock);"
            rst2 = MyCommon.LRT_Select
            For Each row2 In rst2.Rows
              Sendb("  <option value=""" & MyCommon.NZ(row2.Item("BeepTypeID"), 0) & """")
              If TierRecordDT.Rows.Count > 0 Then
                If (MyCommon.NZ(row2.Item("BeepTypeID"), 0) = MyCommon.NZ(TierRecordDT.Rows(0).Item("Beep"), 0)) Then
                  Sendb(" selected=""selected""")
                End If
              End If
              Sendb(">" & Copient.PhraseLib.Lookup(row2.Item("PhraseID"), LanguageID))
              Send("</option>")
            Next
            Send("</select>")
            
            ' Beep duration
            If TierRecordDT.Rows.Count > 0 Then
              BeepDurDisplay = IIf(MyCommon.NZ(TierRecordDT.Rows(0).Item("Beep"), 0) = 3, "inline", "none")
            End If
            Send("<span id=""t" & t & "_trDuration"" style=""display:" & BeepDurDisplay & "; float:left; position:relative;"">")
            Sendb("  :&nbsp;<input type=""text"" class=""shortest"" id=""t" & t & "_beepDuration"" name=""t" & t & "_beepDuration"" maxlength=""2""")
            If TierRecordDT.Rows.Count = 0 Then
              Send(" value=""0""" & DisabledAttribute & " />")
            Else
              Send(" value=""" & MyCommon.NZ(TierRecordDT.Rows(0).Item("BeepDuration"), 0) & """" & DisabledAttribute & " />")
            End If
            Send("</span>")
            Send("<br clear=""left"" />")
            Send("<br class=""half"" />")

            l = 1
            MyCommon.QueryStr = "SELECT L.LanguageID, L.Name, L.MSNetCode, L.JavaLocaleCode, L.PhraseTerm, L.RightToLeftText, T.Line1, T.Line2 " & _
                                "FROM Languages AS L " & _
                                "LEFT JOIN CPE_CashierMsgTranslations AS T ON T.LanguageID=L.LanguageID AND T.CashierMsgTierID=" & PKID & " " & _
                                "WHERE (InstalledForUI = 1 Or AvailableForCustFacing = 1) " & _
                                "ORDER BY CASE WHEN L.LanguageID=" & DefaultLanguageID & " THEN 1 ELSE 2 END, L.LanguageID;"
            LanguagesDT = MyCommon.LRT_Select
            For Each row In LanguagesDT.Rows
              Dim MLLanguageCode As String = MyCommon.NZ(row.Item("MSNetCode"), "")
              Dim MLLanguageName As String = Copient.PhraseLib.Lookup(MyCommon.NZ(row.Item("PhraseTerm"), ""), MyCommon.GetAdminUser.LanguageID)
              Dim MLLanguageID As Integer = MyCommon.NZ(row.Item("LanguageID"), 0)
              
              If (MultiLanguageEnabled = True) Or (MultiLanguageEnabled = False AndAlso MyCommon.NZ(row.Item("LanguageID"), 0) = DefaultLanguageID) Then
                Send("<label for=""t" & t & "_line1_" & MLLanguageID & """>" & MLLanguageName & IIf(MLLanguageID = DefaultLanguageID, " (" & Copient.PhraseLib.Lookup("term.default", MyCommon.GetAdminUser.LanguageID) & ")", "") & ":</label><br />")
                ' Line 1 input
                Dim Line1Raw As String = ""
                Sendb("<input type=""text"" id=""t" & t & "_line1_" & MLLanguageID & """ name=""t" & t & "_line1_" & MLLanguageID & """ onfocus=""srcElement=this"" style=""font-family:Courier; width:200px;"" maxlength=""20"" onblur=""javascript:return isDangerousString(this);""")
                If TierRecordDT.Rows.Count = 0 Then
                  Send(" value=""""" & DisabledAttribute & " />")
                Else
                  If l = 1 Then
                    Line1Raw = MyCommon.NZ(TierRecordDT.Rows(0).Item("Line1"), "")
                  Else
                    Line1Raw = MyCommon.NZ(row.Item("Line1"), "")
                  End If
                  Send(" value=""" & IIf(Line1Raw = "", "", Line1Raw.Replace("""", "&quot;")) & """" & DisabledAttribute & " />")
                End If
                Send("<br />")
                ' Line 2 input
                Send("<input type=""text"" id=""t" & t & "_line2"" name=""t" & t & "_line2"" style=""font-family:Courier; width:200px;"" maxlength=""20"" value=""(Trigger code)"" disabled=""disabled"" />")
                Send("<br />")
                Send("<br class=""half"" />")
              End If
              l += 1
            Next
            If MultiLanguageEnabled And TierLevels > 1 And t < TierLevels Then
              Send("<hr />")
            End If
          Next
          
          'Display Selector
          Send("<input type=""hidden"" id=""DisplayImmediate"" name=""DisplayImmediateGrp"" value=""0"" />")
        %>
      </div>
    </div>
    
    <div id="gutter">
    </div>
    
    <div id="column2">
    </div>
  </div>
</form>

<script runat="server">
  Function Create_Message(ByVal OfferID As String, ByVal Line1 As String, ByVal Line2 As String, ByVal Line2Tag As String, ByVal Phase As Integer, ByVal TpROID As Integer, ByRef DeliverableID As Long) As Boolean
    Dim MyCommon As New Copient.CommonInc
    Dim Status As Integer = 0
    
    Try
      MyCommon.QueryStr = "dbo.pa_CPE_AddCashierMessage"
      MyCommon.Open_LogixRT()
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt, 4).Value = OfferID
      MyCommon.LRTsp.Parameters.Add("@TpROID", SqlDbType.Int, 4).Value = TpROID
      MyCommon.LRTsp.Parameters.Add("@Phase", SqlDbType.Int, 4).Value = Phase
      MyCommon.LRTsp.Parameters.Add("@Status", SqlDbType.Int, 4).Direction = ParameterDirection.Output
      MyCommon.LRTsp.Parameters.Add("@DeliverableID", SqlDbType.Int, 4).Direction = ParameterDirection.Output
      MyCommon.LRTsp.ExecuteNonQuery()
      Status = MyCommon.LRTsp.Parameters("@Status").Value
      DeliverableID = MyCommon.LRTsp.Parameters("@DeliverableID").Value
      MyCommon.Close_LRTsp()
    Catch ex As Exception
      Status = -1
      Send(ex.ToString)
    Finally
      MyCommon.Close_LogixRT()
      MyCommon = Nothing
    End Try
    
    Return (Status = 0)
  End Function
  
  Sub Create_MessageTiers(ByVal MessageID As Long, ByVal DisplayImmediate As Integer, ByVal TierLevel As Long, Optional ByVal DefaultLanguageID As Integer = 1)
    Dim MyCommon As New Copient.CommonInc
    Dim Localization As Copient.Localization
    Dim MLI As New Copient.Localization.MultiLanguageRec
    
    Dim Line1 As String = Request.QueryString("t" & TierLevel & "_line1_" & DefaultLanguageID)
    Dim Line2 As String = Request.QueryString("t" & TierLevel & "_line2_" & DefaultLanguageID)
    Dim Beep As Integer = MyCommon.Extract_Val(Request.QueryString("t" & TierLevel & "_beep"))
    Dim BeepDuration As Integer = MyCommon.Extract_Val(Request.QueryString("t" & TierLevel & "_beepduration"))
    
    Dim Line1Clean As String = ""
    Dim Line2Clean As String = ""
    Dim PKID As Integer = 0
    
    Localization = New Copient.Localization(MyCommon)
    MyCommon.QueryStr = "dbo.pa_CPE_AddCashierMessageTiers"
    
    MyCommon.Open_LogixRT()
    MyCommon.Open_LRTsp()
    
    If Line1 <> "" Then Line1Clean = Replace(Line1, "|", "")
    If Line2 <> "" Then Line2Clean = Replace(Line2, "|", "")
    
    MyCommon.LRTsp.Parameters.Add("@MessageID", SqlDbType.Int, 4).Value = MessageID
    MyCommon.LRTsp.Parameters.Add("@TierLevel", SqlDbType.Int, 4).Value = TierLevel
    MyCommon.LRTsp.Parameters.Add("@Line1", SqlDbType.NVarChar, 30).Value = Line1Clean
    MyCommon.LRTsp.Parameters.Add("@Line2", SqlDbType.NVarChar, 30).Value = Line2Clean
    MyCommon.LRTsp.Parameters.Add("@Beep", SqlDbType.Int, 4).Value = Beep
    MyCommon.LRTsp.Parameters.Add("@BeepDuration", SqlDbType.Int, 4).Value = BeepDuration
    MyCommon.LRTsp.Parameters.Add("@DisplayImmediate", SqlDbType.Bit).Value = DisplayImmediate
    MyCommon.LRTsp.Parameters.Add("@PKID", SqlDbType.Int).Direction = ParameterDirection.Output
    MyCommon.LRTsp.ExecuteNonQuery()
    PKID = MyCommon.LRTsp.Parameters("@PKID").Value
    MyCommon.Close_LRTsp()
    MyCommon.Close_LogixRT()
    
    'Save multilanguage values
    If (MyCommon.Fetch_SystemOption(124) = "1") Then
      Dim LanguagesDT As DataTable
      Dim row As DataRow
      MyCommon.Open_LogixRT()
      MyCommon.QueryStr = "SELECT LanguageID FROM Languages WHERE (InstalledForUI = 1 Or AvailableForCustFacing = 1) " & _
                          "ORDER BY CASE WHEN LanguageID=" & DefaultLanguageID & " THEN 1 ELSE 2 END, LanguageID;"
      LanguagesDT = MyCommon.LRT_Select
      For Each row In LanguagesDT.Rows
        Line1Clean = ""
        Line2Clean = ""
        Line1 = Request.QueryString("t" & TierLevel & "_line1_" & MyCommon.NZ(row.Item("LanguageID"), 0))
        Line2 = Request.QueryString("t" & TierLevel & "_line2_" & MyCommon.NZ(row.Item("LanguageID"), 0))
        If Line1 <> "" Then Line1Clean = Replace(Line1, "|", "")
        If Line2 <> "" Then Line2Clean = Replace(Line2, "|", "")
        If Line1Clean <> "" OrElse Line2Clean <> "" Then
          MyCommon.QueryStr = "INSERT INTO CPE_CashierMsgTranslations (CashierMsgTierID, LanguageID, Line1, Line2) " & _
                              "VALUES (" & PKID & ", " & row.Item("LanguageID") & ", '" & Line1Clean & "', '" & Line2Clean & "');"
          MyCommon.LRT_Execute()
        End If
      Next
      MyCommon.Close_LogixRT()
      MyCommon = Nothing
    End If
  End Sub
</script>
<%
  If (CloseAfterSave) Then
    Send("<script type=""text/javascript"">")
    Send("  window.close();")
    Send("</script>")
  End If
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd("mainform", "t1_line1")
  Logix = Nothing
  MyCommon = Nothing
%>
