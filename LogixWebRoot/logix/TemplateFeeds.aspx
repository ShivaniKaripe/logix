<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
  Dim CopientFileName As String = "TemplateFeeds.aspx"
  Dim CopientFileVersion As String = "7.3.1.138972" 
  Dim CopientProject As String = "Copient Logix"
  Dim CopientNotes As String = ""
  
  Dim AdminUserID As Long
  'Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  
  MyCommon.AppName = "TemplateFeeds.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  If (Request.QueryString("PageName") <> "") Then
    Response.Expires = 0
    Response.Clear()
    Response.ContentType = "text/html"
    GenerateFieldList(Request.QueryString("OfferID"), Request.QueryString("PageName"), Boolean.Parse(Request.QueryString("PageEditable")))
  Else
    Send("<b>" & Copient.PhraseLib.Lookup("feeds.noarguments", LanguageID) & "!</b>")
  End If
  Response.Flush()
  Response.End()
%>

<script runat="server">
  Public DefaultLanguageID
  Public MyCommon As New Copient.CommonInc
  
  Sub GenerateFieldList(ByVal OfferID As String, ByVal PageName As String, ByVal PageEditable As Boolean)
    Dim rst As DataTable
    Dim row As DataRow
    Dim i As Integer
    Dim CheckedAttr As String = ""
    Dim Status As String = ""
    
    Response.Cache.SetCacheability(HttpCacheability.NoCache)
    
    Try
      
            MyCommon.QueryStr = "select UI.FieldID, UI.FieldName, TFP.Editable from UIFields UI with (NoLock) " & _
                                "left join TemplateFieldPermissions TFP with (NoLock) on TFP.FieldID = UI.FieldID and TFP.OfferID = " & OfferID & " " & _
                                "where UI.PageName = '" & PageName & "' Order By UI.FieldName;"
            rst = MyCommon.LRT_Select
      If (rst.Rows.Count > 0) Then
        Send("<div id=""templatefields"">")
        Send("<table id=""tblTempFields"" style=""width:100%;"" summary="""">")
        'For i = 1 To 20
        Send("  <tr>")
        Send("    <td colspan=""3""><b>" & Copient.PhraseLib.Lookup("temp.fieldlevelperms", LanguageID) & "</b></td>")
        Send("  </tr>")
        For Each row In rst.Rows
          ' determine whether the row should be checked 
          If (IsDBNull(row.Item("Editable"))) Then ' No record in the TFP table, thus not set
            CheckedAttr = ""
            Status = Copient.PhraseLib.Lookup("term." & IIf(PageEditable, "unlocked", "locked"), LanguageID)
          Else
            If (PageEditable And Not row.Item("Editable")) Then  'Page is editable but field is not
              CheckedAttr = " checked=""checked"""
              Status = Copient.PhraseLib.Lookup("term.locked", LanguageID)
            ElseIf (Not PageEditable And row.Item("Editable")) Then 'Field is editable but page is not
              CheckedAttr = " checked=""checked"""
              Status = Copient.PhraseLib.Lookup("term.unlocked", LanguageID)
            Else
              CheckedAttr = ""
              Status = Copient.PhraseLib.Lookup("term.locked", LanguageID)
            End If
          End If
          Send("  <tr style=""background-color:#e0e0e0;"">")
          Sendb("    <td><input type=""checkbox"" id=""chkTempField-" & MyCommon.NZ(row.Item("FieldID"), "") & """ name=""chkTempField""")
          Send(" value=""" & MyCommon.NZ(row.Item("FieldID"), "") & """" & CheckedAttr & " onclick=""updateLockStatus(this);"" /></td>")
          Send("    <td><label for=""chkTempField-" & MyCommon.NZ(row.Item("FieldID"), "") & """>" & MyCommon.NZ(row.Item("FieldName"), "") & "</label></td>")
          Send("    <td>" & Status & "</td>")
          Send("  </tr>")
        Next
        'Next i
        Send("</table>")
        Send("</div>")
      End If
    Catch ex As Exception
      
    Finally
      MyCommon.Close_LogixRT()
      MyCommon = Nothing
    End Try
  End Sub
</script>
