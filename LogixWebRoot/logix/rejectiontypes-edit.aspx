<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="Copient.CommonInc" %>
<%@ Import Namespace="System.Collections.Generic" %>

<%
  ' *****************************************************************************
  ' * FILENAME:  rejectiontypes-edit.aspx
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
  dim StrTemp As String = ""
  
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

  Dim ProductGroupID As Long
  Dim GName As String
  Dim CreatedDate As String
  Dim LastUpload As String = Nothing
  Dim LastUploadMsg As String = ""
  Dim rstProdTypes As DataTable = Nothing
  Dim tempid As String
  Dim GroupSize As Integer
  Dim outputStatus As Integer
  Dim DefaultIDType As Integer
  Dim File As HttpPostedFile
  Dim InstallPath As String
  Dim rowCount As Integer
  Dim ProdAvailableCount As Integer
  Dim ProdAssignedCount As Integer
  Dim squery As String
  Dim dtProdAvailable As DataTable
  Dim dtProdAssigned As DataTable
  Dim ProductList As String
  Dim Products() As String
  Dim bAdd As Boolean
  Dim bAddAll As Boolean
  Dim bRemove As Boolean
  Dim typeST As DataTable
  Dim rowST As DataRow
  Dim iType As Integer
  Dim deployDate As String
  Dim longDate As New DateTime
  Dim longDateString As String
  Dim statusMessage As String = ""
  Dim ExtProductID As String = ""
  Dim IDLength As Integer = 0
  Dim GNameTitle As String = ""
  Dim GFullName As String = ""
  Dim XID As String = ""
  Dim ShowActionButton As Boolean = False
  Dim ListBoxSize, LinkSize, LabelSize As Integer
  Dim ShowAllItems As Boolean
  Dim OfferCtr As Integer = 0
  Dim IsSpecialGroup As Boolean = False
  Dim CanEditSpecialGroup As Boolean = False
  Dim OfferID As Integer = 0
  Dim EngineID As Integer = -1
  Dim CreatedFromOffer As Boolean = False
  Dim UploadOperation As Integer = 0
  Dim prodDT As DataTable
  Dim HasExcludedNodes, HasExcludedItems As Boolean
  Dim ShowViewSelected As Boolean = False
  Dim rstItems As DataTable = Nothing
  Dim CpeEngineOnly As Boolean = False
  Dim ProductTypeID As Integer = 0
  Dim ItemPKID As Integer = -1
  Dim SelectedOption As String = ""
  Dim iProductIdNumericOnly As Integer = 97
  Dim NewProductGroupID As Long = 0
  Dim HTMLBuf As New StringBuilder()

  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "rejectiontypes-edit.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixXS()
  MyCommon.Open_LogixEX()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  StrTemp = Request.QueryString("ReadOnlyAttribute")
  
  BannersEnabled = (MyCommon.Fetch_SystemOption(66) = "1")

  If (Logix.UserRoles.ViewEditUSAMRejections OrElse Logix.UserRoles.ViewEditCAMRejections) Then
    AllowEditing = True
  End If
  MyCommon.QueryStr = "select top 1 SubTypeID from PromoEngineSubTypes order by SubTypeID DESC;"
  rst = MyCommon.LRT_Select
  If rst.Rows.Count > 0 Then
    MaxEngineSubTypeID = MyCommon.NZ(rst.Rows(0).Item("SubTypeID"), 0)
  End If

  Send_HeadBegin("term.rejectiontypes")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
%>
<style type="text/css">
  td {
    vertical-align: top;
    }
  .editvalue, .savevalue, .cancelvalue {
    font-size: 10px;
    width: 55px;
    }
  .descinput {
    font-size: 12px;
    width: 290px;
    }
  * html .descinput {
    width: 275px;
    }
  #NewValueExtID {
    color: #aaaaaa;
    font-size: 12px;
    width: 75px;
    }
  #NewValueDesc {
    color: #aaaaaa;
    font-size: 12px;
    width: 140px;
    }
  #NewValue {
    font-size: 10px;
    }
</style>
<%
  Send_Scripts()
%>
<script type="text/javascript">

function isValidID() {
    var retVal = true;
    var elemID = document.getElementById("sourcetypeidtxt");
    if (elemID != null) {
        if (elemID.value.length == 0 ) {
            retVal = false;
            alert('<%Sendb(Copient.PhraseLib.Lookup("term.invalid", LanguageID) + " SourceTypeID") %>');
        }
          else if (elemID.value.trim() == '' || /^\d+$/.test(elemID.value) == false) {
            retVal = false;
            alert('<%Sendb(Copient.PhraseLib.Lookup("product.mustbenumeric", LanguageID)) %>');
        }

    }
    return retVal;
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
  document.mainform.submit();
}

function clearDefaultButton(buttonGroup) {
  for (i=0; i < buttonGroup.length; i++) {
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
    Response.Redirect("rejectiontypes-edit.aspx")
  End If

 infoMessage=""
    If(Request.QueryString("description") <>"") then
        If Not (Regex.IsMatch(Request.QueryString("description"), "^[a-z0-9][a-z0-9_\s]+$", RegexOptions.IgnoreCase)) Then
            infoMessage = "Description should not contain special characters"
        End If
   End If

    If (Request.QueryString("addupdate") <> "" andalso infoMessage ="") Then
        MyCommon.QueryStr = "select * from sourcetypes as S with (NoLock) where S.SourceTypeID=" & Request.QueryString("sourcetypeidtxt") & ";"
        rst = MyCommon.LEX_Select()
        If rst.Rows.Count = 0 Then
          MyCommon.QueryStr = "Insert into SourceTypes (SourceTypeID,Description,ActionTypeID) values (" & _
          Request.QueryString("sourcetypeidtxt") & ", '" & Request.QueryString("description") & "'," & Int(Request.QueryString("actiontype")) & ");"
        Else
          MyCommon.QueryStr = "update SourceTypes set SourceTypeID=" & Request.QueryString("sourcetypeidtxt") & ", Description='" & Request.QueryString("description") & _
          "', ActionTypeID=" & Int(Request.QueryString("actiontype")) & " " & _
          "where SourceTypeID=" & Request.QueryString("sourcetypeidtxt") & ";"
        End If
          MyCommon.LEX_Execute()
    ElseIf (Request.QueryString("remove") <> "" AndAlso Request.QueryString("SourceTypeID") <> "") Then
        For i = 0 To Request.QueryString.GetValues("SourceTypeID").GetUpperBound(0)
          MyCommon.QueryStr = "select * from sourcetypes as S with (NoLock) where S.SourceTypeID=" & Request.QueryString.GetValues("SourceTypeID")(i)
          rst = MyCommon.LEX_Select()
          If rst.Rows.Count > 0 Then
            tempid = rst.Rows(0).Item("SourceTypeID")
          End If
          rst = Nothing
          MyCommon.QueryStr = "dbo.pt_SourceTypes_Delete_ByID"
          MyCommon.Open_LEXsp()
          MyCommon.LEXsp.Parameters.Add("@PKID", SqlDbType.BigInt).Value = Request.QueryString.GetValues("SourceTypeID")(i)
          MyCommon.LEXsp.ExecuteNonQuery()
          MyCommon.Close_LEXsp()
          MyCommon.Activity_Log(5, ProductGroupID, AdminUserID, Copient.PhraseLib.Lookup("history.pgroup-remove", LanguageID) & " " & tempid)
        Next
        'MyCommon.QueryStr = "update productgroups with (RowLock) set updatelevel=updatelevel+1, LastUpdate=getdate() where ProductGroupID=" & ProductGroupID
        'MyCommon.LRT_Execute()
        'Response.Status = "301 Moved Permanently"
        Response.AddHeader("Location", "pgroup-edit.aspx?ProductGroupID=" & ProductGroupID)
    End If
    MyCommon.QueryStr = "select Sourcetypeid,Description,Descriptionphraseid,Actiontypeid from sourcetypes order by SourceTypeID asc;"
    rstItems = MyCommon.LEX_Select()
    ListBoxSize = rstItems.Rows.Count
    GroupSize = ListBoxSize
%>

<form action="#" id="mainform" name="mainform">
<div id="intro">
  <h1 id="title">
    <%
      Sendb(Copient.PhraseLib.Lookup("term.rejectiontypes", LanguageID))
    %>
  </h1>
  <div id="controls">
  </div>
</div>
<div id="main">
  <%
    If (infoMessage <> "") Then
      Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")
    End If
  %>
  <div id="column1">
    <div class="box" id="addproducts">
      <h2>
        <span>
          <% Sendb(Copient.PhraseLib.Lookup("rejection.addremoverejectiontype", LanguageID))%>
        </span>
      </h2>
      <br class="half" />

      <span style="position: relative">
        <%
          If (ShowAllItems OrElse GroupSize <= 100) AndAlso rstItems IsNot Nothing Then
            Sendb(Copient.PhraseLib.Lookup("term.total", LanguageID) & " " & Copient.PhraseLib.Lookup("term.Rejection", LanguageID) & _
                " " & Copient.PhraseLib.Lookup("term.types", LanguageID) & ": " & " ( " & rstItems.Rows.Count & " )")
          Else If (ShowAllItems OrElse GroupSize >= 100) AndAlso rstItems IsNot Nothing Then
            Sendb(Copient.PhraseLib.Lookup("term.total", LanguageID) & " " & Copient.PhraseLib.Lookup("term.Rejection", LanguageID) & _
                " " & Copient.PhraseLib.Lookup("term.types", LanguageID) & ": " & " ( " & rstItems.Rows.Count & " )" & ", " & _
                Copient.PhraseLib.Lookup("customer-supplemental-edit.First100Shown", LanguageID).ToString.ToLower & "<br />")
          Else
          End If
        %>
      </span>
      <div id="itemsDiv" onscroll="handlePageClick(this);" class="boxscroll">
        <select name="SourceTypeID" id="SourceTypeID" size="<% Sendb(ListBoxSize)%>" onchange="enableRemove();" multiple="multiple" style="overflow: hidden; <% if (rstItems IsNot Nothing AndAlso rstItems.Rows.Count = 0) then Sendb("visibility:hidden;") %>">
          <%
            Dim descriptionItem As String
            If (GroupSize > 0) Then
                  For Each row In rstItems.Rows
                      If Not IsDBNull(row.Item("Descriptionphraseid")) Then
                          descriptionItem = MyCommon.NZ(row.Item("SourceTypeID"), " ") & "-" & Copient.PhraseLib.Lookup(row.Item("Descriptionphraseid"), LanguageID)
                      Else
                          descriptionItem = MyCommon.NZ(row.Item("SourceTypeID"), " ") & "-" & MyCommon.NZ(row.Item("Description"), " ")
                      End If
                       Send("     <option value=""" & row.Item("SourceTypeID") & """>" & descriptionItem & "</option>")
                      
                  Next
            End If
          %>
        </select>
      </div>
      <%
        If (Not ShowAllItems AndAlso GroupSize > 100) Then
          Send("<input class=""regular"" id=""btnShowAll"" name=""btnShowAll"" type=""button"" value=""" & Copient.PhraseLib.Lookup("list.showall", LanguageID) & """ onclick=""submitShowAll();"" />")
        End If%>
      <br />
      <table cellpadding="1" cellspacing="1">
        <tr>
          <td><% Sendb(Copient.PhraseLib.Lookup("term.sourcetypeid", LanguageID))%>:</td>
          <td><% Sendb(Copient.PhraseLib.Lookup("term.action", LanguageID) & " ")%> <% Sendb(Copient.PhraseLib.Lookup("term.type", LanguageID))%>:</td>
        </tr>
        <tr>
          <td><input type="text" id="sourcetypeidtxt" name="sourcetypeidtxt" maxlength="9" style="<%Sendb(IIf(CreatedFromOffer, "width:115px;", "width:137px;")) %>" value="" /></td>
          <td>
            <select id="actiontype" name="actiontype" style="width:175px;">
            <%
                MyCommon.QueryStr = "select Distinct A.ActionTypeID,A.Description,A.PhraseID from ActionTypes A  with (NoLock)"
              rst2 = MyCommon.LEX_Select
              For Each row2 In rst2.Rows
                    Send("     <option value=""" & row2.Item("ActionTypeID") & """>" & Copient.PhraseLib.Lookup(row2.Item("PhraseID"), LanguageID) & "</option>")
              Next
            %>
            </select>
          </td>
        </tr>
      </table>

      <div>
        <label for="description"><% Sendb(Copient.PhraseLib.Lookup("term.description", LanguageID))%>:</label><br />
        <input type="text" id="description" style="<%Sendb(IIf(CreatedFromOffer, "width:310px;", "width:347px;")) %>" name="description" maxlength="200" value="" /><br />
        <br class="half" />
      </div>
      <%
        'If (Logix.UserRoles.ViewEditUSAMRejections OrElse Logix.UserRoles.ViewEditCAMRejections) Then
          Sendb("    <input type=""submit"" class=""large"" id=""addupdate"" name=""addupdate"" value=""" & Copient.PhraseLib.Lookup("term.add", LanguageID) & "\" & Copient.PhraseLib.Lookup("term.update", LanguageID) & """ onclick=""return isValidID();"" />")
          Sendb("    <input type=""submit"" disabled=""disabled"" class=""large"" id=""remove"" name=""remove"" onclick=""if(confirm('" & Copient.PhraseLib.Lookup("confirm.remove", LanguageID) & "')){}else{return false}"" style=""width:150px;"" value=""" & Copient.PhraseLib.Lookup("term.remove", LanguageID) & """ /><br />")
        'End If 
     %>
      <hr class="hidden" />
    </div>
  </div>
  
  <div id="gutter">
  </div>
  
  <br clear="all" />
</div>
</form>

<script type="text/javascript">
    function enableRemove() {
        document.getElementById("remove").disabled = false;
    }
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
  MyCommon.Close_LogixEX()
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixXS()
  Send_BodyEnd("mainform", "ExtID")
  MyCommon = Nothing
  Logix = Nothing
%>
