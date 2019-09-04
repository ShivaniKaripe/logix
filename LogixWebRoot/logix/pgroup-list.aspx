<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>

<%  ' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>
<%
    ' *****************************************************************************
    ' * FILENAME: pgroup-list.aspx 
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
    Dim row As System.Data.DataRow
    Dim Shaded As String = "shaded"
    Dim idNumber As Integer
    Dim idSearch As String
    Dim idSearchText As String
    Dim PageNum As Integer = 0
    Dim MorePages As Boolean
    Dim linesPerPage As Integer = 20
    Dim sizeOfData As Integer
    Dim i As Integer = 0
    Dim PrctSignPos As Integer
    Dim infoMessage As String = ""
    Dim Handheld As Boolean = False
    Dim IsSpecialGroup As Boolean = False
    Dim HideSpecialGroups As Boolean = False
  Dim AdvSearchSQL As String = ""
  Dim CriteriaMsg As String = ""
  Dim CriteriaTokens As String = ""
  Dim CriteriaError As Boolean = False
  Dim WhereClause As String = ""
  Dim WhereBuf As New StringBuilder()
    Dim bStoreUser As Boolean = False
    Dim sValidLocIDs As String = ""
    Dim sValidSU As String = ""
    Dim wherestr As String = "" 
    Dim sJoin As String = "" 
    Dim iLen As Integer = 0
    Dim rst As DataTable
  Dim bEnableRestrictedAccessToUEOfferBuilder As Boolean = IIf(MyCommon.Fetch_SystemOption(249)="1",True,False)            
  Dim bStaticEnabled As Boolean = False
  Dim bCreateStaticPG As Boolean = False

  
    If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
        Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
    End If
  
    Response.Expires = 0
    MyCommon.AppName = "pgroup-list.aspx"
    MyCommon.Open_LogixRT()
    AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
    bStaticEnabled = MyCommon.Fetch_SystemOption(280)
    If bStaticEnabled Then
      bCreateStaticPG = Logix.UserRoles.CreateStaticProductGroups
    End If

    'Store User
    If(MyCommon.Fetch_CM_SystemOption(131) = "1") Then
      MyCommon.QueryStr = "select LocationID from StoreUsers where UserID=" & AdminUserID & ";"
      rst = MyCommon.LRT_Select
      iLen = rst.Rows.Count
      If iLen > 0 Then
        bStoreUser = True
        sValidSU = AdminUserID
        For i=0 to (iLen-1)
          If i=0 Then 
            sValidLocIDs = rst.Rows(0).Item("LocationID")
          Else 
            sValidLocIDs &= "," & rst.Rows(i).Item("LocationID")
          End If
        Next
      
        MyCommon.QueryStr = "select UserID from StoreUsers where LocationID in (" & sValidLocIDs & ") and NOT UserID=" & AdminUserID & ";"
        rst = MyCommon.LRT_Select
        iLen = rst.Rows.Count
        If iLen > 0 Then
          For i=0 to (iLen-1)
            sValidSU &= "," & rst.Rows(i).Item("UserID") 
          Next
        End If
      End If
    End If
  
    PageNum = Request.QueryString("pagenum")
    If PageNum < 0 Then PageNum = 0
    MorePages = False
  
    Send_HeadBegin("term.productgroups")
    Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
    Send_Metas()
    Send_Links(Handheld)
    Send_Scripts()
%>
<script type="text/javascript">
  function launchAdvSearch() {
    self.name = "PGroupListWin";
    <%
      Send("openPopup(""advanced-search.aspx?PGroupList=1"");")
    %>
  }
  function editSearchCriteria() {
    var tokenStr = document.frmIter.advTokens.value;
    
    self.name = "PGroupListWin";
    <%
      Send("openPopup(""advanced-search.aspx?PGroupList=1&amp;tokens="" + tokenStr);")
    %>
  }
</script>
<%
    Send_HeadEnd()
    Send_BodyBegin(1)
    Send_Bar(Handheld)
    Send_Help(CopientFileName)
    Send_Logos()
    Send_Tabs(Logix, 4)
    Send_Subtabs(Logix, 40, 1)
  
    If (Logix.UserRoles.AccessProductGroups = False) Then
        Send_Denied(1, "perm.pgroup-access")
        GoTo done
    End If
  
  ' handle an Advance Search Criteria
  If (Request.Form("mode") = "advancedsearch") Then
    Dim TempStr As String = ""
    Dim CritBuf As New StringBuilder()
    Dim CritTokenBuf As New StringBuilder()
    
    If ( hasOption( "createdby" ) AndAlso Request.Form("createdbyOption") <> "") Then
      WhereBuf.Append(" and CreatedByAdminID IN (select Distinct AdminUserID from AdminUsers where ")
      WhereBuf.Append(GetOptionString(MyCommon, Integer.Parse(Request.Form("createdbyOption")), Request.Form("createdby"), "UserName"))
      WhereBuf.Append(") ")
      If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
      CritBuf.Append(Copient.PhraseLib.Lookup("term.createdby", LanguageID) & " " & GetOptionType(Integer.Parse(Request.Form("createdbyOption"))) & " '" & Request.Form("createdby").Trim & "'")
      CritTokenBuf.Append("CreatedBy," & Integer.Parse(Request.Form("createdbyOption")) & "," & Request.Form("createdby").Trim & ",|")
    End If
    
    If ( hasOption( "lastupdatedby" ) AndAlso Request.Form("lastupdatedbyOption") <> "") Then
      WhereBuf.Append(" and LastUpdatedByAdminID IN (select Distinct AdminUserID from AdminUsers where ")
      WhereBuf.Append(GetOptionString(MyCommon, Integer.Parse(Request.Form("lastupdatedbyOption")), Request.Form("lastupdatedby"), "UserName"))
      WhereBuf.Append(") ")
      If (CritBuf.Length > 0) Then CritBuf.Append(" " & StrConv(Copient.PhraseLib.Lookup("term.and", LanguageID), VbStrConv.Lowercase) & " ")
      CritBuf.Append(Copient.PhraseLib.Lookup("term.lastupdatedby", LanguageID) & " " & GetOptionType(Integer.Parse(Request.Form("lastupdatedbyOption"))) & " '" & Request.Form("lastupdatedby").Trim & "'")
      CritTokenBuf.Append("LastUpdatedBy," & Integer.Parse(Request.Form("lastupdatedbyOption")) & "," & Request.Form("lastupdatedby").Trim & ",|")
    End If
    
    CriteriaMsg &= CritBuf.ToString
    CriteriaTokens = CritTokenBuf.ToString
  End If
  
    Dim SortText As String = "ProductGroupID"
    Dim SortDirection As String
  Dim FilterUser As String
  
    FilterUser = Server.HtmlEncode(Request.QueryString("filterUser"))
  If (FilterUser = "") Then FilterUser = AdminUserID.ToString
  
    If (Server.HtmlEncode(Request.QueryString("SortText")) <> "") Then
        SortText = Request.QueryString("SortText")
    End If
  
    If (Server.HtmlEncode(Request.QueryString("pagenum")) = "") Then
        If (Server.HtmlEncode(Request.QueryString("SortDirection")) = "ASC") Then
            SortDirection = "DESC"
        ElseIf (Server.HtmlEncode(Request.QueryString("SortDirection")) = "DESC") Then
            SortDirection = "ASC"
        Else
            SortDirection = "DESC"
        End If
    Else
        SortDirection = Server.HtmlEncode(Request.QueryString("SortDirection"))
    End If
  
    ' hide the special exclusion groups when UE is the only major promo engine installed.
    HideSpecialGroups = MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.UE) AndAlso Not MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CPE) AndAlso Not MyCommon.IsEngineInstalled(Copient.CommonInc.InstalledEngines.CM)
    
    MyCommon.QueryStr = "pt_ProductGroupList_Select"
    MyCommon.Open_LRTsp()
    MyCommon.LRTsp.Parameters.Add("@adminuserid", SqlDbType.BigInt).Value = AdminUserID
  If (Request.Form("mode") = "advancedsearch" OrElse Request.Form("advSql") <> "") Then
    If (Request.Form("advSql") <> "") Then WhereBuf = New StringBuilder(Server.UrlDecode(Request.Form("advSql")))
    If (Request.Form("advCrit") <> "" AndAlso CriteriaMsg = "") Then CriteriaMsg = Server.UrlDecode(Request.Form("advCrit"))
    If (Request.Form("advTokens") <> "" AndAlso CriteriaTokens = "") Then CriteriaTokens = Server.UrlDecode(Request.Form("advTokens"))
        MyCommon.LRTsp.Parameters.Add("@WhereBuf", SqlDbType.Varchar, 500).Value = WhereBuf.ToString
        AdvSearchSQL = WhereBuf.ToString
    ElseIf (Server.HtmlEncode(Request.QueryString("searchterms")) <> "") Then
        If (Integer.TryParse(Server.HtmlEncode(Request.QueryString("searchterms")), idNumber)) Then
            idSearch = idNumber.ToString
        Else
            idSearch = "-1"
        End If
        idSearchText = MyCommon.Parse_Quotes(Server.HtmlEncode(Request.QueryString("searchterms")))
        PrctSignPos = idSearchText.IndexOf("%")
        If (PrctSignPos > -1) Then
            idSearch = -1
            idSearchText = idSearchText.Replace("%", "[%]")
        End If
        If (idSearchText.IndexOf("_") > -1) Then idSearchText = idSearchText.Replace("_", "[_]")
        If (idSearchText.IndexOf("&amp;") > -1) Then idSearchText = idSearchText.Replace("&amp;", "&")
        MyCommon.LRTsp.Parameters.Add("@idSearchText", SqlDbType.VarChar).Value = MyCommon.Parse_Quotes(Server.HtmlDecode(idSearchText))
        MyCommon.LRTsp.Parameters.Add("@idSearch", SqlDbType.BigInt).Value = Convert.ToInt32(idSearch)
       If(bEnableRestrictedAccessToUEOfferBuilder AndAlso Logix.UserRoles.AccessTranslatedUEOffers) Then 
            MyCommon.LRTsp.Parameters.Add("@FilterTransformedPG", SqlDbType.Bit).Value = False
       End If
    Else
        MyCommon.LRTsp.Parameters.Add("@idSearchText", SqlDbType.VarChar).Value = String.Empty
        MyCommon.LRTsp.Parameters.Add("@idSearch", SqlDbType.VarChar).Value = String.Empty
    End If
    Dim flag As Boolean = False
    If HideSpecialGroups Then
        flag = True
    End If
    MyCommon.LRTsp.Parameters.Add("@HideSpecialGroups", SqlDbType.Bit).Value = flag
    MyCommon.LRTsp.Parameters.Add("@SortText", SqlDbType.VarChar).Value = SortText
    MyCommon.LRTsp.Parameters.Add("@SortDirection", SqlDbType.VarChar).Value = SortDirection
  
    'Response.Write( "Testing query -> " & MyCommon.QueryStr & " <- against database <br />" )	
    Dim checkPGlistBasedOnBuyerID As Boolean = True
    If ((Logix.UserRoles.ViewProductgroupRegardlessBuyer = True And MyCommon.IsEngineInstalled(9) = True)Or(MyCommon.IsEngineInstalled(9) = false)) Then
        checkPGlistBasedOnBuyerID = False
    End If
    MyCommon.LRTsp.Parameters.Add("@checkPGlistBasedOnBuyerID", SqlDbType.Bit).Value = checkPGlistBasedOnBuyerID
    
    'Store Users
    if bStoreUser Then
      sJoin = "Full Outer Join ProductGroupLocUpdate pglu with (NoLock) on PG.ProductGroupID=pglu.ProductGroupID " 
      wherestr = " and (LocationID in (" & sValidLocIDs & ") or (CreatedByAdminID in (" & sValidSU & ") and Isnull(LocationID,0)=0)) " 
    End If
    MyCommon.LRTsp.Parameters.Add("@sJoin", SqlDbType.VarChar).Value = sJoin
    MyCommon.LRTsp.Parameters.Add("@wherestr", SqlDbType.VarChar).Value = wherestr
    
    dst = MyCommon.LRTsp_select
    
 
    sizeOfData = dst.Rows.Count
    ' set i
    i = linesPerPage * PageNum
  
    If (sizeOfData = 1 AndAlso Server.HtmlEncode(Request.QueryString("searchterms")) <> "") Then
        Response.Status = "301 Moved Permanently"
        Response.AddHeader("Location", "pgroup-edit.aspx?ProductGroupID=" & dst.Rows(i).Item("ProductGroupID"))
    End If
%>
<div id="intro">
    <h1 id="title">
        <% Sendb(Copient.PhraseLib.Lookup("term.productgroups", LanguageID))%>
    </h1>
    <div id="controls">
        <form action="pgroup-edit.aspx" id="controlsform" name="controlsform">
        <%
            If (bCreateStaticPG) Then
              Sendb("<input type=""submit"" accesskey=""n"" class=""regular"" id=""newstatic"" name=""newstatic"" value=""" & Copient.PhraseLib.Lookup("term.new", LanguageID) & " " & Copient.PhraseLib.Lookup("term.static", LanguageID) & """" & "style=""margin-right: 25px"" />")
            End If
            If (Logix.UserRoles.CreateProductGroups) Then
                Send_New()
            End If
            'If MyCommon.Fetch_SystemOption(75) Then
            '  If (Logix.UserRoles.AccessNotes) Then
            '    Send_NotesButton(7, 0, AdminUserID)
            '  End If
            'End If
        %>
        </form>
    </div>
</div>
<div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
    <% Send_Listbar(linesPerPage, sizeOfData, PageNum, Server.HtmlEncode(Request.QueryString("searchterms")), "&SortText=" & SortText & "&SortDirection=" & SortDirection)%>
 <% If (CriteriaMsg <> "") Then
      Send("<div id=""criteriabar""" & IIf(CriteriaError, " style=""background-color:red;""", "") & ">" & CriteriaMsg & "<a href=""javascript:editSearchCriteria();"" class=""white"" style=""padding-left:15px;"">[" & Copient.PhraseLib.Lookup("term.edit", LanguageID) & "]</a><a href=""pgroup-list.aspx"" class=""white"" style=""padding-left:15px;"">[" & Copient.PhraseLib.Lookup("term.clear", LanguageID) & "]</a></div>")
    End If
    %>
    <table class="list" summary="<% Sendb(Copient.PhraseLib.Lookup("term.productgroups", LanguageID)) %>">
        <thead>
            <tr>
                <th align="left" class="th-xid" scope="col">
                    <a href="pgroup-list.aspx?searchterms=<%Sendb(Server.HtmlEncode(Request.QueryString("searchterms")))%>&SortText=ExtGroupID&SortDirection=<% Sendb(SortDirection) %>">
                        <% Sendb(Copient.PhraseLib.Lookup("term.xid", LanguageID))%>
                    </a>
                    <%
                        If SortText = "ExtGroupID" Then
                            If SortDirection = "ASC" Then
                                Sendb("<span class=""sortarrow"">&#9660;</span>")
                            Else
                                Sendb("<span class=""sortarrow"">&#9650;</span>")
                            End If
                        Else
                        End If
                    %>
                </th>
                <th align="left" class="th-id" scope="col">
                    <a href="pgroup-list.aspx?searchterms=<%Sendb(Server.HtmlEncode(Request.QueryString("searchterms")))%>&SortText=ProductGroupID&SortDirection=<% Sendb(SortDirection) %>">
                        <% Sendb(Copient.PhraseLib.Lookup("term.id", LanguageID))%>
                    </a>
                    <%  If SortText = "ProductGroupID" Then
                            If SortDirection = "ASC" Then
                                Sendb("<span class=""sortarrow"">&#9660;</span>")
                            Else
                                Sendb("<span class=""sortarrow"">&#9650;</span>")
                            End If
                        Else
                        End If
                    %>
                </th>
                <%  If (MyCommon.Fetch_UE_SystemOption(170) = "1") Then%>
                <th align="left" class="th-id" scope="col">
                    <a href="pgroup-list.aspx?searchterms=<%Sendb(Server.HtmlEncode(Request.QueryString("searchterms")))%>&SortText=BuyerID&SortDirection=<% Sendb(SortDirection) %>">
                        <% Sendb(Copient.PhraseLib.Lookup("term.buyerid", LanguageID))%>
                    </a>
                    <%  If SortText = "BuyerID" Then
                            If SortDirection = "ASC" Then
                                Sendb("<span class=""sortarrow"">&#9660;</span>")
                            Else
                                Sendb("<span class=""sortarrow"">&#9650;</span>")
                            End If
                        Else
                        End If
                    %>
                </th>
                <%End If%>
                <th align="left" class="th-name" scope="col">
                    <a href="pgroup-list.aspx?searchterms=<%Sendb(Server.HtmlEncode(Request.QueryString("searchterms")))%>&SortText=Name&SortDirection=<% Sendb(SortDirection) %>">
                        <% Sendb(Copient.PhraseLib.Lookup("term.name", LanguageID))%>
                    </a>
                    <%  If SortText = "Name" Then
                            If SortDirection = "ASC" Then
                                Sendb("<span class=""sortarrow"">&#9660;</span>")
                            Else
                                Sendb("<span class=""sortarrow"">&#9650;</span>")
                            End If
                        Else
                        End If
                    %>
                </th>
                <%  If (bStaticEnabled) Then%>
                <th align="left" class="th-add" scope="col">
                    <a href="pgroup-list.aspx?searchterms=<%sendb(Request.QueryString("searchterms")) %>&SortText=IsStatic&SortDirection=<% Sendb(SortDirection) %>">
                        <% Sendb(Copient.PhraseLib.Lookup("term.static", LanguageID))%>
                    </a>
                    <%If SortText = "Static" Then
                        If SortDirection = "ASC" Then
                          Sendb("<span class=""sortarrow"">&#9660;</span>")
                        Else
                          Sendb("<span class=""sortarrow"">&#9650;</span>")
                        End If
                      End If
                    %>
                </th>
                <%End If%>

                <th align="left" class="th-datetime" scope="col">
                    <a href="pgroup-list.aspx?searchterms=<%Sendb(Server.HtmlEncode(Request.QueryString("searchterms")))%>&SortText=CreatedDate&SortDirection=<% Sendb(SortDirection) %>">
                        <% Sendb(Copient.PhraseLib.Lookup("term.created", LanguageID))%>
                    </a>
                    <%  If SortText = "CreatedDate" Then
                            If SortDirection = "ASC" Then
                                Sendb("<span class=""sortarrow"">&#9660;</span>")
                            Else
                                Sendb("<span class=""sortarrow"">&#9650;</span>")
                            End If
                        Else
                        End If
                    %>
                </th>
                <th align="left" class="th-datetime" scope="col">
                    <a href="pgroup-list.aspx?searchterms=<%Sendb(Server.HtmlEncode(Request.QueryString("searchterms")))%>&SortText=LastUpdate&SortDirection=<% Sendb(SortDirection) %>">
                        <% Sendb(Copient.PhraseLib.Lookup("term.edited", LanguageID))%>
                    </a>
                    <%  If SortText = "LastUpdate" Then
                            If SortDirection = "ASC" Then
                                Sendb("<span class=""sortarrow"">&#9660;</span>")
                            Else
                                Sendb("<span class=""sortarrow"">&#9650;</span>")
                            End If
                        Else
                        End If
                    %>
                </th>
            </tr>
        </thead>
        <tbody>
            <%
                While (i < sizeOfData And i < linesPerPage + linesPerPage * PageNum)
                    IsSpecialGroup = (MyCommon.NZ(dst.Rows(i).Item("PointsNotApplyGroup"), False) OrElse MyCommon.NZ(dst.Rows(i).Item("NonDiscountableGroup"), False))
                    Send("<tr class=""" & Shaded & """>")
            
                    Send("  <td>" & IIf(MyCommon.NZ(dst.Rows(i).Item("ExtGroupID"), "") = "0", "", MyCommon.NZ(dst.Rows(i).Item("ExtGroupID"), "")) & "</td>")
                    Send("  <td>" & dst.Rows(i).Item("ProductGroupID") & "</td>")
                    If (MyCommon.Fetch_UE_SystemOption(170) = "1") Then
                        'Send("  <td>" & MyCommon.NZ(dst.Rows(i).Item("BuyerID"), "") & "</td>")
                        Dim externalbuyerid As String = MyCommon.GetExternalBuyerId(MyCommon.NZ(dst.Rows(i).Item("BuyerID"), ""))
                        Send("  <td>" & externalbuyerid & "</td>")
                    End If
                    If (dst.Rows(i).Item("ProductGroupID") > 1) Then
                        Sendb("  <td><div style=""width:330px;word-wrap:break-word;"">")
                        If (IsSpecialGroup) Then
                            Sendb("<img src=""../images/not.png"" />&nbsp;")
                        End If
                        If (MyCommon.IsEngineInstalled(9) And MyCommon.Fetch_UE_SystemOption(168) = "1" And Not IsDBNull(dst.Rows(i).Item("Buyerid"))) Then
                            Dim buyerid As Integer = dst.Rows(i).Item("Buyerid")
                            Dim externalBuyerid = MyCommon.GetExternalBuyerId(buyerid)
                            ' Send("<a href=""pgroup-edit.aspx?ProductGroupID=" & dst.Rows(i).Item("ProductGroupID") & """>" & MyCommon.SplitNonSpacedString("Buyer " & externalBuyerid & " - " & dst.Rows(i).Item("Name"), 30) & "</a></div></td>")
                            Send("<a href=""pgroup-edit.aspx?ProductGroupID=" & dst.Rows(i).Item("ProductGroupID") & """>Buyer " & externalBuyerid & " - " & dst.Rows(i).Item("Name")& "</a></div></td>")
                        Else
                            'Send("<a href=""pgroup-edit.aspx?ProductGroupID=" & dst.Rows(i).Item("ProductGroupID") & """>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Name"), 30) & "</a></div></td>")
                            Send("<a href=""pgroup-edit.aspx?ProductGroupID=" & dst.Rows(i).Item("ProductGroupID") & """>" & dst.Rows(i).Item("Name") & "</a></div></td>")
                        End If
                        
                    Else
                        'Send("  <td>" & MyCommon.SplitNonSpacedString(dst.Rows(i).Item("Name"), 30) & "</td>")
                        Send("  <td><div style=""width:330px;word-wrap:break-word;"">" & dst.Rows(i).Item("Name") & "</div></td>")
                    End If

                    If (bStaticEnabled) Then
                      If dst.Rows(i).Item("IsStatic") = "1" Then
                        Send("  <td>Yes</td>")
                      Else
                        Send("  <td></td>")
                      End If
                    End If
                        If (Not IsDBNull(dst.Rows(i).Item("CreatedDate"))) Then
                            Send("  <td>" & Logix.ToShortDateTimeString(dst.Rows(i).Item("CreatedDate"), MyCommon) & "</td>")
                        Else
                            Send("  <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
                        End If
                        If (Not IsDBNull(dst.Rows(i).Item("LastUpdate"))) Then
                            Send("  <td>" & Logix.ToShortDateTimeString(dst.Rows(i).Item("LastUpdate"), MyCommon) & "</td>")
                        Else
                            Send("  <td>" & Copient.PhraseLib.Lookup("term.unknown", LanguageID) & "</td>")
                        End If
                        Send("</tr>")
                        If Shaded = "shaded" Then
                            Shaded = ""
                        Else
                            Shaded = "shaded"
                        End If
                        i = i + 1
                End While
            %>
        </tbody>
    </table>
</div>
<script runat="server">
  
private function hasOption( byref optionName as string ) as Boolean 
    
    dim val as String = Request.Form( optionName )
    return val isnot nothing AndAlso val.Trim().Length > 0
        
end function

Function GetOptionString(ByRef MyCommon As Copient.CommonInc, ByVal OptionIndex As Integer, _
                           ByVal OptionValue As String, ByVal FieldName As String) As String
    Dim FieldBuf As New StringBuilder()
    FieldBuf.Append(FieldName & " ")
    Select Case OptionIndex
      Case 1 ' contains
        FieldBuf.Append(" like '%" & MyCommon.Parse_Quotes(OptionValue.Trim) & "%' ")
      Case 2 ' exact
        FieldBuf.Append(" = '" & MyCommon.Parse_Quotes(OptionValue.Trim) & "' ")
      Case 3 ' starts with
        FieldBuf.Append(" like '" & MyCommon.Parse_Quotes(OptionValue.Trim) & "%' ")
      Case 4 ' ends with
        FieldBuf.Append(" like '%" & MyCommon.Parse_Quotes(OptionValue.Trim) & "' ")
      Case 5 ' excludes
        FieldBuf = New StringBuilder()
        FieldBuf.Append(" (" & FieldName & " not like '%" & MyCommon.Parse_Quotes(OptionValue.Trim) & "%' or " & FieldName & " is null) ")
      Case 6 ' is
        FieldBuf.Append(" = " & MyCommon.Parse_Quotes(OptionValue.Trim) & " ")
      Case 7 ' is not
        FieldBuf.Append(" <> " & MyCommon.Parse_Quotes(OptionValue.Trim) & " ")
      Case Else ' default to contains
        FieldBuf.Append(" like '%" & MyCommon.Parse_Quotes(OptionValue.Trim) & "%' ")
    End Select
    Return FieldBuf.ToString
  End Function
  
  Function GetOptionType(ByVal OptionIndex As Integer) As String
    Dim OptionType As String = "contains"
    Select Case OptionIndex
      Case 1 ' contains
        OptionType = Copient.PhraseLib.Lookup("term.contains", LanguageID)
      Case 2 ' exact
        OptionType = "="
      Case 3 ' starts with
        OptionType = Copient.PhraseLib.Lookup("term.startswith", LanguageID)
      Case 4 ' ends with
        OptionType = Copient.PhraseLib.Lookup("term.endswith", LanguageID)
      Case 5 ' excludes
        OptionType = Copient.PhraseLib.Lookup("term.excludes", LanguageID)
      Case 6 ' is
        OptionType = Copient.PhraseLib.Lookup("term.is", LanguageID)
      Case 7 ' is not
        OptionType = Copient.PhraseLib.Lookup("term.IsNot", LanguageID)
      Case Else ' default to contains
        OptionType = Copient.PhraseLib.Lookup("term.contains", LanguageID)
    End Select
    Return OptionType
  End Function
  
</script>

<form id="frmIter" name="frmIter" method="post" action="">
  <input type="hidden" id="advSql" name="advSql" value="<% Sendb(Server.UrlEncode(AdvSearchSQL)) %>" />
  <input type="hidden" id="advCrit" name="advCrit" value="<% Sendb(Server.UrlEncode(CriteriaMsg)) %>" />
  <input type="hidden" id="advTokens" name="advTokens" value="<%Sendb(Server.UrlEncode(CriteriaTokens)) %>" />
</form>
<!-- overwrite the iteration links and post the form -->
<%
    'If MyCommon.Fetch_SystemOption(75) Then
    '  If (Logix.UserRoles.AccessNotes) Then
    '    Send_Notes(7, 0, AdminUserID)
    '  End If
    'End If
done:
    Send_BodyEnd("searchform", "searchterms")
    MyCommon.Close_LogixRT()
    Logix = Nothing
    MyCommon = Nothing
%>
