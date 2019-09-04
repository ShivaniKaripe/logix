?<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: Customer-transaction-items.aspx 
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
  Dim rst As DataTable
  Dim rst2 As DataTable
  Dim rst3 As DataTable
  Dim dt As DataTable
  Dim row As DataRow
  Dim Shaded As String = " class=""shaded"""
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
   
  Dim TransNum As String
  Dim SortURL As String = ""
  Dim SortCol As String = "Items"
  Dim SortDir As String = "desc"
  Dim ContextStyle As String = "color:#000000;"
  
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If
  
  Response.Expires = 0
  MyCommon.AppName = "customer-transaction-items.aspx"
  MyCommon.Open_LogixRT()
  MyCommon.Open_LogixTRX()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  TransNum = Request.QueryString("TransNum")
  
  SortUrl = "customer-transaction-items.aspx?TransNum=" & TransNum 

  Send_HeadBegin("term.customerid", "term.transactionnumber", TransNum)
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
  Send_Scripts()
%>
<script type="text/javascript">

</script>
<%
  Send_HeadEnd()
 
  Send_BodyBegin(2)
%>
<form action="#" id="mainform" name="mainform" >
  <div id="intro">
    <h1 id="title">
    <%    
      Sendb(Copient.PhraseLib.Lookup("term.transaction", LanguageID) & " #" & TransNum) 
    %>
    </h1>

  </div>
  <div id="main">
    <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
   
   <input type="hidden" id="TransNum" name="TransNum" value="<% sendb(TransNum) %>" />
    
    <div id="column">
      <div class="box" id="Items">
        <h2>
          <span>
            <% Sendb(Copient.PhraseLib.Lookup("term.items", LanguageID))%>
          </span>
        </h2>
        <table summary=""" & Copient.PhraseLib.Lookup("term.transactionhistory", LanguageID) & """>
          <thead>
            <tr>
              <th align="left" class="th-Item" scope="col">
                <a href="<% Sendb(SortUrl & "&amp;sortcol=Item&amp;sortdir=" & IIf(SortDir = "desc", "asc", "desc")) %>">
                  <% Sendb(Copient.PhraseLib.Lookup("term.Item", LanguageID))%></a>
                <%
                  If SortCol = "Item" Then
                    If SortDir = "asc" Then
                      Sendb("<span class=""sortarrow"">&#9660;</span>")
                    Else
                      Sendb("<span class=""sortarrow"">&#9650;</span>")
                    End If
                  End If
                %>
              </th>
              <th align="left" class="th-Quantity" scope="col">
                <a href="<% Sendb(SortUrl & "&amp;sortcol=Quantity&amp;sortdir=" & IIf(SortDir = "desc", "asc", "desc")) %>">
                  <% Sendb(Copient.PhraseLib.Lookup("term.Quantity", LanguageID))%></a>
                <%
                  If SortCol = "Quantity" Then
                    If SortDir = "asc" Then
                      Sendb("<span class=""sortarrow"">&#9660;</span>")
                    Else
                      Sendb("<span class=""sortarrow"">&#9650;</span>")
                    End If
                  End If
                %>
              </th>
              <th align="left" class="th-Price" scope="col">
                <a href="<% Sendb(SortUrl & "&amp;sortcol=Price&amp;sortdir=" & IIf(SortDir = "desc", "asc", "desc")) %>">
                  <% Sendb(Copient.PhraseLib.Lookup("term.Price", LanguageID))%></a>
                <%
                  If SortCol = "Price" Then
                    If SortDir = "asc" Then
                      Sendb("<span class=""sortarrow"">&#9660;</span>")
                    Else
                      Sendb("<span class=""sortarrow"">&#9650;</span>")
                    End If
                  End If
                %>
              </th>
              <th align="left" class="th-Description" scope="col">
                <a href="<% Sendb(SortUrl & "&amp;sortcol=Description&amp;sortdir=" & IIf(SortDir = "desc", "asc", "desc")) %>">
                  <% Sendb(Copient.PhraseLib.Lookup("term.Description", LanguageID))%></a>
                <%
                  If SortCol = "Description" Then
                    If SortDir = "asc" Then
                      Sendb("<span class=""sortarrow"">&#9660;</span>")
                    Else
                      Sendb("<span class=""sortarrow"">&#9650;</span>")
                    End If
                  End If
                %>
              </th>
            </tr>
          </thead>
          <br/>
          <tbody>
            
            <%
            Try
              MyCommon.QueryStr = "Select ItemID, Quantity, Price, Description from TransactionItem with (NoLock) where LogixTransNum='" & TransNum & "';"
              rst = MyCommon.LTRX_Select
              If rst.Rows.Count > 0 Then
                For Each row In rst.Rows
                  Dim ItemID As String = MyCommon.NZ(row.Item("ItemID"), "")
                  Dim Quantity As Decimal= MyCommon.NZ(row.Item("Quantity"), 0)
                  Dim Price As Decimal = MyCommon.NZ(row.Item("Price"), 0)
                  Dim Description As String = MyCommon.NZ(row.Item("Description"), "")

            %>
                <tr>
                  <td style=""" & ContextStyle & """>
                  <%Send(ItemID)%>
                  </td>
                  <td style=""" & ContextStyle & """>
                  <%Send(Quantity)%>
                  </td>
                  <td style=""" & ContextStyle & """>
                  <%Send(Price)%>
                  </td>
                  <td style=""" & ContextStyle & """>
                  <%Send(Description)%>
                  </td>
                </tr>  
            <%

                Next
              Else
                Send("<tr>")
                Send("  <td colspan=""7"" style=""text-align:center""><i>" & Copient.PhraseLib.Lookup("customer-inquiry.nohistory", LanguageID) & "</i></td>")
                Send("</tr>")
              End If
              Catch ex As Exception
              infomessage = ex.ToString
            End Try
            %>
              
            </tr>
          </tbody>
        </table>
      </div>
    </div>

  </div>
</form>

<script runat="server">
  
</script>
<script type="text/javascript">

</script>
<%
done:
  MyCommon.Close_LogixRT()
  MyCommon.Close_LogixTRX()
  Send_BodyEnd("mainform")
  MyCommon = Nothing
  Logix = Nothing
%>
