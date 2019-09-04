<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" ValidateRequest="false" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>

<%@ Import Namespace="System.Data" %>
<%
  ' *****************************************************************************
  ' * FILENAME: cgroup-edit-generatebarcodes.aspx 
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
  Dim row As DataRow
  Dim rst2 As DataTable
  Dim row2 As DataRow
  Dim TierLevel As String
  Dim Submitted As Boolean = False
  Dim infoMessage As String = ""
  Dim Handheld As Boolean = False
  Dim CustomerGroupID As Long
  Dim Parent as String
  Dim DisableNumberOfBarcodes as Boolean 
  If Request IsNot Nothing AndAlso Request.Browser IsNot Nothing AndAlso Request.Browser.Platform IsNot Nothing AndAlso Request.ServerVariables IsNot Nothing AndAlso Request.ServerVariables("HTTP_USER_AGENT") IsNot Nothing Then
    Handheld = DetectHandheld(Request.Browser("IsMobileDevice"), Request.Browser.Platform.ToString(), Request.ServerVariables("HTTP_USER_AGENT").ToString())
  End If

  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  Response.Expires = 0
  MyCommon.AppName = "cgroup-edit-generatebarcodes.aspx"
  CustomerGroupID = Request.QueryString("CustomerGroupID")
  Parent = Request.QueryString("Parent")
   DisableNumberOfBarcodes= MyCommon.NZ( Request.QueryString("DisableNumberOfBarcodes"),false)
  Send_HeadBegin("term.preview")
  Send_Comments(CopientProject, CopientFileName, CopientFileVersion, CopientNotes)
  Send_Metas()
  Send_Links(Handheld)
   
  ' Calculate how many pixels wide the pmsgpreviewbody div should be, based on the page width
  Dim FixedWidthFont As Boolean = true
  Dim PageWidthPixels As Integer
  Dim PageWidth As Integer = 20
  If (PageWidth = 0) Or (PageWidth > 50) Then
    PageWidthPixels = 400
  Else
    PageWidthPixels = PageWidth * 8
  End If

 ' Generate custom CSS to style the DIV
  Send("<style type=""text/css"">")
  Send("#generatebarcodes {")
  Send("}")
  Send("* html #generatebarcodes {")
  Send("  white-space: nowrap;")
  Send("}")
  Send("#generatebarcodesbody {")
  Send("  border: 0;")
  Send("  padding: 0;")
  Send("  font-size: 13px;")
  If (FixedWidthFont = False) Then
    Send("  font-family: Verdana, Arial;")
  Else
    Send("  font-family: monospace;")
  End If
  Send("  overflow-x: hidden;")
  
  Send("}")
  Send("</style>")
    
  Send_Scripts()
  Send_HeadEnd()
  '3 - straight popup
  '2 - ChangeParentDocument onunload event
  Send_BodyBegin(2)
  
  If (Logix.UserRoles.AccessOffers = False) Then
    Send_Denied(2, "perm.offers-access")
    GoTo done
  End If
%>

<script type="text/javascript">
  window.name="GenerateBarcodes"

   function ChangeParentDocument()
   {
   //add validiation flag to query string and things to pass back
   //if necessary, pull query string from elements in the document
   //all input elements, unless disabled, are put into the query string
     if (opener != null && !opener.closed) {
       //opener.location = '/logix/cgroup-edit.aspx?CustomerGroupID=<%Sendb(CustomerGroupID)%>&LocationID=<%Sendb(Request.QueryString("locationid"))%>&RedemptionRestrictionID=<%Sendb(Request.QueryString("locationtype"))%>&UPC=<%Sendb(Request.QueryString("upc"))%>&SVProgramID=<%Sendb(Request.QueryString("storedvalueprograms-available"))%>&ValidateBarcode=True';
       var UPC = document.getElementById("upc").value;
       var SVProgramID = document.getElementById("storedvalueprograms-available").value;
       var LocationID = document.getElementById("locationid").value;
       var LocationGroupID = document.getElementById("locationgroups-available").value;
       var RedemptionRestrictionID = document.getElementById("locationtype").value;
       var Validate = document.getElementById("validatebarcodes").value;
       var NumOfBarcodes = document.getElementById("NumberOfBarcodes").value;
	   var offerID = document.getElementById("offerid").value;
       opener.location = '/logix/<%Sendb(Request.QueryString("Parent"))%>?CustomerGroupID=<%Sendb(Request.QueryString("CustomerGroupID"))%>&LocationID=' + LocationID + '&RedemptionRestrictionID=' + RedemptionRestrictionID + '&LocationGroupID=' + LocationGroupID + '&UPC=' + UPC + '&SVProgramID=' + SVProgramID + '&ValidateBarcode=' + Validate + '&NumOfBarcodes=' + NumOfBarcodes + '&OfferID=' + offerID; 
       window.opener.focus();
       window.close();
     }
   }
   function enableSubmit()
   {
      document.getElementById("submitinfo").style.visibility = 'visible';
   }
  function SetToValidate()
   {
      document.getElementById("validatebarcodes").value = 'True';
   }
   
   function checkGroup()
   {
      var type = document.getElementById("locationtype").value;
      if (type == 0)
      {
         document.getElementById("locationidlabel").style.visibility = 'hidden';
         document.getElementById("locationid").style.visibility = 'hidden';
         document.getElementById("locationgrouplist").style.visibility = 'hidden';
         document.getElementById("locationgroups-available").style.visibility = 'hidden';
      }
      if (type == 1)
      {
         document.getElementById("locationidlabel").style.visibility = 'visible';
         document.getElementById("locationid").style.visibility = 'visible';
         document.getElementById("locationgrouplist").style.visibility = 'hidden';
         document.getElementById("locationgroups-available").style.visibility = 'hidden';
      }
      if (type == 2)
      {
         document.getElementById("locationidlabel").style.visibility = 'hidden';
         document.getElementById("locationid").style.visibility = 'hidden';
         document.getElementById("locationgrouplist").style.visibility = 'visible';
         document.getElementById("locationgroups-available").style.visibility = 'visible';
      }
   }
   
</script>

<form action="cgroup-edit.aspx" id="mainform" name="mainform" onsubmit="SetToValidate()" >
  <div id="intro">
  <h1 id="title">
    <% Send(Copient.PhraseLib.Lookup("term.GenerateBarcodes", LanguageID))%>
   </h1>
</div>
<a name="h00" id="h00"></a>
<div id="main">
  <% Send("<input type=""hidden"" id=""validatebarcodes"" name=""validatebarcodes"" />") %>
  <% If (infoMessage <> "") Then Send("<div id=""infobar"" class=""red-background"">" & infoMessage & "</div>")%>
  <%
    TierLevel = MyCommon.NZ(Request.QueryString("TierLevel"), "")
    If TierLevel.Length > 0 Then
      Sendb(TierLevel)
    End If
  %>
  <div id="column2x">
    <div id="generatebarcodes">
      <div id="generatebarcodesbody">
           <div class="box" id="storedvalueprograms">
          <h2><span><b><%Sendb(Copient.PhraseLib.Lookup("term.input", LanguageID))%></b></span></h2>
          <label for="storedvalueprograms-available"><b><% Sendb(Copient.PhraseLib.Lookup("generatebarcodes.AvailableSV", LanguageID) & ":")%></b></label>
           <br clear="all" />
        <br class="half" />
        <%
                Send("<span id=""storedvalueprogramlist"">")
                Send("<select class=""longest"" multiple=""multiple"" id=""storedvalueprograms-available"" name=""storedvalueprograms-available"" size=""10"" onchange=""enableSubmit();"" >")
                MyCommon.QueryStr = "select SVProgramID, Name from StoredValuePrograms where Deleted = 'False'"
                rst = MyCommon.LRT_Select
                If (rst.Rows.Count > 0) Then
                  For Each row In rst.Rows
                      Send("<option value=""" & row.Item("SVProgramID") & """>" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
                  Next
                End If
                Send("</select>")
                Send("</span>")
                Send("<br />")
            %>
			<table>
			<tr>
			<td>
          <label for="upc"><b><%Sendb(Copient.PhraseLib.Lookup("term.upc", LanguageID))%>:</b></label><br />
		   <%
               Sendb("<input type=""text"" class=""regular"" id=""upc"" name=""upc"" maxlength=""10"" value="""" /><br />")
			   %>
			   </td>
		  <td>
		            <label for="upc"><b><%Sendb(Copient.PhraseLib.Lookup("term.numofbarcodes", LanguageID))%>:</b></label><br />
		   <%
               Sendb("<input type=""text"" class=""regular"" id=""NumberOfBarcodes"" name=""NumberOfBarcodes"" maxlength=""10"" value=""1"" "& IIf(DisableNumberOfBarcodes = true, "disabled=""disabled""", "") &" /><br />")
		    %>
		  </td>
		  </tr>

			<tr>
			<td>
          <label for="locationtype"><b><%Sendb(Copient.PhraseLib.Lookup("generatebarcodes.LocationsValid", LanguageID))%>:</b></label><br />

            <%
               Send("<select id=""locationtype"" name=""locationtype"" onchange=""checkGroup();"">")
            Send("<option value=""0"" selected=""selected"">" & Copient.PhraseLib.Lookup("term.alllocations", LanguageID) & "</option>")
            Send("<option value=""1"">" & Copient.PhraseLib.Lookup("term.IndividualLocation", LanguageID) & "</option>")
            Send("<option value=""2"">" & Copient.PhraseLib.Lookup("term.locationgroup", LanguageID) & "</option>")
               Send("</select>")
               Send("</br>")
            %>
			</td>
			<td>
			<label for="offerid"><b><%Sendb(Copient.PhraseLib.Lookup("term.offerid", LanguageID))%>:</b></label><br />
			<input type="text" class="regular" id="offerid" name="offerid" value="" />
			</td>
			</tr>
			
					  </table>
          <br />
          <label id="locationidlabel" style="visibility: hidden;"><b><%Sendb("External Location Code")%>:</b></label><br />
            <%
               Sendb("<input type=""text"" class=""regular"" id=""locationid"" name=""locationid"" value="""" style=""visibility: hidden;""/><br />")
            %>
            <%
                Send("<span id=""locationgrouplist"" style=""visibility: hidden;"">")
                Send("<select class=""longest"" multiple=""multiple"" id=""locationgroups-available"" name=""locationgroups-available"" size=""10"" style=""visibility: hidden;"" >")
                MyCommon.QueryStr = "select LocationGroupID, Name from LocationGroups where Deleted = 'False'"
                rst2 = MyCommon.LRT_Select
                If (rst2.Rows.Count > 0) Then
                  For Each row In rst2.Rows
                      Send("<option value=""" & row.Item("LocationGroupID") & """>" & MyCommon.NZ(row.Item("Name"), "") & "</option>")
                  Next
                End If
                Send("</select>")
                Sendb("</span>")
                Sendb("<br />")
                Sendb("<br />")
            %>
            <%
                  Sendb("<input type=""submit"" class=""regular"" id=""submitinfo"" name=""submitinfo"" value=""" & Copient.PhraseLib.Lookup("term.submit", LanguageID) & """ style=""visibility: hidden;""/>    ")
                  Sendb("<input type=""button"" class=""regular"" id=""cancel"" name=""cancel"" value=""" & Copient.PhraseLib.Lookup("term.cancel", LanguageID) & """ onClick=""self.close()""/>    ")
                  Send("<br />")
                  Send("<br />")
            %>
        </div>
      </div>
    </div>
  </div>
  <div id="gutter">
  </div>
</div>
</form>

<%
done:
  MyCommon.Close_LogixRT()
  Send_BodyEnd("mainform")
  Logix = Nothing
  MyCommon = Nothing
%>