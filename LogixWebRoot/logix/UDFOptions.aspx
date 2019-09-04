<%@ Page Language="vb" Debug="true" CodeFile="LogixCB.vb" Inherits="LogixCB" %>
<%' version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%@ Import Namespace="System.Data" %>

<%
   
  Dim CopientFileName As String = "UDFOptions.aspx"
  Dim CopientFileVersion As String = "7.3.1.138972" 
  Dim CopientProject As String = "Copient Logix"
  Dim CopientNotes As String = ""
  Dim AdminUserID As Integer
  Dim MyCommon As New Copient.CommonInc
  Dim Logix As New Copient.LogixInc
  MyCommon.AppName = "UDFOptions.aspx"
  MyCommon.Open_LogixRT()
  AdminUserID = Verify_AdminUser(MyCommon, Logix)
  
  Response.Expires = 0
  Response.Clear()
  Response.ContentType = "text/html"
    Select Case Request.QueryString("Mode")
		Case "AdvSearchOption"
			getUDFAdvSearchOption(MyCommon.Extract_Val(Request.QueryString("udf")), MyCommon.Extract_Val(Request.QueryString("optionSelect")), MyCommon.Extract_Val(Request.QueryString("row")))
		Case "AdvSearch"
			getSelectedUDFAdvSearch(MyCommon.Extract_Val(Request.QueryString("udf")), MyCommon.Extract_Val(Request.QueryString("row")),Request.QueryString("value1"),Request.QueryString("value2"))
        Case "OfferAdd"
            getOfferUDF(MyCommon.Extract_Val(Request.QueryString("udf")), MyCommon.Extract_Val(Request.QueryString("OfferID")))
		Case "OfferDel"
            deleteUDFfromOffer(MyCommon.Extract_Val(Request.QueryString("OfferID")), MyCommon.Extract_Val(Request.QueryString("udf")))        
    End Select
    Response.Flush()
  Response.End()
  %>
  <script runat="server">
  Public MyCommon As New Copient.CommonInc
  Public Logix As New Copient.LogixInc
  
  Sub getUDFAdvSearchOption(ByVal UDFOption As Integer, ByVal optionSelect As Integer, ByVal rowNum As Integer)
    Dim dst As DataTable
    Dim row As DataRow
	
    Try
      If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
	
      MyCommon.QueryStr = "select * from UserDefinedFields where AdvancedSearch = 1"
      dst = MyCommon.LRT_Select
      If Not IsDBNull(dst.Rows(UDFOption)) Then
        row = dst.Rows(UDFOption)

        Select Case row.Item("DataType")
            Case 0, 1, 4, 5, 6, 7
            Send("<select id=""udfOption-" & rowNum & """ name=""udfOption-" & rowNum & """ class=""mediumshort"">")
            Send("<option value=""1""" & IIf(optionSelect = 1, " selected=""selected""", "") & ">")
            Sendb(Copient.PhraseLib.Lookup("term.contains", LanguageID))
            Send("</option>")
            Send("<option value=""2""" & IIf(optionSelect = 2, " selected=""selected""", "") & ">")
            Sendb(Copient.PhraseLib.Lookup("term.exact", LanguageID))
            Send("</option>")
            Send("<option value=""3""" & IIf(optionSelect = 3, " selected=""selected""", "") & ">")
            Sendb(Copient.PhraseLib.Lookup("term.startswith", LanguageID))
            Send("</option>")
            Send("<option value=""4""" & IIf(optionSelect = 4, " selected=""selected""", "") & ">")
            Sendb(Copient.PhraseLib.Lookup("term.endswith", LanguageID))
            Send("</option>")
            Send("<option value=""5""" & IIf(optionSelect = 5, " selected=""selected""", "") & ">")
            Sendb(Copient.PhraseLib.Lookup("term.excludes", LanguageID))
            Send("</option>")
            Send("</select>")
          Case 2
            Send("<select id=""udfOption-" & rowNum & """ name=""udfOption-" & rowNum & """ class=""mediumshort"" onchange=""handleDateToFrom(this.selectedIndex, 'trUdfEnd-" & rowNum & "', 'udfEnd-" & rowNum & "');"">")
            Send("<option value=""0""" & IIf(optionSelect = 0, " selected=""selected""", "") & ">")
            Sendb(Copient.PhraseLib.Lookup("term.on", LanguageID))
            Send("</option>")
            Send("<option value=""1""" & IIf(optionSelect = 1, " selected=""selected""", "") & ">")
            Sendb(Copient.PhraseLib.Lookup("term.before", LanguageID))
            Send("</option>")
            Send("<option value=""2""" & IIf(optionSelect = 2, " selected=""selected""", "") & ">")
            Sendb(Copient.PhraseLib.Lookup("term.after", LanguageID))
            Send("</option>")
            Send("<option value=""3""" & IIf(optionSelect = 3, " selected=""selected""", "") & ">")
            Sendb(Copient.PhraseLib.Lookup("term.between", LanguageID))
            Send("</option>")
            Send("</select>")
          Case 3
            Send("<select id=""udfOption-" & rowNum & """ name=""udfOption-" & rowNum & """ class=""mediumshort"">")
            Send("<option value=""0""" & IIf(optionSelect = 0, " selected=""selected""", "") & ">")
            Send("&nbsp; </option>")
            Send("<option value=""6""" & IIf(optionSelect = 6, " selected=""selected""", "") & ">")
            Send(Copient.PhraseLib.Lookup("term.yes", LanguageID))
            Send("</option>")
            Send("<option value=""7""" & IIf(optionSelect = 7, " selected=""selected""", "") & ">")
            Sendb(Copient.PhraseLib.Lookup("term.no", LanguageID))
            Send("</option>")
            Send("</select>")
        End Select
      End If
    Catch ex As Exception
      Send(ex.ToString)
    Finally
      MyCommon.Close_LogixRT()
      MyCommon = Nothing
    End Try
  End Sub
  
  Sub getSelectedUDFAdvSearch(ByVal UDFOption As Integer, ByVal rowNum As Integer, ByVal value1 As String, ByVal value2 As String)
    Dim dst As DataTable
    Dim row As DataRow
		
    Try
      If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
		
      MyCommon.QueryStr = "select * from UserDefinedFields where AdvancedSearch = 1"
      dst = MyCommon.LRT_Select
      If Not IsDBNull(dst.Rows(UDFOption)) Then
        row = dst.Rows(UDFOption)
        Select Case row.Item("DataType")
            Case 0, 1, 4, 5, 6, 7
            Send("<input id=""udf-" & rowNum & """ name=""udf-" & rowNum & """ type=""text"" value=""" & value1 & """ class=""medium"" />") 'Sendb(udfs("udf-" & rowNum)) 
          Case 2
            Send("<input id=""udf-" & rowNum & """ name=""udf-" & rowNum & """ type=""text"" value=""" & value1 & """  class=""mediumshort"" />") 'value="<%Sendb(udfs("udf-" & rowNum)) %>
            Send("<img src=""../images/calendar.png"" class=""calendar"" id=""enddate2picker"" alt=""" & Copient.PhraseLib.Lookup("term.datepicker", LanguageID) & """" & _
             " title=""" & Copient.PhraseLib.Lookup("term.datepicker", LanguageID) & """ onclick=""displayDatePicker('udf-" & rowNum & "',event);"" />")
            Send("<tr>") 'IIF(udfs("udf-" & rowNum & "End") ="","style =""display: none;""","")) %>>
            Send("<input id=""udfEnd-" & rowNum & """ name=""udfEnd-" & rowNum & """ type=""text"" value=""" & value2 & """ class=""mediumshort"" />") 'value="<%Sendb(udfs("udf-" & rowNum & "End")) %>
            Send("<img src=""../images/calendar.png"" class=""calendar"" id=""enddate2picker"" alt=""" & Copient.PhraseLib.Lookup("term.datepicker", LanguageID) & """" & _
              " title=" & Copient.PhraseLib.Lookup("term.datepicker", LanguageID) & """ onclick=""displayDatePicker('udfEnd-" & rowNum & "',event);"" />")
        End Select
      End If
    Catch ex As Exception
      Send(ex.ToString)
    Finally
      MyCommon.Close_LogixRT()
      MyCommon = Nothing
    End Try
  End Sub
      
      
  Sub getOfferUDF(ByVal UDFoption As Integer, ByVal OfferID As Integer)
    Dim dst As DataTable
    Dim udfValuesdst As DataTable
    Dim UDFPK As Int64
    Dim Status As Integer
		
    Dim AllowEditing As Boolean
		
    Try
      If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
      AdminUserID = Verify_AdminUser(MyCommon, Logix)
      AllowEditing = Logix.UserRoles.EditUserDefinedFields

      MyCommon.QueryStr = "select udf.UDFPK, udf.Description from UserDefinedFields as udf " & _
                "where not exists (select UDFPK from UserDefinedFieldsValues as v where deleted = 0 and udf.UDFPK = v.UDFPK and v.OfferID = " & OfferID & ")"
      dst = MyCommon.LRT_Select
      If dst.Rows.Count > 0 Then
        UDFPK = dst.Rows(UDFoption).Item("UDFPK")
      End If

      MyCommon.QueryStr = "dbo.pt_UserDefinedFieldsValues_Insert"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@UDFPK", SqlDbType.BigInt).Value = UDFPK
      MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.Int).Value = OfferID
      MyCommon.LRTsp.Parameters.Add("@Status", SqlDbType.Int).Direction = ParameterDirection.Output
      MyCommon.LRTsp.ExecuteNonQuery()
      Status = MyCommon.LRTsp.Parameters("@Status").Value
      MyCommon.Close_LRTsp()
              
      MyCommon.Activity_Log(3, OfferID, AdminUserID, Copient.PhraseLib.Lookup("history.offer-adduserdefinedfield", LanguageID) & " " & MyCommon.TruncateString(dst.Rows(UDFoption).Item("Description"), 20))
              
              
      If Status = 1 Then
        MyCommon.QueryStr = "Select  UDFPK, Description, DataType,coalesce(UDFPS_ID,-1) as UDFPS_ID from UserDefinedFields as udf with (NoLock) inner join UserDefinedFieldsTypes as type on udf.DataType = type.UDFTypeID where udf.UDFPK = " & UDFPK
        dst = MyCommon.LRT_Select
        Send("<tr id = ""TRudfVal-" & UDFPK & """>")
				
        If Logix.UserRoles.DeleteUserDefinedFields Then
          Send("<td>")
          Send("    <input type=""button"" value=""X"" title=""" & Copient.PhraseLib.Lookup("term.delete", LanguageID) & """ name=""ex"" id=""ex-" & MyCommon.NZ(UDFPK, 0) & """ class=""ex""" & " onclick=""javascript:deleteUDF(" & MyCommon.NZ(UDFPK, 0) & ", " & OfferID & ")"" />")
          Send("</td>")
        End If
				
        MyCommon.QueryStr = "Select UDFVPK,UDFPK,Value,IsDefault,DisplayOrder from UserDefinedField_ValidValues where UDFPK=" & UDFPK & " order by DisplayOrder"
        udfValuesdst = MyCommon.LRT_Select
                  
        '  Send("    <td>" & MyCommon.NZ(row.Item("ExtID"), "") & "</td>")
        Send("    <td><span title=""" & Convert.ToString(MyCommon.NZ(dst.Rows(0).Item("Description"), "")).Replace(System.Convert.ToChar(34).ToString(), "&quot;") & """style=""display: inline-block; width: 128px;min-width: 20px; max-width: 128px;   overflow: hidden;   text-overflow: ellipsis;"">" & MyCommon.TruncateString(MyCommon.NZ(dst.Rows(0).Item("Description"), ""), 27) & "</span></td>")
        Send("    <td>")

        ''DataType values
        'UDFTypeID	Type	        ExtTypeID	    ColumnName
        '
        '0	        String	        String	        StringValue
        '1	        Integer	        Int32	        IntValue
        '2	        Date	        Date	        DateValue
        '3	        Boolean	        Boolean	        BooleanValue
        '7	        List Box	    ListBox	        ListBox
        '8	        Numeric Range	NumericRange	NumericRange
        '9	        Likert	        Likert	        Likert
                  
                  
        'Presentation Styles
        'UDFPS_ID	PresentationStyle	        PresentationStyleID
                  
        '1	        Drop-Down List	            DropDownList
        '2	        Radio Buttons - Horizontal	HorizontalRadioButtons
        '3	        Radio Buttons - Vertical	VerticalRadioButtons
        '4	        List Box	                ListBox
        '5	        Check Box	                CheckBox
        '6	        Text Box	                TextBox
                  
                  
        'valid presentation styles by Data type
        'String
        '     Text box 
        '     List box 
        '     Drop down 
        'Integer
        '     Text box 
        '     List box 
        '     Drop down
        'Date
        '     Text box
        'Boolean
        '     Horizontal radio
        '     Vertical radio
        '     Hoizontal CheckBox
        '     Drop down
        'ListBox
        '     List box
        '     Drop down
        '     Horizontal radio buttons
        '     Vertical radio buttons
        'Numeric Range
        '     Text box
        'Likert (effectively a subset of List Box data type, i.e. mandated 5 items)
        '     Horizontal radio buttons
        '     Vertical radio buttons
                                    
        Dim abc As New System.Collections.Generic.List(Of Long)
        Select Case dst.Rows(0).Item("DataType")
          Case 0, 4, 6 ' String , ListBox, Likert   ' not all presentation styles are valid for each of these data types, but those combinations are restricted in the user defined fields definition
            ' however, since these data types don't require any special handling, I'm consolidating them here.
            If dst.Rows(0).Item("UDFPS_ID") = -1 Then
              If dst.Rows(0).Item("DataType") = 0 Then
                Send("<input type=""text"" class=""short""  id = ""udfVal-" & UDFPK & """ name = ""udfVal-" & UDFPK & """ disabled=""disabled""  value ="""" />")
                Send("<input type=""button"" class=""regular"" name = ""udfVal-" & UDFPK & """ id=""udfVal-" & UDFPK & """  " & IIf(AllowEditing, "", "disabled=""disabled""") & " value=""..."" title=""Click here to edit the text""  style=""FONT-SIZE: 8pt; WIDTH: 20px; HEIGHT: 24px"" onclick=""javascript:assigntextboxcontrol(this, 'foldercreate', true);"" />")
              End If
            Else
              Select Case dst.Rows(0).Item("UDFPS_ID")
                Case 1 'drop down list
                  Send(PresentationStyleUI.createDropDown(UDFPK, udfValuesdst, -1, AllowEditing))
                Case 2 'Horizontal radio
                  Send(PresentationStyleUI.createRadioButtons(UDFPK, udfValuesdst, True, -1, AllowEditing))
                Case 3 'Vertical radio
                  Send(PresentationStyleUI.createRadioButtons(UDFPK, udfValuesdst, False, -1, AllowEditing))
                Case 4 'listbox
                  Send(PresentationStyleUI.createListBox(UDFPK, udfValuesdst, AllowEditing, abc))
                Case 6, -1 'text box
                  'MyCommon.QueryStr = " Select value from UserDefinedField_ValidValues where UDFPK = " + Convert.ToString(dst.Rows(0).Item("UDFPK")) + " and isDefault = 1 "
                  'Dim df As DataTable = MyCommon.LRT_Select
                             
                  Dim valueString As String
                  If udfValuesdst.Rows.Count = 1 Then
                    If Convert.ToBoolean(udfValuesdst.Rows(0).Item("isDefault")) Then
                      'save it to the temp table too.
                      valueString = Convert.ToString(udfValuesdst.Rows(0).Item("value"))
                      UpdateUDFStringValues(UDFPK, OfferID, valueString)
                      Send(PresentationStyleUI.createTextBox(UDFPK, New System.Data.DataTable, AllowEditing, valueString))
                                              
                                              
                    Else
                      Send(PresentationStyleUI.createTextBox(UDFPK, New System.Data.DataTable, AllowEditing, ""))
                    End If
                  Else
                    Send(PresentationStyleUI.createTextBox(UDFPK, New System.Data.DataTable, AllowEditing, ""))
                  End If
              End Select
              End If
            Case 7 ' Image URL
              If dst.Rows(0).Item("UDFPS_ID") = -1 Then
                If dst.Rows(0).Item("DataType") = 0 Then
                  Send("<input type=""text"" class=""short""  id = ""udfVal-" & UDFPK & """ name = ""udfVal-" & UDFPK & """ disabled=""disabled""  value ="""" />")
                  Send("<input type=""button"" class=""regular"" name = ""udfVal-" & UDFPK & """ id=""udfVal-" & UDFPK & """  " & IIf(AllowEditing, "", "disabled=""disabled""") & " value=""..."" title=""Click here to edit the text""  style=""FONT-SIZE: 8pt; WIDTH: 20px; HEIGHT: 24px"" onclick=""javascript:assigntextboxcontrol(this, 'foldercreate', true);"" />")
                End If
              Else
                If udfValuesdst.Rows.Count = 1 Then
                  If Convert.ToBoolean(udfValuesdst.Rows(0).Item("isDefault")) Then
                    'save it to the temp table too.
                    Dim sValue As String
                    sValue = Convert.ToString(udfValuesdst.Rows(0).Item("value"))
                    UpdateUDFStringValues(UDFPK, OfferID, sValue)
                  Send(PresentationStyleUI.createUrlBox(UDFPK, New System.Data.DataTable, AllowEditing, sValue.Replace(System.Convert.ToChar(34).ToString(), "&quot;")))
                    Dim sValue1 As String = "/logix/show-image.aspx?caller=udf&src=" & sValue
                    Send("<img align=""center"" src=""" & sValue1 & """ id=""Image_" & UDFPK & """ width=""50"" height=""50"" title=""Click to view full-sized image"" onclick=""showFullSizedImage('" + sValue1 & "');""" + " />")
                  Else
                  Send(PresentationStyleUI.createUrlBox(UDFPK, New System.Data.DataTable, AllowEditing, ""))
                  End If
                Else
                Send(PresentationStyleUI.createUrlBox(UDFPK, New System.Data.DataTable, AllowEditing, ""))
                End If
            End If
          Case 1 ' Int      
            If dst.Rows(0).Item("UDFPS_ID") = -1 Then
              Send("<input type=""text"" id = ""udfVal-" & UDFPK & """ name = ""udfVal-" & UDFPK & """" & IIf(AllowEditing, "", "disabled=""disabled""") & " maxlength=""11"" value ="""" />")
            Else
              Select Case dst.Rows(0).Item("UDFPS_ID")
                Case 1 'drop down list 
                  Send(PresentationStyleUI.createDropDown(UDFPK, udfValuesdst, -1, AllowEditing))
                Case 4 'listbox
                  Send(PresentationStyleUI.createListBox(UDFPK, udfValuesdst, AllowEditing, abc))
                Case 6, -1 'text box
                  'Send("<input type=""text"" id = ""udfVal-" & UDFPK & """ name = ""udfVal-" & UDFPK & """" & IIf(AllowEditing, "", "disabled=""disabled""") & " maxlength=""11"" value ="""" />")
                  If udfValuesdst.Rows.Count = 1 Then 'look to see if there's a default
                    If Convert.ToBoolean(udfValuesdst.Rows(0).Item("isDefault")) Then
                      Send(PresentationStyleUI.createNumberTextBox(UDFPK, Convert.ToInt64(udfValuesdst.Rows(0).Item("value")), AllowEditing, 11))
                    Else
                      Send(PresentationStyleUI.createNumberTextBox(UDFPK, "", AllowEditing, 11))
                    End If
                  Else
                    Send(PresentationStyleUI.createNumberTextBox(UDFPK, "", AllowEditing, 11))
                  End If
              End Select
            End If
            'Send("<input type=""text"" id = ""udfVal-" & UDFPK & """ name = ""udfVal-" & UDFPK & """" & IIf(AllowEditing, "", "disabled=""disabled""") & " maxlength=""11"" value ="""" />")
          Case 2 'Date
            If dst.Rows(0).Item("UDFPS_ID") = -1 Then
              Send("<input class=""short"" id=""udfVal-" & UDFPK & """ name=""udfVal-" & UDFPK & """" & IIf(AllowEditing, "", "disabled=""disabled""") & " maxlength=""10"" type=""text""  />")
              Send("<img src=""/images/calendar.png"" class=""calendar"" id=""udf-datevalue-picker"" alt=""" & Copient.PhraseLib.Lookup("term.datepicker", LanguageID) & """ title=""" & Copient.PhraseLib.Lookup("term.datepicker", LanguageID) & """ " & IIf(AllowEditing, "onclick=""displayDatePicker('udfVal-" & UDFPK & "', event);""", "") & " />")
            Else
              Send(PresentationStyleUI.createDate(UDFPK, udfValuesdst, AllowEditing, LanguageID, ""))
            End If
          Case 3 'bool
            If dst.Rows(0).Item("UDFPS_ID") = -1 Then
              Send("<input type=""checkbox"" id= ""udfVal-" & UDFPK & """name = ""udfVal-" & UDFPK & """ " & IIf(AllowEditing, "", "disabled=""disabled""") & "/>")
              Send(Copient.PhraseLib.Lookup("term.trueifchecked", LanguageID))
            Else
                              
              Select Case dst.Rows(0).Item("UDFPS_ID")
                Case 2 'Horizontal radio
                  Send(PresentationStyleUI.createRadioButtons(UDFPK, udfValuesdst, True, -1, AllowEditing))
                Case 3 'Vertical radio
                  Send(PresentationStyleUI.createRadioButtons(UDFPK, udfValuesdst, False, -1, AllowEditing))
                Case 5 'Hoizontal CheckBox      
                  Send(PresentationStyleUI.createCheckBox(UDFPK, New DataTable, AllowEditing, LanguageID))
                Case 1 'Drop down
                  Send(PresentationStyleUI.createDropDown(UDFPK, udfValuesdst, -1, AllowEditing))
              End Select
            End If
          Case 5 ' Numeric range
            Send(PresentationStyleUI.createNumericRange(UDFPK, udfValuesdst, "", AllowEditing))
        End Select
        Send("  </td>")
        'Send("    <td>")
        'Send("    </td>")
        Send("</tr>")
      End If
			
    Catch ex As Exception
      Send(ex.ToString)
    Finally
      MyCommon.Close_LogixRT()
      MyCommon = Nothing
    End Try
  End Sub
      
  Sub UpdateUDFStringValues(ByVal UDFPK As Long, ByVal OfferID As Long, ByVal UDFStringText As String)
    Dim SpecialCharacters As String = MyCommon.Fetch_SystemOption(171)
    UDFStringText = CleanString(UDFStringText, SpecialCharacters)
    If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
    Try
      MyCommon.QueryStr = "dbo.pt_UDFStringValuesUpdate"
      MyCommon.Open_LRTsp()
      MyCommon.LRTsp.Parameters.Add("@UDFPK", SqlDbType.BigInt).Value = UDFPK
      MyCommon.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Value = OfferID
      MyCommon.LRTsp.Parameters.Add("@Value", SqlDbType.NVarChar, 1000).Value = UDFStringText
      MyCommon.LRTsp.ExecuteNonQuery()
      MyCommon.Close_LRTsp()
    Catch ex As Exception
      Send(ex.ToString())
    Finally
      'MyCommon.Close_LogixRT()
    End Try
  End Sub
      
  Sub deleteUDFfromOffer(ByVal OfferID As Int64, ByVal UDFPK As Int64)
    Dim dst As DataTable
    Dim row As DataRow
		
    Try
      If MyCommon.LRTadoConn.State = ConnectionState.Closed Then MyCommon.Open_LogixRT()
      'MyCommon.QueryStr = "delete from UserDefinedFieldsValues where UDFPK = " & UDFPK & " and OfferID = " & OfferID
      MyCommon.QueryStr = "Update UserDefinedFieldsValues set deleted = 1 where  UDFPK = " & UDFPK & " and OfferID = " & OfferID
      MyCommon.LRT_Execute()
      MyCommon.QueryStr = "select Description from UserDefinedFields where UDFPK = " & UDFPK
      dst = MyCommon.LRT_Select
      'MyCommon.Activity_Log(3, OfferID, AdminUserID,"Deleted User Defined Field: " &  MyCommon.TruncateString(dst.Rows(0).Item("Description"), 20))
			
      'get options for the udf selector
      MyCommon.QueryStr = "select udf.UDFPK, udf.Description from UserDefinedFields as udf " & _
                "where not exists (select UDFPK from UserDefinedFieldsValues as v where deleted = 0 and udf.UDFPK = v.UDFPK and v.OfferID = " & OfferID & ")"
      dst = MyCommon.LRT_Select
      MyCommon.QueryStr = "delete from OfferUDFStringValues where OfferID = " & OfferID & " and UDFPK = " & UDFPK
      MyCommon.LRT_Execute()
      Send("<select class=""medium"" id=""UDFDataType"" name=""UDFDataType"">")
      For Each row In dst.Rows
        Send("<option value=""UDF-" & row.Item("UDFPK") & """ >" & row.Item("Description") & "</option>")
      Next
      Send("</select>")
    Catch ex As Exception
      Send(ex.ToString)
    Finally
      MyCommon.Close_LogixRT()
      MyCommon = Nothing
    End Try
  End Sub
  </script>