<%@ Control Language="C#" AutoEventWireup="true" CodeFile="UDFListControl.ascx.cs"
  Inherits="logix_UserControls_UDFListControl" %>
<% // version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%
  //' *****************************************************************************
  //  ' * FILENAME: UDFListControl.ascx 
  //  ' *~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  //  ' * Copyright © 2002 - 2014.  All rights reserved by:
  //  ' *
  //  ' * NCR Corporation
  //  ' * 2651 Satellite Blvd
  //  ' * Duluth, GA 30096     
  //  ' *~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  //  ' *
  //  ' * PROJECT : NCR Advanced Marketing Solution
  //  ' *
  //  ' * MODULE  : Logix
  //  ' *
  //  ' * PURPOSE : 
  //  ' *
  //  ' * NOTES   : 
  //  ' *
  //  ' * Version : 7.3.1.138972 
  //  ' *
  //  ' *****************************************************************************
%>
<div class="box" id="UserDefinedFields">
  <h2>
    <span>User Defined Fields</span>
  </h2>
  <%if (IsTemplate)
    {%>
  <span class="temp">
    <input type="checkbox" class="tempcheck" id="temp-udf" name="Disallow_UserDefinedFields"
      <% if(Disallow_UserDefinedFields){ Response.Write(" checked=\"checked\""); }%><% if(bUseTemplateLocks) {Response.Write(" disabled=\"disabled\"");} %> />
    <label for="temp-udf">
      <% Response.Write(Copient.PhraseLib.Lookup("term.locked", LanguageID));%></label>
  </span>
  <br class="printonly" />
  <% }%>
  <%        
    if (bUseTemplateLocks == true && Disallow_UserDefinedFields == true)
    {
      disableUDF = "disabled";
    }
  %>
  <% 
    bool isTranslatedOffer = lCommon.IsTranslatedUEOffer(Convert.ToInt32(OfferID), lCommon);
    bool bEnableRestrictedAccessToUEOfferBuilder = lCommon.Fetch_SystemOption(249) == "1" ? true : false;
    bool bEnableAdditionalLockoutRestrictionsOnOffers = lCommon.Fetch_SystemOption(260) == "1" ? true : false;
    bool bOfferEditable = lCommon.IsOfferEditablePastLockOutPeriod(EditOfferPastLockoutdate, lCommon, Convert.ToInt32(OfferID));
    if (bEnableRestrictedAccessToUEOfferBuilder && isTranslatedOffer)
      AllowEditing = false;
    if (lLogix.UserRoles.AddUserDefinedFields)
    {
      Response.Write(@"<div id = ""udfSelect"" style=""display:inline;"">");
      Response.Write(@"<select class=""medium"" id=""UDFDataType"" name=""UDFDataType"" " + disableUDF + "  >");
      lCommon.QueryStr = @"select udf.UDFPK, udf.Description from UserDefinedFields as udf where not exists (select UDFPK from UserDefinedFieldsValues as v where deleted = 0 and udf.UDFPK = v.UDFPK and v.OfferID = " + OfferID + ")";

      //lCommon.Write_Event_Log(lCommon.QueryStr,"UDFListControl.ascx");
      System.Data.DataTable rst2 = lCommon.LRT_Select();
      foreach (System.Data.DataRow row2 in rst2.Rows)
      {
        Response.Write(@"<option value=""UDF-" + row2["UDFPK"] + @""" >" + row2["Description"] + "</option>");
      }
      Response.Write(@"</select>");
      Response.Write(@"</div>");
      if((bEnableRestrictedAccessToUEOfferBuilder && isTranslatedOffer) || (bEnableAdditionalLockoutRestrictionsOnOffers && !bOfferEditable))
        Response.Write(@"<input type=""button"" value=""" + Copient.PhraseLib.Lookup("term.add", LanguageID) + @"""  " + @"disabled=""disabled""" + @" onclick=""javascript:addUDF(" + Server.HtmlEncode(OfferID) + @");"" />");
      else
        Response.Write(@"<input type=""button"" value=""" + Copient.PhraseLib.Lookup("term.add", LanguageID) + @"""  " + ((rst2.Rows.Count == 0) || (disableUDF != "") ? "disabled=true" : "") + @" onclick=""javascript:addUDF(" + Server.HtmlEncode(OfferID) + @");"" />");
      Response.Write(@"<br />");
      Response.Write(@"<br />");
      Response.Write(@"<br />");
    }

    Response.Write(@"<table id=""udftable"" summary=""" + Copient.PhraseLib.Lookup("term.values", LanguageID) + @""">");
    Response.Write(@"<thead id=""udfhead"">");
    Response.Write("  <tr>");
    if (lLogix.UserRoles.DeleteUserDefinedFields)
    {
      if (disableUDF == "")
      {
        Response.Write(@"    <th style=""width:32px;"">" + (Copient.PhraseLib.Lookup("term.delete", LanguageID)).Substring(0, 3) + "</th>");
      }
    }
    Response.Write(@"    <th class=""th-udfdescription"">" + Copient.PhraseLib.Lookup("term.description", LanguageID) + "</th>");
    Response.Write(@"    <th class=""th-udfvalue"">" + Copient.PhraseLib.Lookup("term.value", LanguageID) + "</th>");
    Response.Write(@"  </tr>");
    Response.Write(@" </thead>");
    Response.Write(@"<tbody   id = ""UDFList"">");


    //              lCommon.QueryStr = @"select udf.UDFPK, udf.Description, offer.StringValue, offer.IntValue, offer.DateValue, offer.BooleanValue, offer.UDFVPK, udf.DataType,udf.UDFPS_ID from userdefinedfieldsvalues 
    //		                        as offer with (NoLock) inner join UserDefinedFields as udf on offer.UDFPK = udf.UDFPK where deleted = 0 and offer.OfferID = " + OfferID;

    lCommon.QueryStr = @"select udf.UDFPK, udf.Description, offer.StringValue, offer.IntValue, offer.DateValue, offer.BooleanValue, offer.UDFVPK, udf.DataType,udf.UDFPS_ID from userdefinedfieldsvalues 
		                        as offer with (NoLock) inner join UserDefinedFields as udf on offer.UDFPK = udf.UDFPK where deleted = 0 and offer.OfferID =" + OfferID + @" 
		                        and udf.udfpk in (
                                    select udfpk from UserDefinedFieldsValues where offerid = " + OfferID + @"  group by udfpk having COUNT (udfpk) = 1 )

                                UNION

                                select DISTINCT udf.UDFPK, udf.Description, '' as StringValue, 0 as IntValue, '1/1/2000' as DateValue, 0 as BooleanValue, -1 as UDFVPK, udf.DataType,udf.UDFPS_ID from userdefinedfieldsvalues 
		                                                        as offer with (NoLock) inner join UserDefinedFields as udf on offer.UDFPK = udf.UDFPK where deleted = 0 and offer.OfferID =" + OfferID + @"
		                                                        and udf.udfpk in (
                                select udfpk from UserDefinedFieldsValues where offerid = " + OfferID + @" group by udfpk having COUNT (udfpk) > 1 )
                                ;";





    //lCommon.Write_Event_Log(lCommon.QueryStr, "UDFListControl.ascx");                 
    System.Data.DataTable udfrst = lCommon.LRT_Select();
    //'''BEGIN User Defined Fields UI initialization block
    int rowcount = 0;
    if (udfrst.Rows.Count > 0)
    {
      foreach (System.Data.DataRow row in udfrst.Rows)
      {
        Response.Write(@"  <tr id = ""TRudfVal-" + row["UDFPK"] + @""">");
        if (lLogix.UserRoles.DeleteUserDefinedFields)
        {
          if (disableUDF == "")              //'to delete not only user rights, if the offer was created from a template, template permissions also should allow. Template permissions override user rights
          {
            if((bEnableRestrictedAccessToUEOfferBuilder && isTranslatedOffer) || (bEnableAdditionalLockoutRestrictionsOnOffers && !bOfferEditable))
              Response.Write(@"    <td><input type=""button"" value=""X"" disabled=""disabled"" title=""" + Copient.PhraseLib.Lookup("term.delete", LanguageID) + @""" name=""ex"" id=""ex-" + lCommon.NZ(row["UDFPK"], 0) + @""" class=""ex""" + @" onclick=""javascript:deleteUDF(" + lCommon.NZ(row["UDFPK"], 0) + ", " + Server.HtmlEncode(OfferID) + @")"" /></td>");
            else
              Response.Write(@"    <td><input type=""button"" value=""X"" title=""" + Copient.PhraseLib.Lookup("term.delete", LanguageID) + @""" name=""ex"" id=""ex-" + lCommon.NZ(row["UDFPK"], 0) + @""" class=""ex""" + @" onclick=""javascript:deleteUDF(" + lCommon.NZ(row["UDFPK"], 0) + ", " + Server.HtmlEncode(OfferID) + @")"" /></td>");
          }
        }


        Response.Write(@"    <td><span title=""" + Convert.ToString(lCommon.NZ(row["Description"], "")).Replace(System.Convert.ToChar(34).ToString(), "&quot;") + @""" style=""display: inline-block; width: 128px;min-width: 20px; max-width: 128px;   overflow: hidden;   text-overflow: ellipsis;"">" + lCommon.TruncateString(Convert.ToString(lCommon.NZ(row["Description"], "")), 27) + @"</span></td>");
        //Response.Write(@"    <td><span style=""display: inline-block; width: 128px;min-width: 20px; max-width: 128px;   overflow: hidden;   text-overflow: ellipsis;"">" + Convert.ToString(lCommon.NZ(row["Description"], "")) + @"</span></td>");
        Response.Write("    <td class=\"td-udfvalue\">");


        long valueSelection = -1;
        List<long> valueSelectionArray = null;

        System.Data.DataTable udfValuesdst = new System.Data.DataTable();
        if (row["UDFPS_ID"] != System.DBNull.Value)
        {
          lCommon.QueryStr = @"Select UDFVPK,UDFPK,Value,IsDefault,DisplayOrder from UserDefinedField_ValidValues where UDFPK=" + row["UDFPK"] + @" order by DisplayOrder";

          udfValuesdst = lCommon.LRT_Select();
          if (row["UDFVPK"] == System.DBNull.Value)
          {
            valueSelection = -1;
          }
          else
          {
            if (row["UDFPS_ID"] != System.DBNull.Value && Convert.ToInt64(row["UDFPS_ID"]) == 4)//possibly have a multi selected list box
            {
              lCommon.QueryStr = "Select UDFVPK from UserDefinedFieldsValues where UDFPK=" + row["UDFPK"] + " and OfferID = " + OfferID;
              lCommon.Open_LogixRT();
              System.Data.DataTable multiValue = lCommon.LRT_Select();

              if (multiValue.Rows.Count > 0)
              {
                valueSelectionArray = new List<long>();
                foreach (System.Data.DataRow dr in multiValue.Rows)
                {
                  valueSelectionArray.Add(Convert.ToInt64(dr["UDFVPK"]));
                }
              }
            }
            else
            {
              valueSelection = Convert.ToInt64(row["UDFVPK"]);
            }
          }
        }
        else
        {
        }
        //''add html elements

        switch (System.Convert.ToInt64(row["DataType"]))
        {
          case 0://' String , ListBox, Likert' String
          case 4:
          case 6:
            if (System.Convert.ToInt64(row["DataType"]) == 0 && row["UDFPS_ID"] == System.DBNull.Value) //handle existing data same way it's always been handled
            {
              Response.Write(@"<input type=""text"" class=""short""  id = ""udfVal-" + row["UDFPK"] + @""" name = ""udfVal-" + row["UDFPK"] + @""" disabled=""disabled"" value =""" + lCommon.TruncateString(System.Convert.ToString(lCommon.NZ(row["StringValue"], "")), 10).Replace(System.Convert.ToChar(34).ToString(), "&quot;") + @""" />");
              if (disableUDF != String.Empty)
              {
                Response.Write(@"<input type=""button"" class=""regular"" name = ""udfVal-" + row["UDFPK"] + @""" id=""udfVal-" + row["UDFPK"] + @"""  disabled=""disabled"";                                  value =""..."" title=""Click here to edit the text""  style=""FONT-SIZE: 8pt; WIDTH: 20px; HEIGHT: 24px"" onclick=""javascript:assigntextboxcontrol(this, 'foldercreate', true);"" />");
              }
              else
              {
                Response.Write(@"<input type=""button"" class=""regular"" name = ""udfVal-" + row["UDFPK"] + @""" id=""udfVal-" + row["UDFPK"] + @"""  " + (AllowEditing ? "" : @"disabled=""disabled""") + @" value =""..."" title=""Click here to edit the text""  style=""FONT-SIZE: 8pt; WIDTH: 20px; HEIGHT: 24px"" onclick=""javascript:assigntextboxcontrol(this, 'foldercreate', true);"" />");
              }
            }
            else
            {
              switch (System.Convert.ToInt64(row["UDFPS_ID"]))
              {
                case 1: //'drop down list
                  Response.Write(PresentationStyleUI.createDropDown(System.Convert.ToInt64(row["UDFPK"]), udfValuesdst, valueSelection, AllowEditing && (disableUDF == String.Empty)));
                  break;
                case 2: //'Horizontal radio                                        
                  Response.Write(PresentationStyleUI.createRadioButtons(System.Convert.ToInt64(row["UDFPK"]), udfValuesdst, true, valueSelection, AllowEditing && (disableUDF == String.Empty)));
                  break;
                case 3: //'Vertical radio                                      
                  Response.Write(PresentationStyleUI.createRadioButtons(System.Convert.ToInt64(row["UDFPK"]), udfValuesdst, false, valueSelection, AllowEditing && (disableUDF == String.Empty)));
                  break;
                case 4: //'listbox
                  Response.Write(PresentationStyleUI.createListBox(System.Convert.ToInt64(row["UDFPK"]), udfValuesdst, AllowEditing && (disableUDF == String.Empty), valueSelectionArray));
                  break;
                case 6: //'text box
                case -1:
                  String stringValue = "";
                  if (row["StringValue"] == DBNull.Value)
                  {
                    //we only want to use default value when the UDF is first added to the offer, here we'll take NULL as an empty string
                    stringValue = "";
                  }
                  else
                  {
                    stringValue = Convert.ToString(row["StringValue"]);

                  }
                  Response.Write(PresentationStyleUI.createTextBox(System.Convert.ToInt64(row["UDFPK"]), udfValuesdst, AllowEditing && (disableUDF == String.Empty), stringValue));
                  break;
              }
            }
            break;
          case 7: // ImageURL
            String stringValue1 = "";
            if (row["StringValue"] == DBNull.Value)
            {
              //we only want to use default value when the UDF is first added to the offer, here we'll take NULL as an empty string
              stringValue1 = "";
            }
            else
            {
              stringValue1 = Convert.ToString(row["StringValue"]);
            }
            //stringValue1 = lCommon.TruncateString(stringValue1, 10);
            String stringValue2;
            stringValue2 = PresentationStyleUI.createUrlBox(System.Convert.ToInt64(row["UDFPK"]), udfValuesdst, AllowEditing && (disableUDF == String.Empty), stringValue1);
            Response.Write(stringValue2);
            Response.Write("&nbsp");
            stringValue2 = "/logix/show-image.aspx?caller=udf&src=" + stringValue1;
            Response.Write(@"<img align=""center"" src=""" + stringValue2 + @""" id=""Image_" + Convert.ToString(row["UDFPK"]) + @""" width=""50"" height=""50"" onerror=""this.src='/images/notfound.png'"" title=""Click to view full-sized image"" " + (AllowEditing ? @"onclick=""showFullSizedImage('" + stringValue2 + @"');""" : "") + " />");
            break;
          case 1:// ' Int		               
            if (System.Convert.ToInt64(row["DataType"]) == 1 && row["UDFPS_ID"] == System.DBNull.Value) //handle existing data same way it's always been handled
            {
              if (disableUDF != String.Empty)
              {
                Response.Write(@"<input type=""text"" id = ""udfVal-" + row["UDFPK"] + @""" name = ""udfVal-" + row["UDFPK"] + @""" disabled=""disabled"" maxlength=""11"" value =""" + row["IntValue"] + @""" />");
              }
              else
              {
                Response.Write(@"<input type=""text"" id = ""udfVal-" + row["UDFPK"] + @""" name = ""udfVal-" + row["UDFPK"] + @"""" + (AllowEditing ? "" : @"disabled=""disabled""") + @" maxlength=""11"" value =""" + row["IntValue"] + @""" />");
              }
            }
            else
            {
              switch (System.Convert.ToInt64(row["UDFPS_ID"]))
              {
                case 1:// 'drop down list 
                  Response.Write(PresentationStyleUI.createDropDown(System.Convert.ToInt64(row["UDFPK"]), udfValuesdst, valueSelection, AllowEditing && (disableUDF == String.Empty)));
                  break;
                case 4:// 'listbox
                  Response.Write(PresentationStyleUI.createListBox(System.Convert.ToInt64(row["UDFPK"]), udfValuesdst, AllowEditing && (disableUDF == String.Empty), valueSelectionArray));
                  break;
                case 6:// 'text box
                case -1:
                  Response.Write(PresentationStyleUI.createNumberTextBox(System.Convert.ToInt64(row["UDFPK"]), (row["IntValue"] == System.DBNull.Value ? "" : System.Convert.ToString(row["IntValue"])), AllowEditing && (disableUDF == String.Empty), 11));
                  break;
              }
            }
            break;
          case 2:// 'Date
            if (System.Convert.ToInt64(row["DataType"]) == 2 && row["UDFPS_ID"] == System.DBNull.Value) //handle existing data same way it's always been handled
            {
              if (disableUDF != String.Empty)
              {
                if (row["DateValue"] == System.DBNull.Value)
                {
                  Response.Write(@"<input class=""short"" id=""udfVal-" + row["UDFPK"] + @""" name=""udfVal-" + row["UDFPK"] + @""" disabled=""disabled"" maxlength=""10"" type=""text""  />");
                }
                else
                {
                  Response.Write(@"<input class=""short"" id=""udfVal-" + row["UDFPK"] + @""" name=""udfVal-" + row["UDFPK"] + @""" disabled=""disabled"" maxlength=""10"" type=""text"" value=""" + lLogix.ToShortDateString(System.Convert.ToDateTime(row["DateValue"]), ref lCommon) + @"""  />");
                }
                Response.Write(@"<img src=""/images/calendar.png"" class=""calendar"" id=""udf-datevalue-picker"" alt=""" + Copient.PhraseLib.Lookup("term.datepicker", LanguageID) + @""" title=""" + Copient.PhraseLib.Lookup("term.datepicker", LanguageID) + @"""  />");
              }
              else
              {
                if (row["DateValue"] == System.DBNull.Value)
                {
                  Response.Write(@"<input class=""short"" id=""udfVal-" + row["UDFPK"] + @""" name=""udfVal-" + row["UDFPK"] + @"""" + (AllowEditing ? "" : @"disabled=""disabled""") + @" maxlength=""10"" type=""text""  />");
                }
                else
                {
                  Response.Write(@"<input class=""short"" id=""udfVal-" + row["UDFPK"] + @""" name=""udfVal-" + row["UDFPK"] + @"""" + (AllowEditing ? "" : @"disabled=""disabled""") + @" maxlength=""10"" type=""text"" value=""" + lLogix.ToShortDateString(System.Convert.ToDateTime(row["DateValue"]), ref lCommon) + @"""  />");
                }
                Response.Write(@"<img src=""/images/calendar.png"" class=""calendar"" id=""udf-datevalue-picker"" alt=""" + Copient.PhraseLib.Lookup("term.datepicker", LanguageID) + @""" title=""" + Copient.PhraseLib.Lookup("term.datepicker", LanguageID) + @""" " + (AllowEditing ? @"onclick=""displayDatePicker('udfVal-" + row["UDFPK"] + @"', event);""" : "") + " />");
              }
            }
            else
            {
              Response.Write(PresentationStyleUI.createDate(System.Convert.ToInt64(row["UDFPK"]), udfValuesdst, AllowEditing && (disableUDF == String.Empty), LanguageID, (row["DateValue"] == System.DBNull.Value ? " " : lLogix.ToShortDateString(System.Convert.ToDateTime(row["DateValue"]), ref lCommon))));
            }
            break;
          case 3:// 'bool

            if (System.Convert.ToInt64(row["DataType"]) == 3 && row["UDFPS_ID"] == System.DBNull.Value) //handle existing data same way it's always been handled
            {
              if (disableUDF != String.Empty)
              {
                Response.Write(@"<input type=""checkbox"" id= ""udfVal-" + row["UDFPK"] + @"""name = ""udfVal-" + row["UDFPK"] + @""" disabled=""disabled"" " + (System.Convert.ToBoolean(lCommon.NZ(row["BooleanValue"], false)) == true ? "checked" : "") + "/>");
              }
              else
              {
                Response.Write(@"<input type=""checkbox"" id= ""udfVal-" + row["UDFPK"] + @"""name = ""udfVal-" + row["UDFPK"] + @""" " + (AllowEditing ? "" : @"disabled=""disabled""") + " " + (System.Convert.ToBoolean(lCommon.NZ(row["BooleanValue"], false)) == true ? "checked" : "") + "/>");
              }
              Response.Write(Copient.PhraseLib.Lookup("term.trueifchecked", LanguageID));
            }
            else
            {

              switch (System.Convert.ToInt64(row["UDFPS_ID"]))
              {
                case 2:// 'Horizontal radio
                  Response.Write(PresentationStyleUI.createRadioButtons(System.Convert.ToInt64(row["UDFPK"]), udfValuesdst, true, valueSelection, AllowEditing && (disableUDF == String.Empty)));
                  break;
                case 3:// 'Vertical radio
                  Response.Write(PresentationStyleUI.createRadioButtons(System.Convert.ToInt64(row["UDFPK"]), udfValuesdst, false, valueSelection, AllowEditing && (disableUDF == String.Empty)));
                  break;
                case 5:// 'Hoizontal CheckBox      
                case -1:
                  Response.Write(PresentationStyleUI.createCheckBox(System.Convert.ToInt64(row["UDFPK"]), new System.Data.DataTable(), AllowEditing && (disableUDF == String.Empty), LanguageID,System.Convert.ToBoolean(lCommon.NZ(row["BooleanValue"], false))));
                  break;
                case 1:// 'Drop down
                  Response.Write(PresentationStyleUI.createDropDown(System.Convert.ToInt64(row["UDFPK"]), udfValuesdst, valueSelection, AllowEditing && (disableUDF == String.Empty)));
                  break;
              }
            }
            break;
          case 5:// ' Numeric range
            Response.Write(PresentationStyleUI.createNumericRange(System.Convert.ToInt64(row["UDFPK"]), udfValuesdst, System.Convert.ToString(row["StringValue"]), AllowEditing && (disableUDF == String.Empty)));
            break;
        }
        Response.Write("  </td>");
        //Response.Write("    <td>");
        //Response.Write("    </td>");
        Response.Write("  </tr>");
      }
    }
    //'''END User Defined Fields UI initialization block

    Response.Write("</tbody>");
    Response.Write("</table>");
  %>
</div>