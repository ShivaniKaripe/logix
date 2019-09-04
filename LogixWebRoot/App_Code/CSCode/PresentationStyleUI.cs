
// version:7.3.1.138972.Official Build (SUSDAY10202) %>

//' *****************************************************************************
//  ' * FILENAME: PresentationStyleUI.cs
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



using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;

/// <summary>
/// Summary description for PresentationStyleUI
/// </summary>
public class PresentationStyleUI : BaseUI
{
	public PresentationStyleUI()
	{		
	}

    public static String createDropDown(Int64 UDFPK, DataTable options, Int64 selectedValue, Boolean AllowEditing)
    {
        String htmlString = "";
        String titleString = "";
        DataRow[] selectedRows;

        if (selectedValue == -1)
        {
           selectedRows = options.Select("IsDefault=true");
        }
        else
        {
           selectedRows = options.Select(String.Format("UDFVPK={0}", selectedValue));
        }

        if (selectedRows.Length > 0)
        {
           titleString="title=\"" + Convert.ToString(selectedRows[0]["Value"]).Trim().Replace(System.Convert.ToChar(34).ToString(), "&quot;") + "\"";
        }
        else if (options.Rows.Count > 0)
        {
           titleString = "title=\"" + Convert.ToString(options.Rows[0]["Value"]).Trim().Replace(System.Convert.ToChar(34).ToString(), "&quot;") + "\"";
        }       

        htmlString += "<select class=\"td-udfvalue\"  id=\"udfVal-" + Convert.ToString(UDFPK) + "\" " + titleString + " name=\"udfVal-" + Convert.ToString(UDFPK) + "\"   " + (AllowEditing ? " " : " disabled=\"disabled\" ") + "  onchange=\"onUDFListBoxChange(this);\" >";
        foreach (DataRow row in options.Rows)
        {
            String isSelected = " ";
            if (Convert.ToBoolean(row["IsDefault"]) && selectedValue==-1)
            {
                isSelected = " selected ";
            }
            else if (Convert.ToInt64(row["UDFVPK"]) == selectedValue)
            {
                isSelected = " selected "; 
            }
            htmlString += "<option id=\"udfvalidValue-" + Convert.ToString(row["UDFVPK"]) + "\" name=\"udfvalidValue-" + Convert.ToString(row["UDFVPK"]) + "\" value=\"" + Convert.ToString(row["UDFVPK"]) + "\" " + isSelected + " >" + Convert.ToString(row["Value"]).Trim() + "</option>";
        }
        htmlString += "</select>";
        return htmlString;
    }

    public static String createListBox(Int64 UDFPK, DataTable options, Boolean AllowEditing, List<Int64> selectedValues = null) 
    {
          String htmlString="";

          Copient.CommonInc MyCommon = new Copient.CommonInc();

          htmlString += "<select multiple class=\"td-udfvalue\"  id=\"udfVal-" + UDFPK + "\" name=\"udfVal-" + UDFPK + "\"  size=\"" + options.Rows.Count + "\"   " + (AllowEditing ? " " : " disabled=\"disabled\" ") + " onchange=\"onUDFListBoxChange(this);\" >";
          foreach(DataRow row in options.Rows)
          {
              String isSelected = " ";
              if (Convert.ToBoolean(row["IsDefault"]) && (selectedValues == null || selectedValues.Count==0))
              {
                  isSelected = " selected ";
              }
              else if (-1 != selectedValues.FindIndex(f=>f == Convert.ToInt64(row["UDFVPK"])))
              {
                  isSelected = " selected ";
              }
              htmlString += "<option title=\"" + HttpUtility.HtmlEncode(Convert.ToString(row["Value"]).Trim()) + "\" id=\"udfvalidValue-" + Convert.ToString(row["UDFVPK"]) + "\" name=\"udfvalidValue-" + Convert.ToString(row["UDFVPK"]) + "\" value=\"" + Convert.ToString(row["UDFVPK"]) + "\" " + isSelected + " >" + HttpUtility.HtmlEncode(Convert.ToString(row["Value"]).Trim()) + "</option>";
          }
          htmlString += "</select>";
          return htmlString;
     }

    public static String createCheckBox(Int64 UDFPK, DataTable options, Boolean AllowEditing, int LanguageID,bool isChecked=false)
    {
        String htmlString = "";

        htmlString += "<input type=\"checkbox\" id=\"udfVal-" + Convert.ToString(UDFPK) + "\" title=\"" + Copient.PhraseLib.Lookup("term.trueifchecked", LanguageID) + "\" name=\"udfVal-" + Convert.ToString(UDFPK) + "\" " + (AllowEditing ? " " : " disabled=\"disabled\" ") + (isChecked?" checked ":" ") + "/>";
        htmlString += Copient.PhraseLib.Lookup("term.trueifchecked", LanguageID);

        return htmlString;
    }

    public static String createTextBox(Int64 UDFPK, DataTable options, Boolean AllowEditing, String stringValue)
    {
        String htmlString = "";
        Copient.CommonInc MyCommon = new Copient.CommonInc();
        String initialValue;

        stringValue = stringValue.Replace(System.Convert.ToChar(34).ToString(), "&#34;");
        initialValue = MyCommon.TruncateString(stringValue, 10);

        htmlString += "<input type=\"text\" class=\"short\"  id=\"udfVal-" + Convert.ToString(UDFPK) + "\" name=\"udfVal-" + Convert.ToString(UDFPK) + "\" disabled=\"disabled\"  value =\""+initialValue +"\" title=\"" + stringValue + "\" />";
        htmlString += "<input type=\"button\" class=\"regular\" name=\"udfVal-" + Convert.ToString(UDFPK) + "\" id=\"udfVal-" + Convert.ToString(UDFPK) + "\"  " + (AllowEditing ? " " : " disabled=\"disabled\" ") + " value=\"...\" title=\"Click here to edit the text\"  style=\"FONT-SIZE: 8pt; WIDTH: 20px; HEIGHT: 24px\" onclick=\"javascript:assigntextboxcontrol(this, 'foldercreate', true);\" />";

        return htmlString;
    }

    public static String createUrlBox(Int64 UDFPK, DataTable options, Boolean AllowEditing, String initialValue)
    {
      String htmlString = "";

      htmlString += "<input type=\"text\" class=\"short\"  id=\"udfVal-" + Convert.ToString(UDFPK) + "\" name=\"udfVal-" + Convert.ToString(UDFPK) + "\" disabled=\"disabled\"  value =\"" + initialValue + "\" title=\"" + initialValue + "\" />";
      htmlString += "<input type=\"button\" class=\"regular\" name=\"udfVal-" + Convert.ToString(UDFPK) + "\" id=\"udfVal-" + Convert.ToString(UDFPK) + "\"  " + (AllowEditing ? " " : " disabled=\"disabled\" ") + " value=\"...\" title=\"Click here to edit the text\"  style=\"FONT-SIZE: 8pt; WIDTH: 20px; HEIGHT: 24px\" onclick=\"javascript:assigntextboxcontrol(this, 'foldercreate', true);\" />";

      return htmlString;
    }

    public static String createNumberTextBox(Int64 UDFPK,String intValue ,Boolean AllowEditing,int maxLength)
    {
        String htmlString = "";

        htmlString += "<input type=\"text\" id=\"udfVal-" + Convert.ToString(UDFPK) + "\" title=\"" + intValue + "\" name=\"udfVal-" + Convert.ToString(UDFPK) + "\" " + (AllowEditing? " ": " disabled=\"disabled\" ") + " maxlength=\""+ Convert.ToString(maxLength) +"\" value=\"" + intValue + "\" onkeypress=\"return isNumber(this,event);\" />";
        return htmlString;
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="UDFPK"></param>
    /// <param name="options"></param>
    /// <param name="intRange"> "{minValue:maxValue}"
    /// </param>
    /// <returns></returns>
    public static String createNumericRange(Int64 UDFPK, DataTable options,String intRange,Boolean AllowEditing)
    {
        String htmlString = "";
        String label = "";
        string defaultValue = "";
        foreach (DataRow dr in options.Rows)
        {
            if (label != String.Empty)
            {
                label += ",";
            }
            label += Convert.ToString(dr["Value"]).Trim();

            if (Convert.ToBoolean(dr["IsDefault"]))
            {
                defaultValue = Convert.ToString(dr["Value"]);
            }
        }

        if (intRange != String.Empty)
        {
            defaultValue = intRange;
        }
        defaultValue = defaultValue.Trim();
        htmlString += htmlString += "<input type=\"text\" id=\"udfVal-" + Convert.ToString(UDFPK) + "\" title=\"" + defaultValue + "\" name=\"udfVal-" + Convert.ToString(UDFPK) + "\"  maxlength=\"100\" value=\"" + defaultValue + "\" onkeypress=\"return isNumber(this,event);\" onblur=\"return validateNumericRange(this,event)\"   " + (AllowEditing ? " " : " disabled=\"disabled\" ") + " />";
        htmlString += "<br><label id=\"udfVal-" + Convert.ToString(UDFPK) + "_label\" for=\"udfVal-" + Convert.ToString(UDFPK) + "\">" + label + "</label>";
        return htmlString;
    }

    public static String createDate(Int64 UDFPK, DataTable options, Boolean AllowEditing, int LanguageID,String initialValue)
      {
          String htmlString = "";
          String valueToUse = initialValue;
          if (options.Rows.Count == 1 && initialValue==String.Empty)
          {
              if (Convert.ToBoolean(options.Rows[0]["isDefault"]) == true)
              {
                  valueToUse = Convert.ToString(options.Rows[0]["Value"]);
              }
          }
          valueToUse = valueToUse.Trim();
          htmlString += "<input class=\"short\" id=\"udfVal-" + Convert.ToString(UDFPK) + "\" title=\"" + valueToUse + "\" name=\"udfVal-" + Convert.ToString(UDFPK) + "\"" + (AllowEditing? " ": " disabled=\"disabled\" ") + " maxlength=\"10\" type=\"text\"  value=\""+valueToUse+"\"/>";
          htmlString += "<img src=\"/images/calendar.png\" class=\"calendar\" id=\"udf-datevalue-picker\" alt=\"" + Copient.PhraseLib.Lookup("term.datepicker", LanguageID) + "\" title=\"" + Copient.PhraseLib.Lookup("term.datepicker", LanguageID) + "\" " + (AllowEditing? " onclick=\"displayDatePicker('udfVal-" + Convert.ToString(UDFPK) + "', event);\" " :  " ") + " />";
          return htmlString;
      }


    /*
     * 
     * Horizontal radio
     <table>
         <tr><td/><td/><td/>...</tr>
         <tr><td/><td/><td/>...</tr>
     </table>
     
     * 
     Vertical radio
     <table>
         <tr><td/><td/></tr>
         <tr><td/><td/></tr>
     * ...
     * ...
     * ...
     </table>
     */
    public static String createRadioButtons(Int64 UDFPK, DataTable options, Boolean isHorizontal, Int64 selectedValue, Boolean AllowEditing) 
      {
          String htmlString = "";
          Copient.CommonInc MyCommon = new Copient.CommonInc();
          
          if (isHorizontal)
          {
              String labelRow = "";
              String radioRow = "";

              foreach (DataRow row in options.Rows)
              {
                  String isSelected = " ";
                  if (Convert.ToBoolean(row["IsDefault"]) && selectedValue == -1)
                  {
                      isSelected = " checked ";
                  }
                  else if (Convert.ToInt64(row["UDFVPK"]) == selectedValue)
                  {
                      isSelected = " checked ";
                  }

                  labelRow += "<td style=\"text-align:center\">";
                  labelRow += "<label title=\"" + HttpUtility.HtmlEncode(Convert.ToString(row["Value"])) + "\" for=\"udfvalidValue-" + Convert.ToString(row["UDFVPK"]) + "\">" + HttpUtility.HtmlEncode(MyCommon.TruncateString(Convert.ToString(row["Value"]).Trim(), (27 / options.Rows.Count))) + "</label>";
                  labelRow += "</td>";
                  
                  radioRow += "<td style=\"text-align:center\">";
                  radioRow += "<input type=\"radio\" id=\"udfvalidValue-" + Convert.ToString(row["UDFVPK"]) + "\" name=\"udfvalidValue-" + Convert.ToString(UDFPK) + "\" value=\"" + Convert.ToString(row["UDFVPK"]) + "\" " + isSelected + "  " + (AllowEditing ? " " : " disabled=\"disabled\" ") + "  />"; ;
                  radioRow += "</td>";
              }
              //htmlString = "<table style=\"overflow: hidden; text-overflow: ellipsis;table-layout: fixed\" nowrap><tr>" + labelRow + " </tr><tr>" + radioRow + "</tr></table>";
              htmlString = "<table style=\"overflow: hidden; table-layout: fixed\" ><tr class=\"td-udfvalue\">" + labelRow + " </tr><tr class=\"td-udfvalue\">" + radioRow + "</tr></table>";
          }
          else
          {
              htmlString = "<table>";
              foreach (DataRow row in options.Rows)
              {
                  String isSelected = " ";
                  if (Convert.ToBoolean(row["IsDefault"]) && selectedValue == -1)
                  {
                      isSelected = " checked ";
                  }
                  else if (Convert.ToInt64(row["UDFVPK"]) == selectedValue)
                  {
                      isSelected = " checked ";
                  }
                  htmlString += "<tr class=\"td-udfvalue\" >";
                  htmlString += "<td style=\"text-align:left\"><input style=\"padding-left:0;padding-right:0\"type=\"radio\" id=\"udfvalidValue-" + Convert.ToString(row["UDFVPK"]) + "\" name=\"udfvalidValue-" + Convert.ToString(UDFPK) + "\" value=\"" + Convert.ToString(row["UDFVPK"]) + "\" " + isSelected + "   " + (AllowEditing ? " " : " disabled=\"disabled\" ") + " /></td>";
                  htmlString += "<td style=\"text-align:left\"><label title=\"" + HttpUtility.HtmlEncode(Convert.ToString(row["Value"])) + "\"  for=\"udfvalidValue-" + Convert.ToString(row["UDFVPK"]) + "\">" + HttpUtility.HtmlEncode(MyCommon.TruncateString(Convert.ToString(row["Value"]).Trim(), 22)) + "</label></td>";
                  htmlString += "</tr>";
              }
              htmlString += "</table>";
          }
          return htmlString;
      }
}