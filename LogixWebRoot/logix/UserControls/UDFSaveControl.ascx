<%@ Control Language="C#" AutoEventWireup="true" CodeFile="UDFSaveControl.ascx.cs" Inherits="logix_UserControls_UDFSaveControl" %>
<% // version:7.3.1.138972.Official Build (SUSDAY10202) %>
<%
//' *****************************************************************************
//  ' * FILENAME: UDFSaveControl.ascx 
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
     
    System.Data.DataTable udfSaverst;
    System.Data.DataTable udfrst;
                            
    
     lCommon.QueryStr = "select V.UDFPK, UDF.DataType, udf.Description, V.deleted,udfps.PresentationStyleID  from UserDefinedFieldsValues as v inner join UserDefinedFields as udf on v.UDFPK = udf.UDFPK left join UserDefinedFieldsPresentationStyles udfps on udfps.UDFPS_ID=udf.UDFPS_ID where OfferID = " + OfferID;
     udfSaverst = lCommon.LRT_Select();

     String value = String.Empty;
     Int32 Status;
     System.Data.DataTable udfvv;
    
    if(udfSaverst.Rows.Count > 0)
    {
        foreach (System.Data.DataRow row in udfSaverst.Rows)
        {
            value = "";
            if (row["PresentationStyleID"] == System.DBNull.Value)//if presentation style is null, just invoke previously existing code
            {
                if (System.Convert.ToInt64(row["DataType"]) == 3)//checkboxes dont return a value when unchecked. 
                {
                    if (Request.QueryString.Get("udfVal-" + System.Convert.ToString(row["UDFPK"]))  != null)
                    {
                        value = "1";
                    }
                    else
                    {
                        value = "0";
                    }
                }

                if (Convert.ToBoolean(row["deleted"]) == true)
                {
                    UDFHistory += Copient.PhraseLib.Lookup("term.deleted", LanguageID) + " " + Copient.PhraseLib.Lookup("term.userdefinedfield", LanguageID) + " - " + lCommon.TruncateString(Convert.ToString(row["Description"]), 20) + ", ";
                    lCommon.QueryStr = "delete from UserDefinedFieldsValues where UDFPK = " + row["UDFPK"] + " and OfferID = " + OfferID;
                    lCommon.LRT_Execute();
                    continue;
                }




                if (Convert.ToInt32(row["DataType"]) != 0) //string data type is handled separately
                {
                    if (value == "")
                    {
                        if (Request.QueryString["udfVal-" + row["UDFPK"]] != null)
                        {
                            value = CleanString(Request.QueryString["udfVal-" + row["UDFPK"]], AllowSpecialCharacters);
                        }
                        else
                        {
                            value = "";
                        }
                    }
                    lCommon.QueryStr = "dbo.pt_UserDefinedFieldsValues_Update";   
                    lCommon.Open_LRTsp();
                    lCommon.LRTsp.Parameters.Add("@UDFPK", System.Data.SqlDbType.BigInt).Value = row["UDFPK"];
                    lCommon.LRTsp.Parameters.Add("@OfferID", System.Data.SqlDbType.BigInt).Value = OfferID;
                    lCommon.LRTsp.Parameters.Add("@DataType", System.Data.SqlDbType.Int).Value = row["DataType"];

                    lCommon.LRTsp.Parameters.Add("@Value", System.Data.SqlDbType.NVarChar, 150).Value = value;
                    lCommon.LRTsp.Parameters.Add("@Status", System.Data.SqlDbType.BigInt).Direction = System.Data.ParameterDirection.Output;
                    lCommon.LRTsp.ExecuteNonQuery();
                    Status = Convert.ToInt32(lCommon.LRTsp.Parameters["@Status"].Value);
                    lCommon.Close_LRTsp();
                    if (Status == 1)
                    {
                        UDFHistory += Copient.PhraseLib.Lookup("term.userdefinedfield", LanguageID) + " - " + lCommon.TruncateString(Convert.ToString(row["Description"]), 20) + ", ";
                    }
                    else if (Status == -3)
                    {
                        infoMessage = Copient.PhraseLib.Lookup("term.castingerror", LanguageID);
                    }
                }
                else
                {
                    lCommon.QueryStr = "dbo.pt_PopulateUDFStringValues";
                    lCommon.Open_LRTsp();
                    lCommon.LRTsp.Parameters.Add("@OfferID", System.Data.SqlDbType.BigInt).Value = OfferID;
                    lCommon.LRTsp.Parameters.Add("@UDFPK", System.Data.SqlDbType.BigInt).Value = row["UDFPK"];
                    lCommon.LRTsp.Parameters.Add("@Status", System.Data.SqlDbType.BigInt).Direction = System.Data.ParameterDirection.Output;
                    lCommon.LRTsp.ExecuteNonQuery();
                    Status = Convert.ToInt32(lCommon.LRTsp.Parameters["@Status"].Value);
                    lCommon.Close_LRTsp();
                    if (Status == 1)
                    {
                        UDFHistory += Copient.PhraseLib.Lookup("term.userdefinedfield", LanguageID) + " - " + lCommon.TruncateString(Convert.ToString(row["Description"]), 20) + ", ";
                    }
                }     
            }
            else//new saving functionality
            {
                 if(Convert.ToBoolean(row["deleted"]) == true)
                {
                    UDFHistory += Copient.PhraseLib.Lookup("term.deleted", LanguageID) + " " + Copient.PhraseLib.Lookup("term.userdefinedfield", LanguageID) + " - " + lCommon.TruncateString(Convert.ToString(row["Description"]), 20) + ", ";
                    lCommon.QueryStr = "delete from UserDefinedFieldsValues where UDFPK = " + row["UDFPK"] + " and OfferID = " + OfferID;
                    lCommon.LRT_Execute(); 
                    continue;
                }
                switch (Convert.ToString(row["PresentationStyleID"]))
                {
                    case "DropDownList":
                    case "ListBox":

                        bool didOnce = false;
                        if (Request.QueryString.Get("udfVal-" + row["UDFPK"]) != null)
                        {
                            value = Request.QueryString.Get("udfVal-" + row["UDFPK"]);//this may be a comma delimited list of multiple UDFVPKs
                        }
                        if (value == "")
                        {
                            continue;
                        }
                        //let's delete the old list and repopulate.
                        lCommon.Open_LogixRT();
                        
                        lCommon.QueryStr = "select count(*) as numToDelete from UserDefinedFieldsValues where OfferID=" + OfferID + " and UDFPK = " + Convert.ToString(row["UDFPK"]) + " and UDFVPK not in (" + value + ")";
                        System.Data.DataTable dtCheckDelete = lCommon.LRT_Select();
                        if (dtCheckDelete != null)
                        {
                            if (dtCheckDelete.Rows.Count > 0)
                            {
                                if (Convert.ToInt16(dtCheckDelete.Rows[0]["numToDelete"]) > 0)
                                {
                                    lCommon.QueryStr = "delete from UserDefinedFieldsValues where OfferID=" + OfferID + " and UDFPK = " + Convert.ToString(row["UDFPK"]) + " and UDFVPK not in (" + value + ")";
                                    lCommon.LRT_Execute();
                                    if (!didOnce)
                                    {
                                        UDFHistory += Copient.PhraseLib.Lookup("term.userdefinedfield", LanguageID) + " - " + lCommon.TruncateString(Convert.ToString(row["Description"]), 20) + ", ";
                                        didOnce = true;
                                    }
                                }
                            }
                        }   
                        lCommon.QueryStr = "select UDFVPK,value from UserDefinedField_ValidValues where UDFPK = " + Convert.ToString(row["UDFPK"]) + " and UDFVPK in (" + value + ") and UDFVPK NOT in (select UDFVPK from UserDefinedFieldsValues where OfferID=" + OfferID + " and UDFPK = " + Convert.ToString(row["UDFPK"])+")";

                        lCommon.Open_LogixRT();
                        udfvv = lCommon.LRT_Select();
                            
                        foreach(System.Data.DataRow dr in udfvv.Rows)
                        {
                            lCommon.QueryStr = "dbo.pt_UserDefinedFieldsValues_Insert";
			                lCommon.Open_LRTsp();
                            lCommon.LRTsp.Parameters.Add("@UDFPK", System.Data.SqlDbType.BigInt).Value = row["UDFPK"];
                            lCommon.LRTsp.Parameters.Add("@OfferID", System.Data.SqlDbType.Int).Value = OfferID;
                            lCommon.LRTsp.Parameters.Add("@UDFVPK", System.Data.SqlDbType.BigInt).Value = dr["UDFVPK"];
			                lCommon.LRTsp.Parameters.Add("@Status", System.Data.SqlDbType.Int).Direction = System.Data.ParameterDirection.Output;
                            lCommon.LRTsp.ExecuteNonQuery();
                                
                            Status = Convert.ToInt32(lCommon.LRTsp.Parameters["@Status"].Value);
                            if (Status == 1)
                            {
                                if (!didOnce)
                                {
                                    UDFHistory += Copient.PhraseLib.Lookup("term.userdefinedfield", LanguageID) + " - " + lCommon.TruncateString(Convert.ToString(row["Description"]), 20) + ", ";
                                    didOnce = true;
                                }
                            }
                            else if (Status == -3)
                            {
                                infoMessage = Copient.PhraseLib.Lookup("term.castingerror", LanguageID);
                            }
                        }
                        lCommon.Close_LRTsp();
                        break;
                    case "HorizontalRadioButtons":
                    case "VerticalRadioButtons":
                        if (Request.QueryString.Get("udfvalidValue-" + Convert.ToString((row["UDFPK"]))) != null)
                        {
                            value = CleanString(Request.QueryString["udfvalidValue-" + Convert.ToString((row["UDFPK"]))], AllowSpecialCharacters);
                        }
                        if (value == "")
                        {
                            continue;
                        }
                        lCommon.QueryStr = "select value from UserDefinedField_ValidValues where UDFPK = " + row["UDFPK"] + " and UDFVPK=" + value;
                        lCommon.Open_LRTsp();
                        udfvv = lCommon.LRT_Select();

                        if (udfvv.Rows.Count == 1)
                        {
                            lCommon.QueryStr = "dbo.pt_UserDefinedFieldsValues_Update";
                            lCommon.Open_LRTsp();
                            lCommon.LRTsp.Parameters.Add("@UDFPK", System.Data.SqlDbType.BigInt).Value = row["UDFPK"];
                            lCommon.LRTsp.Parameters.Add("@OfferID", System.Data.SqlDbType.BigInt).Value = OfferID;
                            lCommon.LRTsp.Parameters.Add("@DataType", System.Data.SqlDbType.Int).Value = row["DataType"];
                            if (Convert.ToInt32(row["DataType"]) == 3)
                            {
                                lCommon.LRTsp.Parameters.Add("@Value", System.Data.SqlDbType.NVarChar, 1000).Value = Convert.ToBoolean(udfvv.Rows[0]["value"]);
                            }
                            else
                            {
                                lCommon.LRTsp.Parameters.Add("@Value", System.Data.SqlDbType.NVarChar, 1000).Value = udfvv.Rows[0]["value"];
                            }
                            lCommon.LRTsp.Parameters.Add("@UDFVPK", System.Data.SqlDbType.BigInt).Value = value;
                            lCommon.LRTsp.Parameters.Add("@Status", System.Data.SqlDbType.BigInt).Direction = System.Data.ParameterDirection.Output;
                            lCommon.LRTsp.ExecuteNonQuery();
                            Status = Convert.ToInt32(lCommon.LRTsp.Parameters["@Status"].Value);
                            if (Status == 1)
                            {
                                UDFHistory += Copient.PhraseLib.Lookup("term.userdefinedfield", LanguageID) + " - " + lCommon.TruncateString(Convert.ToString(row["Description"]), 20) + ", ";
                            }
                            else if (Status == -3)
                            {
                                infoMessage = Copient.PhraseLib.Lookup("term.castingerror", LanguageID);
                            }
                        }
                        lCommon.Close_LRTsp();
                        break;
                    case "CheckBox":
                        if (Convert.ToInt32(row["DataType"]) == 3)//'checkboxes dont return a value when unchecked. 
                        {
                            if (Request.QueryString.Get("udfVal-" + Convert.ToString(row["UDFPK"])) != null)
                            {
                                value = "1";
                            }
                            else
                            {
                                value = "0";
                            }
                            lCommon.QueryStr = "dbo.pt_UserDefinedFieldsValues_Update";
                            lCommon.Open_LRTsp();
                            lCommon.LRTsp.Parameters.Add("@UDFPK", System.Data.SqlDbType.BigInt).Value = row["UDFPK"];
                            lCommon.LRTsp.Parameters.Add("@OfferID", System.Data.SqlDbType.BigInt).Value = OfferID;
                            lCommon.LRTsp.Parameters.Add("@DataType", System.Data.SqlDbType.Int).Value = row["DataType"];
                            lCommon.LRTsp.Parameters.Add("@Value", System.Data.SqlDbType.NVarChar, 1000).Value = value;
                            lCommon.LRTsp.Parameters.Add("@Status", System.Data.SqlDbType.BigInt).Direction = System.Data.ParameterDirection.Output;
                            lCommon.LRTsp.ExecuteNonQuery();
                            Status = Convert.ToInt32(lCommon.LRTsp.Parameters["@Status"].Value);
                            lCommon.Close_LRTsp();
                            if (Status == 1)
                            {
                                UDFHistory += Copient.PhraseLib.Lookup("term.userdefinedfield", LanguageID) + " - " + lCommon.TruncateString(Convert.ToString(row["Description"]), 20) + ", ";
                            }
                            else if (Status == -3)
                            {
                                infoMessage = Copient.PhraseLib.Lookup("term.castingerror", LanguageID);
                            }
                        }
                        break;
                    case "TextBox"://assumes a single row in userdefinedfieldsvalues table
                        if (Convert.ToInt32(row["DataType"]) == 0)
                        {
                            lCommon.QueryStr = "dbo.pt_PopulateUDFStringValues";
                            lCommon.Open_LRTsp();
                            lCommon.LRTsp.Parameters.Add("@OfferID", System.Data.SqlDbType.BigInt).Value = OfferID;
                            lCommon.LRTsp.Parameters.Add("@UDFPK", System.Data.SqlDbType.BigInt).Value = row["UDFPK"];
                            lCommon.LRTsp.Parameters.Add("@Status", System.Data.SqlDbType.BigInt).Direction = System.Data.ParameterDirection.Output;
                            lCommon.LRTsp.ExecuteNonQuery();                                
                                
                            Status = Convert.ToInt32(lCommon.LRTsp.Parameters["@Status"].Value);
                            lCommon.Close_LRTsp();
                            if (Status == 1)
                            {
                                UDFHistory += Copient.PhraseLib.Lookup("term.userdefinedfield", LanguageID) + " - " + lCommon.TruncateString(Convert.ToString(row["Description"]), 20) + ", ";
                            }
                        }
                        else //' date and integer and numeric range
                        {
                            if (Request.QueryString.Get("udfVal-" + Convert.ToString(row["UDFPK"])) != null)
                            {
                            value = CleanString(Request.QueryString["udfVal-" + Convert.ToString(row["UDFPK"])], AllowSpecialCharacters);
                            }
                            if (value == "")
                            {
                                //continue;
                            }

                            bool validateData = true;
                            string ValidateMessage = string.Empty;
                            if (Convert.ToInt32(row["DataType"]) == 5)
                            {
								if (value == String.Empty){
									validateData=false;
									infoMessage="Please enter a valid number for range";
									break;
									}
                                lCommon.QueryStr = "select value from userdefinedfield_validvalues where udfpk = " + Convert.ToString(row["UDFPK"]);
                                lCommon.Open_LRTsp();
                                System.Data.DataTable dtvnr = lCommon.LRT_Select();
                                validateData = false;
                                foreach(System.Data.DataRow drvnr in dtvnr.Rows)
                                {
                                    String rangeValue = Convert.ToString(drvnr["value"]);
                                    String[] rangeValueparts = rangeValue.Split(':');
                                    if (rangeValueparts.Length == 2)
                                    {
                                        rangeValueparts[0] = rangeValueparts[0].Replace("{", "");
                                        rangeValueparts[1] = rangeValueparts[1].Replace("}", "");
                                            
                                        Int64 minInt = Convert.ToInt64(rangeValueparts[0]);
                                        Int64 maxInt = Convert.ToInt64(rangeValueparts[1]);
                                        if (Convert.ToInt64(value) >= minInt && Convert.ToInt64(value) <= maxInt)
                                        {
                                            validateData = true;
                                            break;
                                        }
                                    }
                                    else
                                    {
                                           
                                        if (Convert.ToInt64(value) == Convert.ToInt64(rangeValueparts[0]))
                                        {
                                            validateData = true;
                                            break;
                                        }
                                    }                                        
                                }
                                if (!validateData)
                                {
                                    ValidateMessage = "Invalid number in range.";
                                }
                            }

                            if (validateData && Convert.ToInt32(row["DataType"]) == 1)
                            {
                                int result = 0;
                                if (!string.IsNullOrEmpty(value) && !int.TryParse(value, out result)) 
                                {
                                    validateData = false;
                                    ValidateMessage = "User Defined Field value should not exceed 2147483647 for Integer Type.";
                                }
                            }

                            if (validateData == true)
                            {
                                lCommon.QueryStr = "dbo.pt_UserDefinedFieldsValues_Update";
                                lCommon.Open_LRTsp();
                                lCommon.LRTsp.Parameters.Add("@UDFPK", System.Data.SqlDbType.BigInt).Value = row["UDFPK"];
                                lCommon.LRTsp.Parameters.Add("@OfferID", System.Data.SqlDbType.BigInt).Value = OfferID;
                                lCommon.LRTsp.Parameters.Add("@DataType", System.Data.SqlDbType.Int).Value = row["DataType"];
                                lCommon.LRTsp.Parameters.Add("@Value", System.Data.SqlDbType.NVarChar, 1000).Value = value;
                                lCommon.LRTsp.Parameters.Add("@Status", System.Data.SqlDbType.BigInt).Direction = System.Data.ParameterDirection.Output;
                                lCommon.LRTsp.ExecuteNonQuery();

                                Status = Convert.ToInt32(lCommon.LRTsp.Parameters["@Status"].Value);
                                if (Status == 1)
                                {
                                    UDFHistory += Copient.PhraseLib.Lookup("term.userdefinedfield", LanguageID) + " - " + lCommon.TruncateString(Convert.ToString(row["Description"]), 20) + ", ";
                                }
                                else if (Status == -3)
                                {
                                    infoMessage = Copient.PhraseLib.Lookup("term.castingerror", LanguageID);
                                }
                            }
                            else
                            {                                   
                                infoMessage = ValidateMessage;
                            }
                        }
                        break;
                    default:
                        break;
                }//switch                
            }//else PresentationStyleID != DBNull
        }//foreach
    }
 %>