
// version:7.3.1.138972.Official Build (SUSDAY10202) %>

//' *****************************************************************************
//  ' * FILENAME: UDFListClass.cs
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
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using CMS.AMS;
using CMS.Contract;
using CMS.AMS.Contract;
using CMS.AMS.Models;



public partial class UDFUserControlBase : System.Web.UI.UserControl
{

    private String _infoMessage;
    public String infoMessage
    {
        get
        {
            return _infoMessage;
        }
        set
        {
            _infoMessage = value;
        }
    }

    private String _UDFHistory;
    public String UDFHistory
    {
        get
        {
            return _UDFHistory;
        }
        set
        {
            _UDFHistory = value;
        }
    }   

    //needs to be set by including page
    private bool _IsTemplate = false;
    public bool IsTemplate
    {
        get { return _IsTemplate; }
        set { _IsTemplate = value; }
    }


    private bool _bUseTemplateLocks = false;
    public bool bUseTemplateLocks
    {
        get
        {
            return _bUseTemplateLocks;
        }
        set
        {
            _bUseTemplateLocks = value;
        }
    }

    private bool _Disallow_UserDefinedFields = false;
    public bool Disallow_UserDefinedFields
    {
        get
        {
            return _Disallow_UserDefinedFields;
        }
        set
        {
            _Disallow_UserDefinedFields = value;
        }
    }


    public int LanguageID;


    public String disableUDF = "";
    public Copient.CommonInc lCommon;
    public Copient.LogixInc lLogix;
    public bool AllowEditing = true;
    public bool EditOfferPastLockoutdate = false;
    public String OfferID = "";
    public long AdminUserID; //need to see how the containing page gets this
    public bool Debug = false;

    public String AdminName;
    public String AllowSpecialCharacters = String.Empty;
    protected void Page_Init(object sender, EventArgs e)
    {
        lCommon = new Copient.CommonInc();
        lLogix = new Copient.LogixInc();
        object common = lCommon;

        lCommon.AppName = "UDFControl";
        lCommon.Open_LogixRT();
        Object o = lCommon;
        Object o2 = lLogix;
        AdminUserID = Verify_AdminUser(ref o, ref o2);

        AllowEditing = lLogix.UserRoles.EditUserDefinedFields;
        EditOfferPastLockoutdate = lLogix.UserRoles.EditOfferPastLockoutPeriod;

        if (Request.QueryString["OfferID"] != null)
        {
            OfferID = Request.QueryString["OfferID"];
        }
        else
        {
            OfferID = Request.Form["OfferID"];
        }

        AllowSpecialCharacters = lCommon.Fetch_SystemOption(171);
    }


    public string GetCgiValue(String VarName)
    {
        String TempVal;
        TempVal = String.Empty;

        if (Request.QueryString[VarName] != null)
        {
            TempVal = Request.QueryString[VarName];
        }
        else if (Request.Form[VarName] != null)
        {
            TempVal = Request.Form[VarName];
        }
        return TempVal;
    }

    public long Verify_AdminUser(ref object Common, ref object MyLogix)
    {

        String Authtoken = "";
        String MyURI;
        String TransferKey = "";
        Boolean Debug = false;

        if (AdminUserID != 0)
        {
            if (Debug) { lCommon.Write_Log("auth.txt", "AppName=" + lCommon.AppName + " - Verify_AdminUser was called, but we already know the AdminUserID=" + AdminUserID, true); }
            //'we already know who the AdminUser is ... we shouldn't be looking him up more than once

            Object o = lCommon;
            lLogix.Load_Roles(ref o, AdminUserID);
            return AdminUserID;
        }

        
        //'1st, check the transferkey and see if the user is being transferred into AMS from another product (PrefMan)
        if (GetCgiValue("transferkey") != String.Empty)
        {
            
            if (Debug) { lCommon.Write_Log("auth.txt", "AppName=" + lCommon.AppName + " - Checking the TransferKey (" + GetCgiValue("transferkey") + ")  AdminUserID=" + AdminUserID, true); }
            TransferKey = GetCgiValue("transferkey");
            AdminUserID = lLogix.Auth_TransferKey_Verify(ref lCommon, TransferKey, ref AdminName, ref LanguageID, ref Authtoken);
            if (Debug) { lCommon.Write_Log("auth.txt", "AppName=" + lCommon.AppName + " - After TransferKey_Verify AdminUserID=" + AdminUserID, true); }
            if (AdminUserID != 0)
            {
                Response.Cookies["AuthToken"].Value = Authtoken;
                Object o = lCommon;
                lLogix.Load_Roles(ref o, AdminUserID);
                return AdminUserID;
            }
        }
        else
        {
            
        }
        Authtoken = "";
        if (Request.Cookies["AuthToken"] != null)
        {
            Authtoken = Request.Cookies["AuthToken"].Value;
        }

        if (Debug) { lCommon.Write_Log("auth.txt", "AppName=" + lCommon.AppName + " - AuthToken='" + Authtoken + "'   Transferkey='" + GetCgiValue("transferkey") + "'", true); }
        AdminUserID = 0;
        AdminUserID = lLogix.Auth_Token_Verify(ref lCommon, Authtoken, ref AdminName, ref LanguageID);
        if (Debug) { lCommon.Write_Log("auth.txt", "AppName=" + lCommon.AppName + " - After checking AuthToken, AdminUserID=" + AdminUserID, true); }


        if (AdminUserID == 0)
        {
            MyURI = System.Web.HttpUtility.UrlEncode(Request.Url.AbsoluteUri);
            /*
            Send("<!DOCTYPE html ")
            Send("     PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN""")
            Send("     ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">")
            Send("<html xmlns=""http://www.w3.org/1999/xhtml"">")
            Send("<head>")
            Send("<meta http-equiv=""refresh"" content=""0; url=/logix/login.aspx?mode=invalid&amp;bounceback=" & MyURI & """ />")
            Send("<title>Logix</title>")
            Send("</head>")
            Send("<body bgcolor=""#ffffff"">")
            Send("<!-- Bouncing -->")
            Send("</body>")
            Send("</html>")
            Response.End()*/
        }
        else
        {
            Object o = lCommon;
            lLogix.Load_Roles(ref o, AdminUserID);
        }
        return AdminUserID;
    }

     public String CleanString(String InString, String AdditionalValidCharacters = "")
     {
        String tmpString = String.Empty;
        int z;
        String allowedCharacters= "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789.#$*-_&%@!?/:;+ " + AdditionalValidCharacters;
        if(InString != String.Empty)
        {
            for(z = 0; z< (InString.Length);z++)
            {
                
                if(allowedCharacters.Contains(InString[z]))
                {
                    tmpString = tmpString + InString[z];
                }
            }
        }

        return tmpString;
    }
}

