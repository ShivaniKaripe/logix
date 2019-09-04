using System;
using System.Data;
using System.Xml;
using System.IO;
using System.Web.Services;
using System.Linq;
using CMS;
using CMS.Contract;
using CMS.AMS;
using CMS.AMS.Contract;
using CMS.DB;
using CMS.AMS.Models;
using System.Data.SqlClient;
using System.Collections.Generic;



/// <summary>
/// Summary description for LogixDataDefinition
/// </summary>

[WebService(Namespace = "http://ncr.cms.ams.com/")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
// To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
// [System.Web.Script.Services.ScriptService]
public class LogixDataDefinition : System.Web.Services.WebService
{
    #region private Data objects

    #region PointsPrograms

    [Serializable]
    public class PointProgram
    {
        public long ProgramID { get; set; }
        public string ProgramName { get; set; }
        public string Description { get; set; }
        public long? PromoVarID { get; set; }
        public DateTime? LastLoaded { get; set; }
        public string LastLoadMsg { get; set; }
        public bool ExternalProgram { get; set; }
        public string ExternalID { get; set; }
        public string ExtHostTypeDesc { get; set; }
        public string ExtHostProgramID { get; set; }
        public bool AutoDelete { get; set; }
        public string AdjustmentUPC { get; set; }
        public bool CAMProgram { get; set; }
        public Int32? EngineSubTypePKID { get; set; }
        public Int32? CategoryID { get; set; }
        public Int32 ExtHostTypeID { get; set; }
        public bool DecimalValues { get; set; }
        public Int32 ReturnHandlingTypeID { get; set; }
        public Int32 DisallowRedeemInEarnTrans { get; set; }
        public Int32 AllowNegativeBal { get; set; }
        public bool ScorecardBold { get; set; }
        public string ScorecardDesc { get; set; }
        public Int32? BuckOfferID { get; set; }
        public Int32? BuckTierNumber { get; set; }
        public bool? AllowAnyCustomer { get; set; } // UE engine specific column
        public bool? VisibleToCustomers { get; set; }

        public PointsProgramMigrations PointsProgramMigrationDetails { get; set; }
        public PointsProgramDeletions PointsProgramDeletionsDetails { get; set; }
        public List<ProgramTranslations> PointsProgramTransaltionDetails { get; set; }

    }
    [Serializable]
    public class PointsProgramMigrations
    {
        public long MigrationProgramID { get; set; }
        public DateTime? MigrationDate { get; set; }
        public DateTime? LastMigrationDate { get; set; }
        public DateTime? LastEmailDate { get; set; }
        public bool Deleted { get; set; }

    }
    [Serializable]
    public class PointsProgramDeletions
    {
        public DateTime DeletionDate { get; set; }
        public DateTime? LastDeletionDate { get; set; }
        public DateTime? LastEmailDate { get; set; }
        public bool Deleted { get; set; }
    }
    [Serializable]
    public class ProgramTranslations
    {
        public long PKID { get; set; }
        public long LanguageID { get; set; }
        public string ScorecardDesc { get; set; }
    }
    #endregion

    #region Stored value programs
    [Serializable]
    public class StoredValuePrograms
    {

        public long SVProgramID { get; set; }
        public string Name { get; set; }
        public decimal? Value { get; set; }
        public DateTime? CreatedDate { get; set; }
        public bool Deleted { get; set; }
        public string Description { get; set; }
        public bool OneUnitPerRec { get; set; }
        public bool SVExpireType { get; set; }
        public int? SVExpirePeriodType { get; set; }
        public int? ExpirePeriod { get; set; }
        public string ExpireTOD { get; set; }
        public DateTime? ExpireDate { get; set; }
        public int SVTypeID { get; set; }
        public int UnitOfMeasureLimit { get; set; }
        public bool AllowReissue { get; set; }
        public int CMOAStatusFlag { get; set; }
        public int CPEStatusFlag { get; set; }
        public bool? Visible { get; set; }
        public bool AutoDelete { get; set; }
        public string ExtProgramID { get; set; }
        public int? ScorecardID { get; set; }
        public string ScorecardDesc { get; set; }
        public bool ScorecardBold { get; set; }
        public string AdjustmentUPC { get; set; }
        public int? EngineSubTypePKID { get; set; }
        public int ReturnHandlingTypeID { get; set; }
        public int DisallowRedeemInEarnTrans { get; set; }
        public int AllowNegativeBal { get; set; }
        public int? RedemptionRestrictionID { get; set; }
        public int? MemberRedemptionId { get; set; }
        public bool? FuelPartner { get; set; }
        public bool? AutoRedeem { get; set; }
        public bool? AllowAdjustments { get; set; }
        public bool? ExpireCentralServerTZ { get; set; }
        public bool? AllowAnyCustomer { get; set; } // UE engine specific column
        public bool? VisibleToCustomers { get; set; }

        public List<ProgramTranslations> SVProgramTransaltionDetails { get; set; }

    }
    #endregion

    #endregion

    #region Private Variables

    private CMS.AMS.Common m_common = null;
    private ConnectorInc m_connectorInc = null;
    private CMS.CryptLib m_cryptlib = null;
    private ILogger m_logger = null;
    private IErrorHandler m_errHandler = null;
    private CMS.AMS.AuthLib m_authInc = null;
    private IPhraseLib m_phraseLib = null;
    private IAdminUserData m_adminUser = null;
    private IDBAccess m_dbaccess = null;
    private Copient.CommonInc m_commonInc = null;
    private SystemSettings m_systemSettings = null;
    ICacheData SystemCacheData = null;
    #endregion Private Variables

    #region Private methods - general

    private void Startup()
    {
        CurrentRequest.Resolver.AppName = "LogixDataDefinition";
        m_common = CurrentRequest.Resolver.Resolve<CMS.AMS.Common>();
        if (m_common.LRT_Connection_State() == ConnectionState.Closed) { m_common.Open_LogixRT(); }
        if (m_common.LXS_Connection_State() == ConnectionState.Closed) { m_common.Open_LogixXS(); }
        CurrentRequest.Resolver.RegisterInstance<CommonBase>(m_common);

        m_common.Set_AppInfo();

        m_connectorInc = new ConnectorInc(m_common);
        m_logger = CurrentRequest.Resolver.Resolve<ILogger>();
        m_systemSettings = CurrentRequest.Resolver.Resolve<SystemSettings>();
        m_phraseLib = CurrentRequest.Resolver.Resolve<IPhraseLib>();
        m_errHandler = CurrentRequest.Resolver.Resolve<IErrorHandler>();
        m_authInc = new CMS.AMS.AuthLib(m_common);
        m_adminUser = CurrentRequest.Resolver.Resolve<IAdminUserData>();
        m_dbaccess = CurrentRequest.Resolver.Resolve<IDBAccess>();
        m_commonInc = new Copient.CommonInc();
        m_cryptlib = new CryptLib();
    }

    private void Shutdown()
    {
        if (m_common.LXS_Connection_State() != ConnectionState.Closed) { m_common.Close_LogixXS(); }
        if (m_common.LRT_Connection_State() != ConnectionState.Closed) { m_common.Close_LogixRT(); }

        m_connectorInc = null;
        m_common = null;
        m_logger = null;
        m_errHandler = null;
        m_phraseLib = null;
        m_adminUser = null;
        m_dbaccess = null;
        m_commonInc = null;
        m_systemSettings = null;
    }

    /// <summary>
    /// Processes exception catch for other methods
    /// </summary>
    /// <param name="ex">Exception that occurred</param>
    /// <param name="methodName">Name of the method where the exception occurred</param>
    /// <param name="xmlWriter">XmlWriter from the method where the exception occurred</param>
    /// <param name="stringWriter">StringWriter from the method where the exception occurred</param>
    private void ProcessException(Exception ex, string methodName, ref XmlWriter xmlWriter, ref StringWriter stringWriter)
    {
        m_connectorInc.Init_ResponseXML(methodName, ref stringWriter, ref xmlWriter);
        if (ex is CMSException)
        {
            CMSException.StatusCodes statusCode = CMSException.StatusCodes.GENERAL_ERROR;
            if (ex.Data != null)
            {
                statusCode = (CMSException.StatusCodes)ex.Data["StatusCode"];
            }
            m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, statusCode, ex.Message, false);
        }
        else
        {
            m_common.Error_Processor(ex.ToString());
            m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, CMS.CMSException.StatusCodes.APPLICATION_EXCEPTION, m_phraseLib.Lookup("term.errorprocessingseelog"), false);
            m_common.Write_Log("An error occurred while processing - please see the error log!");
        }
    }

    private void CloseResponseXml(ref XmlWriter xmlWriter, ref StringWriter strWriter, ref XmlDocument xmlResponse)
    {
        m_connectorInc.Close_ResponseXML(ref xmlWriter);
        xmlWriter.Flush();
        xmlWriter.Close();
        xmlResponse.LoadXml(strWriter.ToString());
    }

    private AMSResult<T> ReturnAMSResult<T>(T result, AMSResultType resultType, string message)
    {
        AMSResult<T> returnAMSResult = new AMSResult<T>(result);
        returnAMSResult.ResultType = resultType;
        returnAMSResult.MessageString = message;
        return returnAMSResult;
    }

    #endregion

    #region Private methods -PointPrograms
    private AMSResult<List<PointProgram>> GetPointsProgramsList(string name)
    {

        bool bPointsMigrationEnabled = false;
        bool bPointsDeletionEnabled = false;
        bool bMultiLanguageEnabled = false;
        PointProgram pp = null;
        List<PointProgram> ppLst = null;
        SQLParametersList lstParams = null;
        try
        {

            if ((m_common.Fetch_SystemOption(164) == "1")) { bPointsMigrationEnabled = true; }
            if ((m_common.Fetch_SystemOption(205) == "1")) { bPointsDeletionEnabled = true; }
            if ((m_common.Fetch_SystemOption(124) == "1")) { bMultiLanguageEnabled = true; }

            m_logger.WriteInfo("Getting all active Point Program definitions");

            AMSResult<PointProgram> result = new AMSResult<PointProgram>();

            //Coverity CID - 93967
            string query = "SELECT * FROM PointsPrograms WITH (NoLock) WHERE Deleted=0 ";
            DataTable dt = new DataTable();
            if (!String.IsNullOrEmpty(name))
            {
                lstParams = new SQLParametersList();
                lstParams.Add("@name", SqlDbType.VarChar).Value = "%" + name + "%";
                query = query + " AND ProgramName like @name";
                dt = m_dbaccess.ExecuteQuery(DataBases.LogixRT, CommandType.Text, query, lstParams);
            }
            else
            {
                dt = m_dbaccess.ExecuteQuery(DataBases.LogixRT, CommandType.Text, query, null);
            }
            if (dt.Rows.Count > 0)
            {
                ppLst = new List<PointProgram>();
                foreach (DataRow rw in dt.Rows)
                {
                    #region retreiving data from PointsProgram table
                    pp = new PointProgram();

                    pp.ProgramID = Convert.ToInt32(rw["ProgramID"]);
                    pp.ProgramName = (rw["ProgramName"] != DBNull.Value) ? Convert.ToString(rw["ProgramName"]) : null;
                    pp.Description = (rw["Description"] != DBNull.Value) ? Convert.ToString(rw["Description"]) : null;
                    pp.PromoVarID = (rw["PromoVarID"] != DBNull.Value) ? Convert.ToInt64(rw["PromoVarID"]) : 0;
                    if (rw["LastLoaded"] != DBNull.Value) { pp.LastLoaded = Convert.ToDateTime(rw["LastLoaded"]); }
                    pp.LastLoadMsg = (rw["LastLoadMsg"] != DBNull.Value) ? Convert.ToString(rw["LastLoadMsg"]) : null;
                    if (rw["LastLoaded"] != DBNull.Value) { pp.LastLoaded = Convert.ToDateTime(rw["LastLoaded"]); }
                    pp.AutoDelete = (rw["AutoDelete"] != DBNull.Value) ? Convert.ToBoolean(rw["AutoDelete"]) : false;
                    pp.ScorecardBold = (rw["ScorecardBold"] != DBNull.Value) ? Convert.ToBoolean(rw["ScorecardBold"]) : false;
                    pp.ScorecardDesc = (rw["ScorecardDesc"] != DBNull.Value) ? Convert.ToString(rw["ScorecardDesc"]) : null;
                    pp.AdjustmentUPC = (rw["AdjustmentUPC"] != DBNull.Value) ? Convert.ToString(rw["AdjustmentUPC"]) : null;
                    pp.CAMProgram = (rw["CAMProgram"] != DBNull.Value) ? Convert.ToBoolean(rw["CAMProgram"]) : false;
                    if (rw["EngineSubTypePKID"] != DBNull.Value) { pp.EngineSubTypePKID = Convert.ToInt32(rw["EngineSubTypePKID"]); }
                    if (rw["CategoryID"] != DBNull.Value) { pp.CategoryID = Convert.ToInt32(rw["CategoryID"]); }
                    pp.DecimalValues = (rw["DecimalValues"] != DBNull.Value) ? Convert.ToBoolean(rw["DecimalValues"]) : false;
                    pp.ExternalProgram = (rw["ExternalProgram"] != DBNull.Value) ? Convert.ToBoolean(rw["ExternalProgram"]) : false;
                    if (pp.ExternalProgram)
                    {
                        if (rw["ExtHostTypeID"] == DBNull.Value)
                        {
                            string query1 = "select ExternalID from PromoVariables with (NoLock) where PromoVarID =@PromoVarID";

                            lstParams = new SQLParametersList();
                            lstParams.Add("@PromoVarID", SqlDbType.BigInt).Value = pp.PromoVarID;
                            DataTable dt1 = m_dbaccess.ExecuteQuery(DataBases.LogixXS, CommandType.Text, query1, lstParams);
                            if (dt1.Rows.Count > 0)
                            {
                                pp.ExternalID = (dt1.Rows[0]["ExternalID"] != DBNull.Value) ? m_cryptlib.SQL_StringDecrypt(Convert.ToString(dt1.Rows[0]["ExternalID"]).ToString()) : null;
                            }
                        }
                        else
                        {
                            pp.ExternalID = (rw["ExtHostProgramID"] != DBNull.Value) ? Convert.ToString(rw["ExtHostProgramID"]) : "0";
                            pp.ExtHostTypeID = (rw["ExtHostTypeID"] != DBNull.Value) ? Convert.ToInt32(rw["ExtHostTypeID"]) : 0;
                        }

                    }
                    else
                    {
                        pp.ExternalID = (rw["ExtHostProgramID"] != DBNull.Value) ? Convert.ToString(rw["ExtHostProgramID"]) : "0";
                    }
                    pp.ReturnHandlingTypeID = (rw["ReturnHandlingTypeID"] != DBNull.Value) ? Convert.ToInt32(rw["ReturnHandlingTypeID"]) : 1;
                    pp.DisallowRedeemInEarnTrans = (rw["DisallowRedeemInEarnTrans"] != DBNull.Value) ? Convert.ToInt32(rw["DisallowRedeemInEarnTrans"]) : 0;
                    pp.AllowNegativeBal = (rw["AllowNegativeBal"] != DBNull.Value) ? Convert.ToInt32(rw["AllowNegativeBal"]) : 0;
                    pp.VisibleToCustomers = (rw["VisibleToCustomers"] != DBNull.Value) ? Convert.ToBoolean(rw["VisibleToCustomers"]) : false;
                    if (pp.ExtHostTypeID > 0)
                    {
                        string query1 = "select Description from ExtHostTypes with (NoLock) where ExtHostTypeID=@ExtHostTypeID";
                        lstParams = new SQLParametersList();
                        lstParams.Add("@ExtHostTypeID", SqlDbType.BigInt).Value = pp.ExtHostTypeID;
                        DataTable dt1 = m_dbaccess.ExecuteQuery(DataBases.LogixRT, CommandType.Text, query1, lstParams);

                        if (dt1.Rows.Count > 0)
                            if (dt1.Rows[0]["Description"] != DBNull.Value) { pp.ExtHostTypeDesc = Convert.ToString(dt1.Rows[0]["Description"]); }
                            else
                                pp.ExtHostTypeDesc = null;
                    }
                    if (m_systemSettings.IsEngineInstalled(Engines.UE) == true)
                    {
                        string query1 = "SELECT PKID, AllowAnyCustomer from PointsProgramsPromoEngineSettings " +
                                       "WHERE ProgramID = @ProgramID AND EngineID = @EngineID";
                        lstParams = new SQLParametersList();
                        lstParams.Add("@ProgramID", SqlDbType.BigInt).Value = pp.ProgramID;
                        lstParams.Add("@EngineID", SqlDbType.Int).Value = 9;
                        DataTable dt1 = m_dbaccess.ExecuteQuery(DataBases.LogixRT, CommandType.Text, query1, lstParams);

                        if (dt1.Rows.Count > 0)
                        {
                            if (dt1.Rows[0]["AllowAnyCustomer"] != DBNull.Value) { pp.AllowAnyCustomer = Convert.ToBoolean(dt1.Rows[0]["AllowAnyCustomer"]); }
                        }
                    }

                    #endregion

                    #region retreiving data from PointsProgramMigrations table if migration is enabled

                    if (bPointsMigrationEnabled)
                    {
                        string query1 = "select MigrationProgramID, MigrationDate, LastMigrationDate, LastEmailDate from PointsProgramMigrations with (NoLock) " +
                                        "where ProgramID=@ProgramID and Deleted=0";
                        lstParams = new SQLParametersList();
                        lstParams.Add("@ProgramID", SqlDbType.BigInt).Value = pp.ProgramID;
                        DataTable dt1 = m_dbaccess.ExecuteQuery(DataBases.LogixRT, CommandType.Text, query1, lstParams);
                        PointsProgramMigrations pm = null;
                        if (dt1.Rows.Count > 0)
                        {
                            pm = new PointsProgramMigrations();
                            pm.MigrationProgramID = dt1.Rows[0]["MigrationProgramID"] != DBNull.Value ? Convert.ToInt32(dt1.Rows[0]["MigrationProgramID"]) : 0;
                            pm.MigrationDate = (dt1.Rows[0]["MigrationDate"] == DBNull.Value) ? (DateTime?)null : (DateTime)dt1.Rows[0]["MigrationDate"];
                            pm.LastMigrationDate = (dt1.Rows[0]["LastMigrationDate"] == DBNull.Value) ? (DateTime?)null : (DateTime)dt1.Rows[0]["LastMigrationDate"];
                            pm.LastEmailDate = (dt1.Rows[0]["LastEmailDate"] == DBNull.Value) ? (DateTime?)null : (DateTime)dt1.Rows[0]["LastEmailDate"];

                            if (pm != null)
                                pp.PointsProgramMigrationDetails = pm;
                        }

                    }
                    #endregion

                    #region retreiving data from PointsProgramDeletions table if deletion is enabled
                    if (bPointsDeletionEnabled)
                    {
                        string query1 = "select DeletionDate, LastDeletionDate, LastEmailDate from PointsProgramDeletions with (NoLock) " +
                                       "where ProgramID=@ProgramID and Deleted=0";
                        lstParams = new SQLParametersList();
                        lstParams.Add("@ProgramID", SqlDbType.BigInt).Value = pp.ProgramID;
                        DataTable dt1 = m_dbaccess.ExecuteQuery(DataBases.LogixRT, CommandType.Text, query1, lstParams);
                        PointsProgramDeletions pd = null;
                        if (dt1.Rows.Count > 0)
                        {
                            pd = new PointsProgramDeletions();
                            if (dt1.Rows[0]["DeletionDate"] != DBNull.Value) pd.DeletionDate = Convert.ToDateTime(dt1.Rows[0]["DeletionDate"]);
                            pd.LastDeletionDate = (dt1.Rows[0]["LastDeletionDate"] == DBNull.Value) ? (DateTime?)null : (DateTime)dt1.Rows[0]["LastDeletionDate"];
                            pd.LastEmailDate = (dt1.Rows[0]["LastEmailDate"] == DBNull.Value) ? (DateTime?)null : (DateTime)dt1.Rows[0]["LastEmailDate"];

                            if (pd != null)
                                pp.PointsProgramDeletionsDetails = pd;
                        }

                    }
                    #endregion

                    #region retrieving data PointsProgramTranslations if multi languauge is enabled
                    if (bMultiLanguageEnabled)
                    {
                        string query1 = "select PKID,LanguageID,ScorecardDesc from PointsProgramTranslations with (NoLock) " +
                                       "where ProgramID=@ProgramID ";
                        lstParams = new SQLParametersList();
                        lstParams.Add("@ProgramID", SqlDbType.BigInt).Value = pp.ProgramID;
                        DataTable dt1 = m_dbaccess.ExecuteQuery(DataBases.LogixRT, CommandType.Text, query1, lstParams);
                        List<ProgramTranslations> ptLst = new List<ProgramTranslations>();
                        ProgramTranslations pt = null;
                        if (dt1.Rows.Count > 0)
                        {

                            foreach (DataRow dr in dt1.Rows)
                            {
                                pt = new ProgramTranslations();
                                if (dr["PKID"] != DBNull.Value) pt.PKID = Convert.ToInt32(dr["PKID"]);
                                if (dr["LanguageID"] != DBNull.Value) pt.LanguageID = Convert.ToInt32(dr["LanguageID"]);
                                pt.ScorecardDesc = (dr["ScorecardDesc"] == DBNull.Value) ? "" : Convert.ToString(dr["ScorecardDesc"]);
                                ptLst.Add(pt);
                            }
                        }
                        if (ptLst.Count > 0)
                            pp.PointsProgramTransaltionDetails = ptLst;
                    }

                    #endregion

                    if (pp != null)
                        ppLst.Add(pp);
                }

            }
            if (ppLst != null)
            {
                m_logger.WriteInfo("Points Program List");
                return ReturnAMSResult<List<PointProgram>>(ppLst, AMSResultType.Success, PhraseLib.Lookup(ref m_common, "datadefinition-pointsprogram-list", 1, "phrase not found"));
            }
            else
            {
                m_logger.WriteError("No active Point Programs exist");
                return ReturnAMSResult<List<PointProgram>>(null, AMSResultType.ValidationError, PhraseLib.Lookup(ref m_common, "datadefinition-pointsprogram-no-list", 1, "phrase not found"));
            }
        }
        catch (SqlException sqlEx)
        {
            m_logger.WriteError("Failed to get Points Programs" + sqlEx.ToString());
            return ReturnAMSResult<List<PointProgram>>(null, AMSResultType.SQLException, PhraseLib.Lookup(ref m_common, "datadefinition-pointsprogram-failed", 1, "phrase not found") + sqlEx.ToString());
        }
        catch (Exception ex)
        {
            m_logger.WriteError("Failed to get Points Programs" + ex.ToString());
            return ReturnAMSResult<List<PointProgram>>(null, AMSResultType.Exception, PhraseLib.Lookup(ref m_common, "datadefinition-pointsprogram-failed", 1, "phrase not found") + ex.ToString());
        }
    }

    private void FetchPointProgramDetails(ref XmlWriter writer, XmlDocument xmlInput, List<PointProgram> ppList)
    {
        writer.WriteStartElement("PointPrograms");
        foreach (PointProgram pp in ppList)
        {
            //retrieving  name of the each property  into an array
            string[] propertyNames = pp.GetType().GetProperties().Select(p => p.Name).ToArray();
            writer.WriteStartElement("PointProgram");
            foreach (var prop in propertyNames)
            {
                //retrieving value of each property
                object propValue = pp.GetType().GetProperty(prop).GetValue(pp, null);
                if (propValue != null)
                {
                    switch (prop)
                    {
                        case "PointsProgramTransaltionDetails":
                            List<ProgramTranslations> ptLst = (List<ProgramTranslations>)propValue;
                            writer.WriteStartElement("PointsProgramTranslations");
                            foreach (ProgramTranslations ppt in ptLst)
                            {
                                writer.WriteStartElement("PointsProgramTranslation");
                                string[] propertyNames1 = ppt.GetType().GetProperties().Select(p => p.Name).ToArray();
                                foreach (var prop1 in propertyNames1)
                                {
                                    object propValue1 = ppt.GetType().GetProperty(prop1).GetValue(ppt, null);
                                    if (propValue1 != null && !String.IsNullOrEmpty(Convert.ToString(propValue1)))
                                        writer.WriteElementString(prop1, Convert.ToString(propValue1));
                                }
                                writer.WriteEndElement();
                            }
                            writer.WriteEndElement();
                            break;
                        case "PointsProgramMigrationDetails":
                            PointsProgramMigrations ppm = (PointsProgramMigrations)propValue;
                            writer.WriteStartElement("PointsProgramMigration");

                            string[] propertyNames2 = ppm.GetType().GetProperties().Select(p => p.Name).ToArray();
                            foreach (var prop1 in propertyNames2)
                            {
                                object propValue1 = ppm.GetType().GetProperty(prop1).GetValue(ppm, null);
                                if (propValue1 != null && !String.IsNullOrEmpty(Convert.ToString(propValue1)))
                                    writer.WriteElementString(prop1, Convert.ToString(propValue1));

                            }
                            writer.WriteEndElement();
                            break;
                        case "PointsProgramDeletionsDetails":
                            PointsProgramDeletions ppd = (PointsProgramDeletions)propValue;
                            writer.WriteStartElement("PointsProgramDeletion");

                            string[] propertyNames3 = ppd.GetType().GetProperties().Select(p => p.Name).ToArray();
                            foreach (var prop1 in propertyNames3)
                            {
                                object propValue1 = ppd.GetType().GetProperty(prop1).GetValue(ppd, null);
                                if (propValue1 != null && !String.IsNullOrEmpty(Convert.ToString(propValue1)))
                                    writer.WriteElementString(prop1, Convert.ToString(propValue1));

                            }
                            writer.WriteEndElement();
                            break;
                        default:
                            if (!String.IsNullOrEmpty(Convert.ToString(propValue)))
                                writer.WriteElementString(prop, Convert.ToString(propValue));
                            break;
                    }
                }
            }
            writer.WriteEndElement();
        }
        writer.WriteEndElement();
    }
    #endregion

    #region Private methods - Stored value
    private AMSResult<List<StoredValuePrograms>> GetSVProgramsList(string name)
    {
        bool bMultiLanguageEnabled = false;
        StoredValuePrograms pp = null;
        List<StoredValuePrograms> svpLst = null;
        SQLParametersList lstParams = null;
        try
        {
            if ((m_common.Fetch_SystemOption(124) == "1")) { bMultiLanguageEnabled = true; }

            m_logger.WriteInfo("Getting all active SV Program definitions");

            AMSResult<StoredValuePrograms> result = new AMSResult<StoredValuePrograms>();

            string query = "SELECT SVProgramID, Name, Value, CreatedDate,Deleted, SVP.Description,OneUnitPerRec,  ValuePrecision, " +
                      "SVExpireType, SVExpirePeriodType, ExpirePeriod, ExpireTOD, ExpireDate, SVP.SVTypeID, UnitofMeasureLimit, AllowReissue,CMOAStatusFlag,CPEStatusFlag,Visible, AutoDelete, " +
                      "ExtProgramID,ScorecardID, ScorecardDesc, ScorecardBold, AdjustmentUPC,EngineSubTypePKID, ReturnHandlingTypeID, DisallowRedeemInEarnTrans,RedemptionRestrictionID,  " +
                      "AllowNegativeBal, MemberRedemptionID, FuelPartner, AutoRedeem, AllowAdjustments,ExpireCentralServerTZ,VisibleToCustomers " +
                      "FROM StoredValuePrograms AS SVP WITH (NoLock) " +
                      "LEFT JOIN SVTypes AS SVT WITH (NoLock) ON SVT.SVTypeID=SVP.SVTypeID " +
                      "WHERE Deleted=0";

            //Coverity CID - 93913
            DataTable dt = new DataTable();
            if (!String.IsNullOrEmpty(name))
            {
                lstParams = new SQLParametersList();
                lstParams.Add("@name", SqlDbType.VarChar).Value = "%"+name+"%";
                query = query + " AND Name like @name";
                dt = m_dbaccess.ExecuteQuery(DataBases.LogixRT, CommandType.Text, query, lstParams);
            }
            else
            {
                dt = m_dbaccess.ExecuteQuery(DataBases.LogixRT, CommandType.Text, query, null);
            }
            if (dt.Rows.Count > 0)
            {
                svpLst = new List<StoredValuePrograms>();
                foreach (DataRow rw in dt.Rows)
                {
                    #region retreiving data from Storedvalue Program table
                    pp = new StoredValuePrograms();

                    pp.SVProgramID = Convert.ToInt32(rw["SVProgramID"]);
                    pp.Name = (rw["Name"] != DBNull.Value) ? Convert.ToString(rw["Name"]) : null;
                    if (rw["Value"] != DBNull.Value) { pp.Value = Convert.ToDecimal(rw["Value"]); }
                    pp.CreatedDate = (rw["CreatedDate"] == DBNull.Value) ? (DateTime?)null : (DateTime)rw["CreatedDate"];
                    pp.Description = (rw["Description"] != DBNull.Value) ? Convert.ToString(rw["Description"]) : null;
                    pp.Deleted = (rw["Deleted"] != DBNull.Value) ? Convert.ToBoolean(rw["Deleted"]) : false;
                    pp.OneUnitPerRec = (rw["OneUnitPerRec"] != DBNull.Value) ? Convert.ToBoolean(rw["OneUnitPerRec"]) : false;
                    pp.SVExpireType = (rw["SVExpireType"] != DBNull.Value) ? Convert.ToBoolean(rw["SVExpireType"]) : false;
                    if (rw["SVExpirePeriodType"] != DBNull.Value) { pp.SVExpirePeriodType = Convert.ToInt32(rw["SVExpirePeriodType"]); }
                    if (rw["ExpirePeriod"] != DBNull.Value) { pp.ExpirePeriod = Convert.ToInt32(rw["ExpirePeriod"]); }
                    pp.ExpireTOD = (rw["ExpireTOD"] != DBNull.Value) ? Convert.ToString(rw["ExpireTOD"]) : null;
                    pp.ExpireDate = (rw["ExpireDate"] == DBNull.Value) ? (DateTime?)null : (DateTime)rw["ExpireDate"];
                    pp.SVTypeID = (rw["SVTypeID"] != DBNull.Value) ? Convert.ToInt32(rw["SVTypeID"]) : 0;
                    pp.UnitOfMeasureLimit = (rw["UnitOfMeasureLimit"] != DBNull.Value) ? Convert.ToInt32(rw["UnitOfMeasureLimit"]) : 0;
                    pp.AllowReissue = (rw["AllowReissue"] != DBNull.Value) ? Convert.ToBoolean(rw["AllowReissue"]) : false;
                    pp.CMOAStatusFlag = (rw["CMOAStatusFlag"] != DBNull.Value) ? Convert.ToInt32(rw["CMOAStatusFlag"]) : 0;
                    pp.CPEStatusFlag = (rw["CPEStatusFlag"] != DBNull.Value) ? Convert.ToInt32(rw["CPEStatusFlag"]) : 0;
                    if (rw["Visible"] != DBNull.Value) { pp.Visible = Convert.ToBoolean(rw["Visible"]); }
                    pp.AutoDelete = (rw["AutoDelete"] != DBNull.Value) ? Convert.ToBoolean(rw["AutoDelete"]) : false;
                    pp.ExtProgramID = (rw["ExtProgramID"] != DBNull.Value) ? Convert.ToString(rw["ExtProgramID"]) : null;
                    if (rw["ScorecardID"] != DBNull.Value) { pp.ScorecardID = Convert.ToInt32(rw["ScorecardID"]); }
                    pp.ScorecardDesc = (rw["ScorecardDesc"] != DBNull.Value) ? Convert.ToString(rw["ScorecardDesc"]) : null;
                    pp.ScorecardBold = (rw["ScorecardBold"] != DBNull.Value) ? Convert.ToBoolean(rw["ScorecardBold"]) : false;
                    pp.AdjustmentUPC = (rw["AdjustmentUPC"] != DBNull.Value) ? Convert.ToString(rw["AdjustmentUPC"]) : null;
                    if (rw["EngineSubTypePKID"] != DBNull.Value) { pp.EngineSubTypePKID = Convert.ToInt32(rw["EngineSubTypePKID"]); }
                    pp.ReturnHandlingTypeID = (rw["ReturnHandlingTypeID"] != DBNull.Value) ? Convert.ToInt32(rw["ReturnHandlingTypeID"]) : 0;
                    pp.DisallowRedeemInEarnTrans = (rw["DisallowRedeemInEarnTrans"] != DBNull.Value) ? Convert.ToInt32(rw["DisallowRedeemInEarnTrans"]) : 0;
                    pp.AllowNegativeBal = (rw["AllowNegativeBal"] != DBNull.Value) ? Convert.ToInt32(rw["AllowNegativeBal"]) : 0;
                    if (rw["RedemptionRestrictionID"] != DBNull.Value) { pp.RedemptionRestrictionID = Convert.ToInt32(rw["RedemptionRestrictionID"]); }
                    if (rw["MemberRedemptionId"] != DBNull.Value) { pp.MemberRedemptionId = Convert.ToInt32(rw["MemberRedemptionId"]); }
                    if (rw["FuelPartner"] != DBNull.Value) { pp.FuelPartner = Convert.ToBoolean(rw["FuelPartner"]); }
                    if (rw["AutoRedeem"] != DBNull.Value) { pp.AutoRedeem = Convert.ToBoolean(rw["AutoRedeem"]); }
                    if (rw["AllowAdjustments"] != DBNull.Value) { pp.AllowAdjustments = Convert.ToBoolean(rw["AllowAdjustments"]); }
                    if (rw["ExpireCentralServerTZ"] != DBNull.Value) { pp.ExpireCentralServerTZ = Convert.ToBoolean(rw["ExpireCentralServerTZ"]); }
                    pp.VisibleToCustomers = (rw["VisibleToCustomers"] != DBNull.Value) ? Convert.ToBoolean(rw["VisibleToCustomers"]) : false;

                    if (m_systemSettings.IsEngineInstalled(Engines.UE) == true)
                    {
                        string query1 = "SELECT PKID, AllowAnyCustomer from SVProgramsPromoEngineSettings " +
                                       "WHERE SVProgramID = @SVProgramID AND EngineID = @EngineID";
                        lstParams = new SQLParametersList();
                        lstParams.Add("@SVProgramID", SqlDbType.BigInt).Value = pp.SVProgramID;
                        lstParams.Add("@EngineID", SqlDbType.Int).Value = 9;
                        DataTable dt1 = m_dbaccess.ExecuteQuery(DataBases.LogixRT, CommandType.Text, query1, lstParams);

                        if (dt1.Rows.Count > 0)
                        {
                            if (dt1.Rows[0]["AllowAnyCustomer"] != DBNull.Value) { pp.AllowAnyCustomer = Convert.ToBoolean(dt1.Rows[0]["AllowAnyCustomer"]); }
                        }
                    }

                    #endregion

                    #region retrieving data SVProgramTranslations if multi languauge is enabled
                    if (bMultiLanguageEnabled)
                    {
                        string query1 = "select PKID,LanguageID,ScorecardDesc from SVProgramTranslations with (NoLock) " +
                                       "where SVProgramID=@SVProgramID ";
                        lstParams = new SQLParametersList();
                        lstParams.Add("@SVProgramID", SqlDbType.BigInt).Value = pp.SVProgramID;
                        DataTable dt1 = m_dbaccess.ExecuteQuery(DataBases.LogixRT, CommandType.Text, query1, lstParams);
                        List<ProgramTranslations> ptLst = new List<ProgramTranslations>();
                        ProgramTranslations pt = null;
                        if (dt1.Rows.Count > 0)
                        {

                            foreach (DataRow dr in dt1.Rows)
                            {
                                pt = new ProgramTranslations();
                                if (dr["PKID"] != DBNull.Value) pt.PKID = Convert.ToInt32(dr["PKID"]);
                                if (dr["LanguageID"] != DBNull.Value) pt.LanguageID = Convert.ToInt32(dr["LanguageID"]);
                                pt.ScorecardDesc = (dr["ScorecardDesc"] == DBNull.Value) ? "" : Convert.ToString(dr["ScorecardDesc"]);
                                ptLst.Add(pt);
                            }
                        }
                        if (ptLst.Count > 0)
                            pp.SVProgramTransaltionDetails = ptLst;
                    }

                    #endregion

                    if (pp != null)
                        svpLst.Add(pp);
                }

            }
            if (svpLst != null)
            {
                m_logger.WriteInfo("Stored value Program List");
                return ReturnAMSResult<List<StoredValuePrograms>>(svpLst, AMSResultType.Success, PhraseLib.Lookup(ref m_common, "datadefinition-svprogram-list", 1, "phrase not found"));
            }
            else
            {
                m_logger.WriteError("No active stored value Programs exist");
                return ReturnAMSResult<List<StoredValuePrograms>>(null, AMSResultType.ValidationError, PhraseLib.Lookup(ref m_common, "datadefinition-svprogram-no-list", 1, "phrase not found"));
            }
        }
        catch (SqlException sqlEx)
        {
            m_logger.WriteError("Failed to get stored value Programs" + sqlEx.ToString());
            return ReturnAMSResult<List<StoredValuePrograms>>(null, AMSResultType.SQLException, PhraseLib.Lookup(ref m_common, "datadefinition-svprogram-failed", 1, "phrase not found") + sqlEx.ToString());
        }
        catch (Exception ex)
        {
            m_logger.WriteError("Failed to get Stored value Programs" + ex.ToString());
            return ReturnAMSResult<List<StoredValuePrograms>>(null, AMSResultType.Exception, PhraseLib.Lookup(ref m_common, "datadefinition-svprogram-failed", 1, "phrase not found") + ex.ToString());
        }
    }

    private void FetchSVProgramDetails(ref XmlWriter writer, XmlDocument xmlInput, List<StoredValuePrograms> ppList)
    {
        writer.WriteStartElement("StoredValuePrograms");
        foreach (StoredValuePrograms pp in ppList)
        {
            //retrieving  name of the each property  into an array
            string[] propertyNames = pp.GetType().GetProperties().Select(p => p.Name).ToArray();
            writer.WriteStartElement("StoredValueProgram");
            foreach (var prop in propertyNames)
            {
                //retrieving value of each property
                object propValue = pp.GetType().GetProperty(prop).GetValue(pp, null);
                if (propValue != null)
                {
                    switch (prop)
                    {
                        case "SVProgramTransaltionDetails":
                            List<ProgramTranslations> ptLst = (List<ProgramTranslations>)propValue;
                            writer.WriteStartElement("SVProgramTransaltions");
                            foreach (ProgramTranslations ppt in ptLst)
                            {
                                writer.WriteStartElement("SVProgramTransaltion");
                                string[] propertyNames1 = ppt.GetType().GetProperties().Select(p => p.Name).ToArray();
                                foreach (var prop1 in propertyNames1)
                                {
                                    object propValue1 = ppt.GetType().GetProperty(prop1).GetValue(ppt, null);
                                    if (propValue1 != null && !String.IsNullOrEmpty(Convert.ToString(propValue1)))
                                        writer.WriteElementString(prop1, Convert.ToString(propValue1));
                                }
                                writer.WriteEndElement();
                            }
                            writer.WriteEndElement();
                            break;

                        default:
                            if (!String.IsNullOrEmpty(Convert.ToString(propValue)))
                                writer.WriteElementString(prop, Convert.ToString(propValue));
                            break;
                    }
                }
            }
            writer.WriteEndElement();
        }
        writer.WriteEndElement();
    }
    #endregion

    #region Private methods - Trackable coupons programs
    private AMSResult<List<TrackableCouponProgram>> GetTrackableCouponsProgramsList()
    {
        List<TrackableCouponProgram> tpLst = new List<TrackableCouponProgram>();
        try
        {

            m_logger.WriteInfo("Getting all active Trackable Program definitions");

            string query = "SELECT  ProgramID,ExtProgramID,Name,CreatedDate,LastUpdate,LastLoaded,LastLoadMsg,Deleted,Description  " +
                            ",MaxRedeemCount,ExpireDate FROM TrackableCouponProgram WHERE Deleted=0 ";


            tpLst = m_dbaccess.ExecuteQuery(DataBases.LogixRT, CommandType.Text, query, null).ToGenericList<TrackableCouponProgram>();

            if (tpLst != null && tpLst.Count > 0)
            {
                DataTable dt = new DataTable();
                SQLParametersList lstParams = new SQLParametersList();
                foreach (var tcp in tpLst)
                {
                    lstParams.Clear();
                    query = "select count(1) as AssosiatedCouponCount from TrackableCoupons where ProgramID = @ProgramID";
                    lstParams.Add("@ProgramID", SqlDbType.Int).Value = tcp.ProgramID;
                    dt = m_dbaccess.ExecuteQuery(DataBases.LogixXS, CommandType.Text, query, lstParams);
                    if (dt.IsNotEmpty())
                       tcp.AssosiatedCouponCount = dt.Rows[0]["AssosiatedCouponCount"].ConvertToInt32();
                }
                
                m_logger.WriteInfo("Trackable coupon Program List");
                return ReturnAMSResult<List<TrackableCouponProgram>>(tpLst, AMSResultType.Success, PhraseLib.Lookup(ref m_common, "datadefinition-tcprogram-list", 1, "phrase not found"));
            }
            else
            {
                m_logger.WriteError("No active Trackable coupon Programs exist");
                return ReturnAMSResult<List<TrackableCouponProgram>>(null, AMSResultType.ValidationError, PhraseLib.Lookup(ref m_common, "datadefinition-tcprogram-no-list", 1, "phrase not found"));
            }
        }
        catch (SqlException sqlEx)
        {
            m_logger.WriteError("Failed to get Trackable coupon Programs" + sqlEx.ToString());
            return ReturnAMSResult<List<TrackableCouponProgram>>(null, AMSResultType.SQLException, PhraseLib.Lookup(ref m_common, "datadefinition-tcprogram-failed", 1, "phrase not found") + sqlEx.ToString());
        }
        catch (Exception ex)
        {
            m_logger.WriteError("Failed to get Trackable coupon Programs" + ex.ToString());
            return ReturnAMSResult<List<TrackableCouponProgram>>(null, AMSResultType.Exception, PhraseLib.Lookup(ref m_common, "datadefinition-tcprogram-failed", 1, "phrase not found") + ex.ToString());
        }
    }

    private void FetchTracakableProgramDetails(ref XmlWriter writer, XmlDocument xmlInput, List<TrackableCouponProgram> ppList)
    {
        writer.WriteStartElement("TracakableCouponPrograms");
        foreach (TrackableCouponProgram pp in ppList)
        {
            //retrieving  name of the each property  into an array
            string[] propertyNames = pp.GetType().GetProperties().Select(p => p.Name).ToArray();
            writer.WriteStartElement("TracakableCouponProgram");
            foreach (var prop in propertyNames)
            {
                //retrieving value of each property
                object propValue = pp.GetType().GetProperty(prop).GetValue(pp, null);
                if (propValue != null)
                {
                    if (!String.IsNullOrEmpty(Convert.ToString(propValue)))
                        writer.WriteElementString(prop, Convert.ToString(propValue));
                }
            }
            writer.WriteEndElement();
        }
        writer.WriteEndElement();
    }
    #endregion

    #region WebMethods
    /// <summary>
    /// GetPointsPrograms webmethod returns all the active point programs 
    /// </summary>
    /// <param name="GUID">User should provide valid GUID for authentication </param>
    /// <returns>List of active point programs defined in the system.</returns>

    [WebMethod]
    public XmlDocument GetPointsPrograms(string GUID)
    {
        string methodName = "GetPointsPrograms";
        XmlWriter xmlWriter = null;
        StringWriter strWriter = null;
        XmlDocument xmlResponse = new XmlDocument();
        XmlDocument xmlInput = null;

        try
        {
            Startup();
            m_connectorInc.Init_ResponseXML(methodName, ref strWriter, ref xmlWriter);
            if (!m_authInc.Is_Valid_GUID(GUID))
            {
                m_logger.WriteError(m_phraseLib.Detokenize("term.guidnotvalid", GUID, m_common.Get_AppInfo().AppName));
                m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, CMS.CMSException.StatusCodes.INVALID_GUID, m_phraseLib.Detokenize("term.guidnotvalid", GUID, m_common.Get_AppInfo().AppName), false);
            }

            else
            {
                var response = GetPointsProgramsList(string.Empty);
                CMS.CMSException.StatusCodes statusCode = CMS.CMSException.StatusCodes.GENERAL_ERROR;
                if (response.ResultType == CMS.AMS.Models.AMSResultType.Success)
                {
                    statusCode = CMS.CMSException.StatusCodes.SUCCESS;
                    FetchPointProgramDetails(ref xmlWriter, xmlInput, response.Result);
                    m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, statusCode, m_phraseLib.Lookup("term.success"), true);
                }
                else
                {
                    m_logger.WriteError(response.MessageString);
                    m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, statusCode, response.MessageString, false);
                }
            }
        }
        catch (Exception ex)
        {
            m_logger.WriteError("Failed to get PointPrograms. Please see the error log!");
            m_errHandler.ProcessError(ex);
            ProcessException(ex, methodName, ref xmlWriter, ref strWriter);
        }

        CloseResponseXml(ref xmlWriter, ref strWriter, ref xmlResponse);
        Shutdown();

        return xmlResponse;
    }

    /// <summary>
    ///  GetPointProgramByName webmethod returns point programs for the provided name
    /// </summary>
    /// <param name="GUID">User should provide valid GUID for authentication </param>
    /// <param name="name">User should provide valid name of point program </param>
    /// <returns>Returns single point program details if the entered name is perfect match or returns list of point programs based on the like pattern of the entered name.  </returns>
    [WebMethod]
    public XmlDocument GetPointProgramByName(string GUID, string Name)
    {
        string methodName = "GetPointProgramByName";
        XmlWriter xmlWriter = null;
        StringWriter strWriter = null;
        XmlDocument xmlResponse = new XmlDocument();
        XmlDocument xmlInput = null;

        try
        {
            Startup();
            m_connectorInc.Init_ResponseXML(methodName, ref strWriter, ref xmlWriter);
            if (!m_authInc.Is_Valid_GUID(GUID))
            {
                m_logger.WriteError(m_phraseLib.Detokenize("term.guidnotvalid", GUID, m_common.Get_AppInfo().AppName));
                m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, CMS.CMSException.StatusCodes.INVALID_GUID, m_phraseLib.Detokenize("term.guidnotvalid", GUID, m_common.Get_AppInfo().AppName), false);
            }
            else if (String.IsNullOrEmpty(Name))
            {
                m_logger.WriteError(m_phraseLib.Detokenize("term.invalidnameparameter", GUID, m_common.Get_AppInfo().AppName));
                m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, CMS.CMSException.StatusCodes.MISSING_PARAMETER, m_phraseLib.Detokenize("term.invalidnameparameter", GUID, m_common.Get_AppInfo().AppName), false);
            }
            else
            {
                var response = GetPointsProgramsList(Name);
                CMS.CMSException.StatusCodes statusCode = CMS.CMSException.StatusCodes.GENERAL_ERROR;
                if (response.ResultType == CMS.AMS.Models.AMSResultType.Success)
                {
                    statusCode = CMS.CMSException.StatusCodes.SUCCESS;
                    FetchPointProgramDetails(ref xmlWriter, xmlInput, response.Result);
                    m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, statusCode, m_phraseLib.Lookup("term.success"), true);
                }
                else
                {
                    m_logger.WriteError(response.MessageString);
                    m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, statusCode, response.MessageString, false);
                }
            }
        }
        catch (Exception ex)
        {
            m_logger.WriteError("Failed to get PointProgram details by name. Please see the error log!");
            m_errHandler.ProcessError(ex);
            ProcessException(ex, methodName, ref xmlWriter, ref strWriter);
        }

        CloseResponseXml(ref xmlWriter, ref strWriter, ref xmlResponse);
        Shutdown();

        return xmlResponse;
    }

    /// <summary>
    /// GetSVPrograms webmethod returns all the active SV programs 
    /// </summary>
    /// <param name="GUID">User should provide valid GUID for authentication </param>
    /// <returns>List of active Stored value programs defined in the system.</returns>

    [WebMethod]
    public XmlDocument GetSVPrograms(string GUID)
    {
        string methodName = "GetSVPrograms";
        XmlWriter xmlWriter = null;
        StringWriter strWriter = null;
        XmlDocument xmlResponse = new XmlDocument();
        XmlDocument xmlInput = null;

        try
        {
            Startup();
            m_connectorInc.Init_ResponseXML(methodName, ref strWriter, ref xmlWriter);
            if (!m_authInc.Is_Valid_GUID(GUID))
            {
                m_logger.WriteError(m_phraseLib.Detokenize("term.guidnotvalid", GUID, m_common.Get_AppInfo().AppName));
                m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, CMS.CMSException.StatusCodes.INVALID_GUID, m_phraseLib.Detokenize("term.guidnotvalid", GUID, m_common.Get_AppInfo().AppName), false);
            }
            else
            {
                var response = GetSVProgramsList(string.Empty);
                CMS.CMSException.StatusCodes statusCode = CMS.CMSException.StatusCodes.GENERAL_ERROR;
                if (response.ResultType == CMS.AMS.Models.AMSResultType.Success)
                {
                    statusCode = CMS.CMSException.StatusCodes.SUCCESS;
                    FetchSVProgramDetails(ref xmlWriter, xmlInput, response.Result);
                    m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, statusCode, m_phraseLib.Lookup("term.success"), true);
                }
                else
                {
                    m_logger.WriteError(response.MessageString);
                    m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, statusCode, response.MessageString, false);
                }
            }
        }
        catch (Exception ex)
        {
            m_logger.WriteError("Failed to get SVPrograms. Please see the error log!");
            m_errHandler.ProcessError(ex);
            ProcessException(ex, methodName, ref xmlWriter, ref strWriter);
        }

        CloseResponseXml(ref xmlWriter, ref strWriter, ref xmlResponse);
        Shutdown();

        return xmlResponse;
    }

    /// <summary>
    ///  GetSVProgramByName webmethod returns Stored value programs for the provided name
    /// </summary>
    /// <param name="GUID">User should provide valid GUID for authentication </param>
    /// <param name="name">User should provide valid name of Stored value program </param>
    /// <returns>Returns single Stored value program details if the entered name is perfect match or returns list of Stored value programs based on the like pattern of the entered name.  </returns>
    [WebMethod]
    public XmlDocument GetSVProgramByName(string GUID, string Name)
    {
        string methodName = "GetSVProgramByName";
        XmlWriter xmlWriter = null;
        StringWriter strWriter = null;
        XmlDocument xmlResponse = new XmlDocument();
        XmlDocument xmlInput = null;

        try
        {
            Startup();
            m_connectorInc.Init_ResponseXML(methodName, ref strWriter, ref xmlWriter);
            if (!m_authInc.Is_Valid_GUID(GUID))
            {
                m_logger.WriteError(m_phraseLib.Detokenize("term.guidnotvalid", GUID, m_common.Get_AppInfo().AppName));
                m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, CMS.CMSException.StatusCodes.INVALID_GUID, m_phraseLib.Detokenize("term.guidnotvalid", GUID, m_common.Get_AppInfo().AppName), false);
            }
            else if (String.IsNullOrEmpty(Name))
            {
                m_logger.WriteError(m_phraseLib.Detokenize("term.invalidnameparameter", GUID, m_common.Get_AppInfo().AppName));
                m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, CMS.CMSException.StatusCodes.MISSING_PARAMETER, m_phraseLib.Detokenize("term.invalidnameparameter", GUID, m_common.Get_AppInfo().AppName), false);
            }
            else
            {
                var response = GetSVProgramsList(Name);
                CMS.CMSException.StatusCodes statusCode = CMS.CMSException.StatusCodes.GENERAL_ERROR;
                if (response.ResultType == CMS.AMS.Models.AMSResultType.Success)
                {
                    statusCode = CMS.CMSException.StatusCodes.SUCCESS;
                    FetchSVProgramDetails(ref xmlWriter, xmlInput, response.Result);
                    m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, statusCode, m_phraseLib.Lookup("term.success"), true);
                }
                else
                {
                    m_logger.WriteError(response.MessageString);
                    m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, statusCode, response.MessageString, false);
                }
            }
        }
        catch (Exception ex)
        {
            m_logger.WriteError("Failed to get SVProgram by name. Please see the error log!");
            m_errHandler.ProcessError(ex);
            ProcessException(ex, methodName, ref xmlWriter, ref strWriter);
        }

        CloseResponseXml(ref xmlWriter, ref strWriter, ref xmlResponse);
        Shutdown();

        return xmlResponse;
    }

    /// <summary>
    ///  GetTrackableCouponPrograms webmethod returns all the active trackable coupon programs 
    /// </summary>
    /// <param name="GUID">User should provide valid GUID for authentication</param>
    /// <returns>List of active trackable coupon programs defined in the system.</returns>

    [WebMethod]
    public XmlDocument GetTrackableCouponPrograms(string GUID)
    {
        string methodName = "GetTrackableCouponPrograms";
        XmlWriter xmlWriter = null;
        StringWriter strWriter = null;
        XmlDocument xmlResponse = new XmlDocument();
        XmlDocument xmlInput = null;

        try
        {
            Startup();
            m_connectorInc.Init_ResponseXML(methodName, ref strWriter, ref xmlWriter);
            if (!m_authInc.Is_Valid_GUID(GUID))
            {
                m_logger.WriteError(m_phraseLib.Detokenize("term.guidnotvalid", GUID, m_common.Get_AppInfo().AppName));
                m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, CMS.CMSException.StatusCodes.INVALID_GUID, m_phraseLib.Detokenize("term.guidnotvalid", GUID, m_common.Get_AppInfo().AppName), false);
            }

            else
            {
                var response = GetTrackableCouponsProgramsList();
                CMS.CMSException.StatusCodes statusCode = CMS.CMSException.StatusCodes.GENERAL_ERROR;
                if (response.ResultType == CMS.AMS.Models.AMSResultType.Success)
                {
                    statusCode = CMS.CMSException.StatusCodes.SUCCESS;
                    FetchTracakableProgramDetails(ref xmlWriter, xmlInput, response.Result);
                    m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, statusCode, m_phraseLib.Lookup("term.success"), true);
                }
                else
                {
                    m_logger.WriteError(response.MessageString);
                    m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, statusCode, response.MessageString, false);
                }
            }
        }
        catch (Exception ex)
        {
            m_logger.WriteError("Failed to get Trackable Coupon Programs. Please see the error log!");
            m_errHandler.ProcessError(ex);
            ProcessException(ex, methodName, ref xmlWriter, ref strWriter);
        }

        CloseResponseXml(ref xmlWriter, ref strWriter, ref xmlResponse);
        Shutdown();

        return xmlResponse;
    }
    #endregion

}
