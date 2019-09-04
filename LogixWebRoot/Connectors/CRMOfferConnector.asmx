<%@ WebService Language="C#" Class="CRMOfferConnector" %>

/*

CRM Offer Connector web service
$Id: CRMOfferConnector.asmx 134627 2019-02-26 10:38:50Z sk185403 $

version:7.3.1.138972.Official Build (SUSDAY10202)

Things are currently public for ease of testing...

*/

using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Web;
using System.Web.Services;
using System.Xml;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;

[WebService(Namespace = "http://www.copienttech.com/CRMOfferConnector/")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
public class CRMOfferConnector : System.Web.Services.WebService
{
    //From CLOUDSOL-13 where I worked on issues HEB was having on 11/Mar/13 1:24 PM
    //Resolution to last week's issues:
    //
    //Unable to load-balance web servers on postOffers(). CRMOfferConnector and CRMImportAgent
    //are both using either SystemOptions table OptionID=29 which is WORKSPACE_FILE_PATH_SYSTEM_OPTION
    //and InterfaceOptions table OptionID=44 which is IMPORT_FILES_SOURCE_PATH_INTERFACE_OPTION. 
    //CRMImportAgent uses IMPORT_FILES_SOURCE_PATH_INTERFACE_OPTION if defined else it uses 
    //WORKSPACE_FILE_PATH_SYSTEM_OPTION. CRMOfferConnector was doing the same in the same order. 
    //I needed to reverse the order in CRMOfferConnector to allow for separate specification of file drop
    //and file pickup locations. What appears to be the best and most efficient way to operated is if 
    //the CRMImportAgent operates off its local path picking files up at say:
    //C:\Copient\AgentFiles\WebUploads while each of the web servers gets configure identically
    //to drop the files using UNC path similar to: \\server\Copient\AgentFiles\WebUploads. For 
    //support reasons, WORKSPACE_FILE_PATH_SYSTEM_OPTION must be configured identically across
    //all web servers. There can only be one CRMImportAgent running in the system writing to the 
    //SQL Server tables, but there can be multiple web servers all dropping files for the CRMImportAgent
    //to pick up. 

    private const string sVersion = "7.3.1.138972";

    #region Static Path Functions

    private const int WORKSPACE_FILE_PATH_SYSTEM_OPTION = 29;
    private const int EXPORT_FILES_TARGET_PATH_INTERFACE_OPTION = 43;
    private const int IMPORT_FILES_SOURCE_PATH_INTERFACE_OPTION = 44;
    private static Hashtable NamingTable = new Hashtable(50);
    private char[] invalidFileNameChars = Path.GetInvalidFileNameChars();

    public static string getSystemStringOptionAsFilePath(int optionID, Copient.CommonInc commonInterface)
    {
        return AppendFilePath(commonInterface.Fetch_SystemOption(optionID));
    }

    public static string getInterfaceStringOptionAsFilePath(int optionID, Copient.CommonInc commonInterface)
    {
        return AppendFilePath(commonInterface.Fetch_InterfaceOption(optionID));
    }

    private static string AppendFilePath(string path)
    {

        if (!string.IsNullOrWhiteSpace(path) && !path.EndsWith("\\"))
        {
            path += "\\";
        }

        return path;
    }
    public static string exportFilePath(Copient.CommonInc commonInterface)
    {
        return GetFilePath(EXPORT_FILES_TARGET_PATH_INTERFACE_OPTION, commonInterface);
    }

    public static string importFilePath(Copient.CommonInc commonInterface)
    {
        return GetFilePath(IMPORT_FILES_SOURCE_PATH_INTERFACE_OPTION, commonInterface);
    }

    private static string GetFilePath(int interfaceOption, Copient.CommonInc common)
    {
        string path = getInterfaceStringOptionAsFilePath(interfaceOption, common);
        if (string.IsNullOrWhiteSpace(path))
        {
            path = getSystemStringOptionAsFilePath(WORKSPACE_FILE_PATH_SYSTEM_OPTION, common);
        }
        CreateDirectory(path);
        return path;
    }

    private static void CreateDirectory(string path)
    {
        if (!string.IsNullOrWhiteSpace(path))
        {
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
        }
    }

    #endregion

    public class CRMOfferConnectorException : System.ApplicationException
    {
        public CRMOfferConnectorException(string msg)
          : base(msg)
        { }

    } // end class CRMOfferConnectorException

    public class logger
    {
        private string m_logFileName;
        private Copient.CommonInc m_commonRoutines;

        public logger(string logFileName, Copient.CommonInc commonRoutines)
        {
            m_logFileName = logFileName;
            m_commonRoutines = commonRoutines;
        }

        public void log(string msg)
        {
            m_commonRoutines.Write_Log(m_logFileName, msg, true);
        }

    } // end class logger




    private Copient.CommonInc m_commonRoutines;
    private logger m_logger;

    public CRMOfferConnector()
    {
        m_commonRoutines = new Copient.CommonInc();
        m_commonRoutines.AppName = "CRMOfferConnector";
        m_commonRoutines.Open_LogixRT(); // to get file logging to work, LogixRT has to be opened separately

        m_logger = new logger("CRMOfferConnector." + DateTime.Now.ToString("yyyyMMdd") + ".txt", m_commonRoutines);
        //m_commonRoutines.Write_Log( m_logFileName, "CRMOfferConnector", true );
        //m_commonRoutines.Write_Log( m_logFileName, this.Context.Request.offersToExportTable, false );
    } // end CRMOfferConnector() ctor


    ~CRMOfferConnector()
    { // if only this were deterministic...
        m_commonRoutines.Close_LogixRT();
    }

    public void log(string msg)
    {
        m_logger.log(msg);
    }



    public class ExternalInterfaceID
    {
        public Guid m_eiguid;          // this is their GUID for the connector
        public int m_extInterfaceID;   // this is their InboundCRMEngineID from cpe_incentives or offers

        private const int CONNECTOR_ID = 47;
        private const string m_GUIDValidationQueryFormat =
            @"
                SELECT count(*) AS matches
                FROM Connectors WITH (NoLock)
                WHERE connectorid = {0} AND installed = 'true' AND (
                    ( usesguids = 'false' )
                    OR
                    ( usesguids = 'true' AND ( SELECT count(*) AS c FROM connectorguids WITH (NoLock) WHERE connectorid = {0} AND GUID = {1} ) > 0 )
                )
             "; // end m_GUIDValidationQueryFormat

        private const string m_ExtInterfaceIDValidationQueryFormat =
            @"
                SELECT COUNT(*) AS matches
                FROM ExtCRMInterfaces WITH (NoLock)
                WHERE ExtInterfaceID = {0} AND active = 1 AND deleted = 0
            ";

        private const string m_ExtInterfaceIDIsOutboundEnabled =
            @"
                SELECT OutboundEnabled
                FROM ExtCRMInterfaces WITH (NoLock)
                WHERE ExtInterfaceID = {0} AND active = 1 AND deleted = 0
            ";
        public ExternalInterfaceID()
        { /* default constructor, such that members get default initialization */ }

        public ExternalInterfaceID(Guid uuid, int extid)
        {
            m_eiguid = uuid;
            m_extInterfaceID = extid;
        }

        override public string ToString()
        {
            return String.Format("{0} : {1}", m_extInterfaceID, m_eiguid);
        }

        public string GUIDvalidationQuery()
        {
            return string.Format(m_GUIDValidationQueryFormat, "@ConnectorID", "@Guid");
        }

        public string ExtIDValidationQuery()
        {
            return string.Format(m_ExtInterfaceIDValidationQueryFormat, "@ExtInterfaceID");
        }

        public string outboundEnabled()
        {
            return string.Format(m_ExtInterfaceIDIsOutboundEnabled, "@ExtInterfaceID");
        }

        public bool validateGUID(Copient.CommonInc commonInterface)
        {
            commonInterface.Open_LogixRT();
            commonInterface.QueryStr = GUIDvalidationQuery();
            commonInterface.DBParameters.Add("@ConnectorID", SqlDbType.Int).Value = CONNECTOR_ID;
            commonInterface.DBParameters.Add("@Guid", SqlDbType.NVarChar).Value = m_eiguid.ToString();

            DataTable MatchingGUIDcount = commonInterface.ExecuteQuery(Copient.DataBases.LogixRT);
            if (MatchingGUIDcount.Rows.Count < 0 || ((int)(MatchingGUIDcount.Rows[0]["matches"])) < 1)
            {
                throw new CRMOfferConnectorException(string.Format("The external connector GUID supplied was invalid( {0} )", ToString()));
            }
            return true;

        } //end method validateGUID()

        public bool validateExtInterfaceID(Copient.CommonInc commonInterface)
        {
            commonInterface.Open_LogixRT();
            commonInterface.QueryStr = ExtIDValidationQuery();
            commonInterface.DBParameters.Add("@ExtInterfaceID", SqlDbType.Int).Value = m_extInterfaceID;
            DataTable MatchingExtIDcount = commonInterface.ExecuteQuery(Copient.DataBases.LogixRT);
            if (MatchingExtIDcount.Rows.Count < 0 || ((int)(MatchingExtIDcount.Rows[0]["matches"])) < 1)
            {
                throw new CRMOfferConnectorException(string.Format("The external connector identifier supplied was invalid( {0} )", ToString()));
            }
            return true;
        } // end method validateExtInterfaceID()

        public bool isOutboundEnabled(Copient.CommonInc commonInterface)
        {
            commonInterface.Open_LogixRT();
            commonInterface.QueryStr = outboundEnabled();
            commonInterface.DBParameters.Add("@ExtInterfaceID", SqlDbType.Int).Value = m_extInterfaceID;
            DataTable OutboundEnabledRow = commonInterface.ExecuteQuery(Copient.DataBases.LogixRT);

            if (OutboundEnabledRow.Rows.Count == 1)
            {
                return ((bool)OutboundEnabledRow.Rows[0]["OutboundEnabled"]);
            }
            else
            {
                return false;
            }
        } // end method isOutboundEnabled


        public void verify(Copient.CommonInc commonInterface)
        {
            validateGUID(commonInterface);
            validateExtInterfaceID(commonInterface);
        } // end method verify

    } // end class ExternalInterfaceID

    [WebMethod] //------------------------------------------------------------------------------------------------------------------
    public string heartbeat(ExternalInterfaceID ei)
    {
        log(string.Format("heartbeat( {0} )", ei));
        ei.verify(m_commonRoutines);
        return ei.ToString();
    } // end heartbeat()




    /* This class encapsulates a set of offerid/ExtOfferID/Offer definitions sets */
    public class Offers
    {

        public class OfferDefinition
        {

            public class EncapsulatedXML
            {
                public string m_encoded_xml;

                public EncapsulatedXML()
                { /* do nothing default constructor */ }

                public EncapsulatedXML(string xmlFilename)
                {
                    // m_encoded_xml = xmlFilename;
                    m_encoded_xml = "";
                    if (xmlFilename == null)
                    {
                        throw new CRMOfferConnectorException(string.Format("The offer file name provided was NULL!", xmlFilename));
                    }
                    if (xmlFilename.Length == 0 || !File.Exists(xmlFilename))
                    {
                        throw new CRMOfferConnectorException(string.Format("Offer file was not found!( {0} )", xmlFilename));
                    }
                    string xml = File.ReadAllText(xmlFilename);
                    m_encoded_xml = HttpUtility.HtmlEncode(xml);

                } // end EncapsulatedXML ctor

                public string decode()
                {
                    return HttpUtility.HtmlDecode(m_encoded_xml);
                }


                public bool writeToFile(string xmlFilename)
                {
                    string xml = decode();
                    File.WriteAllText(xmlFilename, xml);
                    return true;
                }

            } // end class Encapsulated XML


            public long m_offerID;
            public string m_extOfferID;
            public string m_filename;
            public EncapsulatedXML m_offerXML;
            public int m_engineID;
            internal int m_buyerId;
            public string m_locations;

            public bool hasExternalID() { return m_extOfferID != null && m_extOfferID.Trim().Length > 0; }

            public OfferDefinition()
            {
                m_filename = "";
                m_offerXML = new EncapsulatedXML();
                m_engineID = (int)Copient.CommonInc.InstalledEngines.CPE;
                m_locations = "";
            }

            public OfferDefinition(DataRow offerInfoRow, string exportedOfferPath)
            {
                m_offerID = (long)offerInfoRow["OfferID"];
                m_extOfferID = (offerInfoRow["extofferid"] is DBNull) ? "" : ((string)offerInfoRow["extofferid"]);
                m_filename = (offerInfoRow["filename"] is DBNull) ? "" : ((string)offerInfoRow["filename"]);
                m_engineID = (int)offerInfoRow["EngineID"];
                if (m_filename.Length > 0)
                {
                    m_offerXML = new EncapsulatedXML(exportedOfferPath + m_filename);
                }
                m_locations = "";
            }

            public const string CPE_EXISTENCE_QUERY_FORMAT =
            @"
                SELECT TOP( 1 ) IncentiveID
                    FROM cpe_incentives
                    WHERE Deleted=0 and ClientOfferID = {0} AND InboundCRMEngineID = {1}
            ";

            public const string CM_EXISTENCE_QUERY_FORMAT =
            @"
                SELECT TOP( 1 ) OfferID as IncentiveID
                    FROM Offers
                    WHERE Deleted=0 and ExtOfferID = {0} AND InboundCRMEngineID = {1}
            ";

            public string existenceQuery()
            {
                if ((m_engineID == (int)Copient.CommonInc.InstalledEngines.CPE) || (m_engineID == (int)Copient.CommonInc.InstalledEngines.UE))
                {
                    return string.Format(CPE_EXISTENCE_QUERY_FORMAT, "@ExtOfferId", "@EngineId");
                }
                else
                {
                    return string.Format(CM_EXISTENCE_QUERY_FORMAT, "@ExtOfferId", "@EngineId");
                }
            }


            // look for it by external offer id
            public bool offerExists(Copient.CommonInc commonInterface, int eid)
            {
                commonInterface.Open_LogixRT();
                commonInterface.QueryStr = existenceQuery();
                commonInterface.DBParameters.Add("@ExtOfferId", SqlDbType.NVarChar).Value = m_extOfferID;
                commonInterface.DBParameters.Add("@EngineId", SqlDbType.Int).Value = eid;

                DataTable matchingOffersCount = commonInterface.ExecuteQuery(Copient.DataBases.LogixRT);
                if (matchingOffersCount.Rows.Count > 0)
                {   //
                    m_offerID = ((long)matchingOffersCount.Rows[0]["IncentiveID"]);
                    return true;
                }
                return false;
            }

            public bool InsertExternalID(Copient.CommonInc commonInterface, string ExtID, int InboundCRMID)
            {
                try
                {
                    commonInterface.QueryStr = "pt_ExtOfferID_Insert";
                    commonInterface.Open_LRTsp();
                    commonInterface.LRTsp.Parameters.Add("@ExtOfferID", SqlDbType.NVarChar, 20).Value = ExtID;
                    commonInterface.LRTsp.Parameters.Add("@InboundCRMEngineID", SqlDbType.Int).Value = InboundCRMID;
                    commonInterface.LRTsp.ExecuteNonQuery();
                    commonInterface.Close_LRTsp();
                    return true;
                }
                catch (Exception ex)
                {
                    return false;
                }
            }

            public const string CM_TRANSFER_EXISTENCE_QUERY_FORMAT =
            @"
                SELECT TOP( 1 ) OfferID
                    FROM Offers
                    WHERE Deleted=0 and ProductionID = {0}
            ;";

            public string transferCmExistenceQuery()
            {
                if (m_engineID == (int)Copient.CommonInc.InstalledEngines.CM)
                {
                    return string.Format(CM_TRANSFER_EXISTENCE_QUERY_FORMAT, "@OfferId");
                }
                else
                {
                    return ("");
                }
            }

            // look for it by production offer id
            public bool transferCmOfferExists(Copient.CommonInc commonInterface)
            {
                commonInterface.Open_LogixRT();
                commonInterface.QueryStr = transferCmExistenceQuery();
                commonInterface.DBParameters.Add("@OfferId", SqlDbType.BigInt).Value = m_offerID;
                DataTable matchingOffersCount = commonInterface.ExecuteQuery(Copient.DataBases.LogixRT);
                if (matchingOffersCount.Rows.Count > 0)
                {   //
                    m_offerID = ((long)matchingOffersCount.Rows[0]["OfferID"]);
                    return true;
                }
                return false;
            }

            public const string UE_TRANSLATE_EXISTENCE_QUERY_FORMAT =
           @"
                SELECT TOP( 1 ) IncentiveID
                    FROM CPE_Incentives 
                    WHERE Deleted=0 and ClientOfferID = {0}
                    AND InboundCRMEngineID = {1}
            ;";

            public string translateUeExistenceQuery()
            {
                if (m_engineID == (int)Copient.CommonInc.InstalledEngines.UE)
                {
                    return string.Format(UE_TRANSLATE_EXISTENCE_QUERY_FORMAT, "@ClientOfferId", "@InboundCRMEngineID");
                }
                else
                {
                    return ("");
                }
            }

            // look for it by production offer id
            public bool translateUeOfferExists(int inboundCRMEngineID, Copient.CommonInc commonInterface)
            {
                commonInterface.Open_LogixRT();
                commonInterface.QueryStr = translateUeExistenceQuery();
                commonInterface.DBParameters.Add("@ClientOfferId", SqlDbType.NVarChar).Value = m_extOfferID;
                commonInterface.DBParameters.Add("@InboundCRMEngineID", SqlDbType.BigInt).Value = inboundCRMEngineID;
                DataTable matchingOffersCount = commonInterface.ExecuteQuery(Copient.DataBases.LogixRT);
                if (matchingOffersCount.Rows.Count > 0)
                {   // 
                    m_offerID = ((long)matchingOffersCount.Rows[0]["IncentiveID"]);
                    return true;
                }
                return false;
            }


            private void logOfferCreation(Copient.CommonInc commonInterface, long userID, int buyerId = -1)
            {
                const int ACTIVITY_TYPE_OFFER = 3;
                const int SYS_OPTION_DEFAULT_LANGUAGE_ID = 1;
                int defaultLanguageID = int.Parse(commonInterface.Fetch_SystemOption(SYS_OPTION_DEFAULT_LANGUAGE_ID));
                commonInterface.Activity_Log(ACTIVITY_TYPE_OFFER, m_offerID, userID, Copient.PhraseLib.Lookup("history.offer-create", defaultLanguageID), buyerId);
            }

            public bool addNewSkeletonOffer(Copient.CommonInc commonInterface, int eid)
            {
                const long ADMIN_USER_ID = 1;
                commonInterface.Open_LRTsp();

                commonInterface.LRTsp.CommandText = "pt_Offers_Insert";
                commonInterface.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 100).Value = string.Format("ImportedOffer_{0}_{1}", eid, m_extOfferID);
                commonInterface.LRTsp.Parameters.Add("@EngineID", SqlDbType.Int).Value = m_engineID;
                commonInterface.LRTsp.Parameters.Add("@EngineSubTypeID", SqlDbType.Int).Value = 0;
                commonInterface.LRTsp.Parameters.Add("@IsTemplate", SqlDbType.Bit).Value = 0;
                commonInterface.LRTsp.Parameters.Add("@AdminID", SqlDbType.Int).Value = ADMIN_USER_ID;
                commonInterface.LRTsp.Parameters.Add("@BuyerID", SqlDbType.Int).Value = (m_buyerId == -1 ? (object)DBNull.Value : m_buyerId);
                commonInterface.LRTsp.Parameters.Add("@OfferID", SqlDbType.BigInt).Direction = ParameterDirection.Output;
                commonInterface.LRTsp.ExecuteNonQuery();

                m_offerID = (long)commonInterface.LRTsp.Parameters["@OfferID"].Value;

                commonInterface.Close_LRTsp();

                //* If import was successful, the connector writes to the ActivityLog a record for the offer stating that it was created by the inbound CRM process.
                logOfferCreation(commonInterface, ADMIN_USER_ID, m_buyerId);

                return true;
            } // end method addNewSkeletonOffer()

            public bool associateExternalIDToOffer(Copient.CommonInc commonInterface)
            {   // update the external Offerid  
                if ((m_engineID == (int)Copient.CommonInc.InstalledEngines.CPE) || (m_engineID == (int)Copient.CommonInc.InstalledEngines.UE) || (m_engineID == (int)Copient.CommonInc.InstalledEngines.Website))
                {
                    commonInterface.QueryStr = string.Format("UPDATE CPE_Incentives SET ClientOfferID = SUBSTRING( {0}, 1, 20 )  WHERE IncentiveID = {1}", "@ExtOfferId", "@OfferId");
                }
                else
                {
                    commonInterface.QueryStr = string.Format("UPDATE Offers SET ExtOfferID = SUBSTRING( {0}, 1, 20 )  WHERE OfferID = {1}", "@ExtOfferId", "@OfferId");
                }
                commonInterface.DBParameters.Add("@ExtOfferId", SqlDbType.NVarChar).Value = m_extOfferID;
                commonInterface.DBParameters.Add("@OfferId", SqlDbType.BigInt).Value = m_offerID;
                commonInterface.ExecuteNonQuery(Copient.DataBases.LogixRT);
                return commonInterface.RowsAffected > 0;
            }

            public const string CPE_CRM_ENGINE_ID_UPDATE_FORMAT =
            @"
                UPDATE CPE_Incentives WITH ( RowLock )
                    SET CRMEngineID = {0}, InboundCRMEngineID = {1}, LastUpdate = GETDATE()
                    WHERE IncentiveID = {2}
            ";

            public const string CM_CRM_ENGINE_ID_UPDATE_FORMAT =
            @"
                UPDATE Offers WITH ( RowLock )
                    SET CRMEngineID = {0}, InboundCRMEngineID = {1}, LastUpdate = GETDATE()
                    WHERE OfferID = {2}
            ";


            public string CRMEngineIDUpdateQuery()
            {
                if ((m_engineID == (int)Copient.CommonInc.InstalledEngines.CPE) || (m_engineID == (int)Copient.CommonInc.InstalledEngines.UE) || (m_engineID == (int)Copient.CommonInc.InstalledEngines.Website))
                {
                    return string.Format(CPE_CRM_ENGINE_ID_UPDATE_FORMAT, "@CRMEngineID", "@InboundCRMEngineID", "@OfferID");
                }
                else
                {
                    return string.Format(CM_CRM_ENGINE_ID_UPDATE_FORMAT, "@CRMEngineID", "@InboundCRMEngineID", "@OfferID");
                }
            }

            // * When an offer is submitted for import via the connector, please ensure that the connector will
            //      set the CPE_Incentive record's InboundCRMEngineID and CRMEngineID fields are set to 5 (eid)
            public bool updateOfferCRMEngineID(Copient.CommonInc commonInterface, int eid, bool outEnabled)
            {
                int CRMEngineID = 0;
                int InboundCRMEngineID = 0;

                if (outEnabled)
                {
                    CRMEngineID = eid;
                    InboundCRMEngineID = eid;
                }
                else
                {
                    CRMEngineID = 5;  //Default CRM communications service
                    InboundCRMEngineID = eid;
                }

                commonInterface.QueryStr = CRMEngineIDUpdateQuery();
                commonInterface.DBParameters.Add("@CRMEngineID", SqlDbType.Int).Value = CRMEngineID;
                commonInterface.DBParameters.Add("@InboundCRMEngineID", SqlDbType.Int).Value = InboundCRMEngineID;
                commonInterface.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = m_offerID;
                commonInterface.ExecuteNonQuery(Copient.DataBases.LogixRT);
                return true;
            } // end method updateOfferCRMEngineID()


            public const string IMPORT_OFFER_FILENAME_FORTMAT = "Offer{0}-{1}.xml";
            public string importFileName()
            {  //* The filename has the format "Offer[IncentiveID]-[DateTimeStamp].xml".
                if (m_filename.Length < 1)
                {
                    string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
                    m_filename = string.Format(IMPORT_OFFER_FILENAME_FORTMAT, m_offerID, timestamp);
                }
                return m_filename;
            }

            public bool writeToFile(string path)
            { //* The connector writes the XML to a file.
                string outFile = path + importFileName();
                return m_offerXML.writeToFile(outFile);
            }


            //* The connector writes a record to the CRMImportQueue table (which includes the offer's ID and the date).
            //    The connector writes a record to the CRMFileImportQueue table (which
            //    includes a reference to the previous record's PKID and the name of the file),
            //    setting the CRMStatusFlag to 0 to indicate that the record is awaiting processing.
            public bool addToImportQueue(Copient.CommonInc commonInterface, int eid,string operation)
            {
                commonInterface.QueryStr = string.Format(@"
                    IF NOT EXISTS( SELECT pkid FROM CRMImportQueue WHERE OfferId = {0} )
                    BEGIN
                        INSERT INTO CRMImportQueue WITH (RowLock) ( OfferId, LastUpdate, Deleted, Operation ) VALUES ( {0}, GETDATE(), 0 , {1})
                    END
                    ELSE
                    BEGIN
                        UPDATE CRMImportQueue WITH (RowLock) SET LastUpdate = GETDATE(), Deleted = 0 , Operation = {1} WHERE OfferId = {0}
                    END
                ", "@OfferId","@Operation");
                commonInterface.DBParameters.Add("@OfferId", SqlDbType.BigInt).Value = m_offerID;
                commonInterface.DBParameters.Add("@Operation", SqlDbType.NVarChar).Value = operation;
                commonInterface.ExecuteNonQuery(Copient.DataBases.LogixRT);
                return commonInterface.RowsAffected > 0;
            } // end addToImportQueue()

            public bool addToImportFileQueue(Copient.CommonInc commonInterface, int eid)
            {
                // add the new file
                commonInterface.QueryStr = string.Format(@"
                    INSERT INTO CRMFileImportQueue ( CRMImportQueuePKID, filename, CRMStatusFlag, LastUpdate, Deleted )
                        SELECT pkid as CRMImportQueuePKID, '{1}' as filename, 0 as CRMStatusFlag, GETDATE() as LastUpdate, 0 as Deleted
                        FROM CRMImportQueue
                        WHERE OfferID = {0} AND Deleted=0
                ", m_offerID, importFileName());
                commonInterface.DBParameters.Add("@OfferId", SqlDbType.BigInt).Value = m_offerID;

                //commonInterface.DBParameters.Add("", SqlDbType.BigInt).Value = m_offerID;

                commonInterface.ExecuteNonQuery(Copient.DataBases.LogixRT);
                return commonInterface.RowsAffected > 0;
            }


        } // end class OfferDefinition

        public OfferDefinition[] m_offers;  // map an extOfferID to a definition of the offer



        // select the files that are ready to be downloaded
        private const string m_offerToExportQueryFormat = @"
            SELECT CPE_Incentives.IncentiveID as OfferID
	            , CPE_incentives.ClientOfferID as extofferid
	            , CRMFileExportQueue.FileName as filename
	            , CPE_Incentives.EngineID as EngineID
            FROM CRMFileExportQueue WITH ( NoLock )
                INNER JOIN crmexportqueue ON crmfileexportqueue.crmexportqueuePKID = crmexportqueue.PKID
                INNER JOIN CPE_Incentives ON crmexportqueue.OfferID = CPE_Incentives.IncentiveID
			        AND crmexportqueue.ExtInterfaceID = {0}
                LEFT JOIN ExtSegmentMap ex ON ex.IncentiveId = CPE_Incentives.IncentiveID
            WHERE CRMFileExportQueue.deleted = 'false' AND crmstatusflag = 0 AND (ex.ExtSegmentID > 0 or ex.ExtSegmentID IS NULL)
            UNION
            SELECT Offers.OfferID as OfferID
	            , Offers.ExtOfferID as extofferid
	            , CRMFileExportQueue.FileName as filename
	            , Offers.EngineID as EngineID
	            FROM CRMFileExportQueue WITH ( NoLock )
                INNER JOIN crmexportqueue ON crmfileexportqueue.crmexportqueuePKID = crmexportqueue.PKID
                INNER JOIN Offers ON crmexportqueue.OfferID = Offers.OfferID
			        AND crmexportqueue.ExtInterfaceID = {0}
            WHERE CRMFileExportQueue.deleted = 'false' AND crmstatusflag = 0
        ";

        public static string offerFilesToExportQuery()
        {
            return string.Format(m_offerToExportQueryFormat, "@ExtInterfaceID");
        }


        public static DataTable offersToExport(ExternalInterfaceID eid, Copient.CommonInc commonInterface)
        {
            commonInterface.Open_LogixRT();
            commonInterface.QueryStr = offerFilesToExportQuery();
            commonInterface.DBParameters.Add("@ExtInterfaceID", SqlDbType.Int).Value = eid.m_extInterfaceID;
            return commonInterface.ExecuteQuery(Copient.DataBases.LogixRT);
        } // end method offersToExport()


        public Offers()
        {
            m_offers = new OfferDefinition[0];
        }

        /* Find all the offers that have been exported for the given External Interface ID */
        public Offers(ExternalInterfaceID eid, Copient.CommonInc commonInterface)
        {
            DataTable offersToExportTable = Offers.offersToExport(eid, commonInterface);
            int numOffers = offersToExportTable.Rows.Count;
            m_offers = new OfferDefinition[numOffers]; // offersToExportTable has Offerid, extofferid, filename
            if (numOffers < 1)
            {
                return;
            }

            string exportedOffersPath = exportFilePath(commonInterface);
            for (int i = 0; i < numOffers; ++i)
            {
                DataRow anOfferRow = offersToExportTable.Rows[i];
                if (anOfferRow != null)
                {
                    OfferDefinition od = new OfferDefinition(anOfferRow, exportedOffersPath);
                    m_offers[i] = od;
                }
            } // end for i

        } // end Offers ctor()




        public bool import(ExternalInterfaceID eid, logger l, Copient.CommonInc commonInterface, ref Dictionary<string, long> offerids, ref Dictionary<long, long> CGIDs)
        {

            if (m_offers == null)
                return false;


            string importFileTargetPath = importFilePath(commonInterface);
            string sXml;
            bool bstatus;
            bool bcstatus;


            // foreach offer in the list
            foreach (OfferDefinition anOffer in m_offers)
            {
                string operation = "add";
                if (anOffer == null || !anOffer.hasExternalID()) continue; // skip blank entries or entries we can't do anything with

                sXml = anOffer.m_offerXML.decode();
                if (sXml.Contains("Engine=\"CM\""))
                {
                    anOffer.m_engineID = (int)Copient.CommonInc.InstalledEngines.CM;
                }
                else if (sXml.Contains("3</EngineID>"))
                {
                    anOffer.m_engineID = (int)Copient.CommonInc.InstalledEngines.Website;
                }
                else if (sXml.Contains("9</EngineID>"))
                {
                    anOffer.m_engineID = (int)Copient.CommonInc.InstalledEngines.UE;
                }
                else
                {
                    anOffer.m_engineID = (int)Copient.CommonInc.InstalledEngines.CPE;
                }

                //* The connector determines if the offer exists in CPE_Incentives.                
                if (!anOffer.offerExists(commonInterface, eid.m_extInterfaceID))
                {   //* If it doesn't exist, it creates a stub CPE_Incentive record and obtains the record's IncentiveID.
                    anOffer.addNewSkeletonOffer(commonInterface, eid.m_extInterfaceID);
                }
                else
                    operation = "edit";

                //Check if document contains ChargebackDepCode but is missing ChargebackDeptID tag or contains ClientOfferID in the header
                //    which overwrites extOfferID in the database or has ExtBannerID defined under /Offer/Banners/Banner.
                if (((sXml.Contains("ClientOfferID")) || (sXml.Contains("ChargebackDeptCode"))) || (sXml.Contains("ExtBannerID")) || (anOffer.m_engineID == 3))
                {
                    XmlDocument xmlDoc = new XmlDocument();
                    XmlNodeList CGNodeList = default(XmlNodeList);
                    XmlNodeList ConditionNodeList = default(XmlNodeList);
                    XmlDocument offerDoc = new XmlDocument();
                    offerDoc.LoadXml(sXml);

                    XmlNode nodeClientOfferId = offerDoc.SelectSingleNode("/Offer/Header/ClientOfferID");
                    XmlNode hdrNode = offerDoc.SelectSingleNode("/Offer/Header");
                    XmlNode nodeChargebackDeptCode = offerDoc.SelectSingleNode("/Offer/Rewards/Discount/ChargebackDeptCode");
                    XmlNode nodeChargebackDeptID = offerDoc.SelectSingleNode("/Offer/Rewards/Discount/ChargebackDeptID");
                    XmlNode nodeDiscount = offerDoc.SelectSingleNode("/Offer/Rewards/Discount");
                    XmlNode nodeBannersExtBannerID = offerDoc.SelectSingleNode("/Offer/Banners/Banner/ExtBannerID");
                    XmlNode nodeBannersBannerID = offerDoc.SelectSingleNode("/Offer/Banners/Banner/BannerID");
                    XmlNode hdrnodeBannersBanner = offerDoc.SelectSingleNode("/Offer/Banners/Banner");

                    XmlNode nodeConditions = offerDoc.SelectSingleNode("/Offer/Conditions");
                    XmlNode nodeRewards = offerDoc.SelectSingleNode("/Offer/Rewards");
                    XmlNode nodeAuxilary = offerDoc.SelectSingleNode("/Offer/Auxilary");

                    if (nodeClientOfferId != null)
                    {
                        hdrNode.RemoveChild(nodeClientOfferId);
                    }

                    if (nodeChargebackDeptCode != null)
                    {
                        string chargebackDeptCode = nodeChargebackDeptCode.InnerText;
                        string bannerID = "0";
                        if (nodeBannersExtBannerID != null)
                        {
                            bannerID = nodeBannersExtBannerID.InnerText;
                        }
                        else if (nodeBannersBannerID != null)
                        {
                            bannerID = nodeBannersBannerID.InnerText;
                        }

                        commonInterface.Open_LogixRT();
                        if ((nodeBannersExtBannerID != null) && (!(bannerID.Equals("0"))))
                        {
                            commonInterface.QueryStr = "SELECT ChargeBackDepts.ChargeBackDeptID, ChargeBackDepts.BannerID from ChargeBackDepts INNER JOIN Banners ON ChargeBackDepts.BannerID=Banners.BannerID WHERE ChargeBackDepts.ExternalID=@chargebackDeptCode and Banners.ExtBannerID=@bannerID ";
                            commonInterface.DBParameters.Add("@chargebackDeptCode", SqlDbType.NVarChar).Value = chargebackDeptCode;
                            commonInterface.DBParameters.Add("@bannerID", SqlDbType.NVarChar).Value = bannerID;
                        }
                        else
                        {
                            commonInterface.QueryStr = "SELECT ChargeBackDeptID from ChargeBackDepts WHERE ExternalID= @chargebackDeptCode  and BannerID =@bannerID";
                            commonInterface.DBParameters.Add("@chargebackDeptCode", SqlDbType.NVarChar).Value = chargebackDeptCode;
                             commonInterface.DBParameters.Add("@bannerID", SqlDbType.Int).Value = bannerID;
                        }

                        DataTable matchingChargebackDeptCount = commonInterface.LRT_Select();
                        if (matchingChargebackDeptCount.Rows.Count == 1)
                        {
                            int chargebackDeptID = ((int)matchingChargebackDeptCount.Rows[0]["ChargeBackDeptID"]);
                            if (nodeChargebackDeptID != null)
                            {
                                nodeChargebackDeptID.InnerText = chargebackDeptID.ToString();
                            }
                            else
                            {
                                XmlElement newNodeChargebackDeptID = offerDoc.CreateElement("ChargebackDeptID");
                                newNodeChargebackDeptID.SetAttribute("DataType", "Int32");
                                newNodeChargebackDeptID.InnerText = chargebackDeptID.ToString();
                                nodeDiscount.PrependChild(newNodeChargebackDeptID);
                            }

                            if ((nodeBannersExtBannerID != null) && (!(bannerID.Equals("0"))))
                            {
                                bannerID = matchingChargebackDeptCount.Rows[0]["BannerID"].ToString();
                                if (nodeBannersBannerID != null)
                                {
                                    nodeBannersBannerID.InnerText = bannerID;
                                }
                                else
                                {
                                    XmlElement newNodeBannersBannerID = offerDoc.CreateElement("BannerID");
                                    newNodeBannersBannerID.SetAttribute("DataType", "Int32");
                                    newNodeBannersBannerID.InnerText = bannerID;
                                    hdrnodeBannersBanner.PrependChild(newNodeBannersBannerID);
                                }
                            }
                        }
                    }

                    if (nodeBannersExtBannerID != null)
                    {
                        hdrnodeBannersBanner.RemoveChild(nodeBannersExtBannerID);
                    }

                    if (anOffer.m_engineID == 3)
                    {
                        List<XmlNode> nodestoremove = new List<XmlNode>();


                        if (nodeConditions != null)
                        {
                            foreach (XmlNode child in nodeConditions)
                            {
                                if ((child.Name != "Customer") && (child.Name != "Point"))
                                {
                                    nodestoremove.Add(child);
                                }
                            }
                        }

                        if (nodeRewards != null)
                        {
                            foreach (XmlNode child in nodeRewards)
                            {
                                if (child.Name != "Membership")
                                {
                                    nodestoremove.Add(child);
                                }
                            }
                        }

                        if (nodeAuxilary != null)
                        {
                            foreach (XmlNode child in nodeAuxilary)
                            {
                                if ((child.Name != "CustomerGroup") && (child.Name != "PointsProgram"))
                                {
                                    nodestoremove.Add(child);
                                }
                            }
                        }

                        foreach (XmlNode node in nodestoremove)
                        {
                            node.ParentNode.RemoveChild(node);
                        }
                    }

                    //process customer conditions first to get the customer group ids 
                    if ((commonInterface.Fetch_SystemOption(246) == "1") && (anOffer.m_engineID == 3 || anOffer.m_engineID == 2))
                    {
                        if ((nodeAuxilary != null))
                        {
                            xmlDoc.LoadXml(nodeAuxilary.OuterXml);

                            CGNodeList = xmlDoc.GetElementsByTagName("CustomerGroup");
                            if ((CGNodeList.Count > 0))
                            {

                                bstatus = ProcessCustomerGroups(CGNodeList, commonInterface);

                                if (bstatus)
                                {

                                    ConditionNodeList = nodeConditions.SelectNodes("Customer");

                                    bcstatus = ProcessCustomerConditions(ConditionNodeList, commonInterface, anOffer.m_offerID, ref CGIDs);

                                }
                            }
                        }
                    }

                    sXml = offerDoc.OuterXml;
                    anOffer.m_offerXML.m_encoded_xml = System.Web.HttpUtility.HtmlDecode(sXml);
                }

                l.log(string.Format("Importing offer with ExtInterfaceID( {0} ), ExtID( {1} ), OfferID( {2} ), EngineID( {3} )", eid.m_extInterfaceID, anOffer.m_extOfferID, anOffer.m_offerID, anOffer.m_engineID));
                if (anOffer.writeToFile(importFileTargetPath) && anOffer.associateExternalIDToOffer(commonInterface))
                {
                    bool m_outboundEnabled = eid.isOutboundEnabled(commonInterface); // this specifies if the CRMEngineID used for outbound is the same value as Inbound CRMEngineID
                    anOffer.updateOfferCRMEngineID(commonInterface, eid.m_extInterfaceID, m_outboundEnabled);
                    anOffer.addToImportQueue(commonInterface, eid.m_extInterfaceID,operation);
                    anOffer.addToImportFileQueue(commonInterface, eid.m_extInterfaceID);
                }

                offerids.Add(anOffer.m_extOfferID, anOffer.m_offerID);

            } // end foreach

            return true;
        }
        public bool import(ExternalInterfaceID eid, logger l, Copient.CommonInc commonInterface)
        {

            if (m_offers == null)
                return false;

            string importFileTargetPath = importFilePath(commonInterface);
            string sXml;
            int  extInterfaceType = 0;

            // foreach offer in the list
            foreach (OfferDefinition anOffer in m_offers)
            {
                string operation = "add";
                if (anOffer == null || !anOffer.hasExternalID()) continue; // skip blank entries or entries we can't do anything with
                sXml = anOffer.m_offerXML.decode();
                if (sXml.Contains("Engine=\"CM\""))
                {
                    anOffer.m_engineID = (int)Copient.CommonInc.InstalledEngines.CM;
                    // ensure that xml contains value for 'ExtOfferID' based on that in the Offer Definition, thus avoiding a discrepancy
                    sXml = CM_SetElementInOfferHeader("ExtOfferID", anOffer.m_extOfferID.ToString(), true, sXml);
                    // ensure that xml contains value for 'InboundCRMEngineID' based on the Interface ID, thus avoiding a discrepancy
                    int iEffectiveInboundId = eid.m_extInterfaceID;
                    sXml = CM_SetXmlInboundId(ref iEffectiveInboundId, sXml, commonInterface, false, ref extInterfaceType);
                    //anOffer.m_offerXML.m_encoded_xml = System.Web.HttpUtility.HtmlEncode(sXml);
                    anOffer.m_offerXML.m_encoded_xml = sXml;
                }
                else
                {
                    if (sXml.Contains("9</EngineID>"))
                    { //UE offer
                        anOffer.m_engineID = (int)Copient.CommonInc.InstalledEngines.UE;
                        if (sXml.Contains("ExternalBuyerId"))
                        {
                            //string xpath = "Offer/Header/ExternalBuyerId";
                            System.Xml.Linq.XDocument doc = System.Xml.Linq.XDocument.Parse(sXml.Trim());
                            var node = from ele in doc.Descendants("ExternalBuyerId") select new { ExternalBuyerId = (string)ele.Value };

                            //Get the  Buyer ID from externalBuyerID
                            if (node != null && node.ToList().Count > 0)
                                anOffer.m_buyerId = GetBuyerID(node.ToList()[0].ExternalBuyerId, commonInterface);
                        }
                    }
                    else
                    {
                        anOffer.m_engineID = (int)Copient.CommonInc.InstalledEngines.CPE;
                    }
                }

                //* The connector determines if the offer exists in CPE_Incentives.
                if (!anOffer.offerExists(commonInterface, eid.m_extInterfaceID))
                {   //* If it doesn't exist, it creates a stub CPE_Incentive record and obtains the record's IncentiveID.

                    //To avoid concurrent beging with creating external id, if creation failed that means offer already created, 
                    //recheck offfer existance to get logix offer id.
                    if (!anOffer.InsertExternalID(commonInterface, anOffer.m_extOfferID, eid.m_extInterfaceID))
                    {
                        anOffer.offerExists(commonInterface, eid.m_extInterfaceID);
                    }
                    anOffer.addNewSkeletonOffer(commonInterface, eid.m_extInterfaceID);
                }
                else
                    operation = "edit";

                // At this point, the offer already existed or has been added as a skeleton.  Either way, if the ext interface is of type '6'
                // (import and translate), update the field in the offer to make sure it gets translated.  Note that this only applies to offers
                // with engine=0 (cm) but the variable extInterfaceType, after being initted to 0, gets set only if engine is 0 anyhow
                if (extInterfaceType == 6)
                {
                    update_offer_translation_column(commonInterface, l, anOffer.m_offerID);
                }

                //Check if document contains ChargebackDepCode but is missing ChargebackDeptID tag or contains ClientOfferID in the header
                //    which overwrites extOfferID in the database or has ExtBannerID defined under /Offer/Banners/Banner.
                if (((sXml.Contains("ClientOfferID")) || (sXml.Contains("ChargebackDeptCode"))) || (sXml.Contains("ExtBannerID")))
                {
                    XmlDocument offerDoc = new XmlDocument();
                    offerDoc.LoadXml(sXml);

                    XmlNode nodeClientOfferId = offerDoc.SelectSingleNode("/Offer/Header/ClientOfferID");
                    XmlNode hdrNode = offerDoc.SelectSingleNode("/Offer/Header");
                    XmlNode nodeChargebackDeptCode = offerDoc.SelectSingleNode("/Offer/Rewards/Discount/ChargebackDeptCode");
                    XmlNode nodeChargebackDeptID = offerDoc.SelectSingleNode("/Offer/Rewards/Discount/ChargebackDeptID");
                    XmlNode nodeDiscount = offerDoc.SelectSingleNode("/Offer/Rewards/Discount");
                    XmlNode nodeBannersExtBannerID = offerDoc.SelectSingleNode("/Offer/Banners/Banner/ExtBannerID");
                    XmlNode nodeBannersBannerID = offerDoc.SelectSingleNode("/Offer/Banners/Banner/BannerID");
                    XmlNode hdrnodeBannersBanner = offerDoc.SelectSingleNode("/Offer/Banners/Banner");

                    if (nodeClientOfferId != null)
                    {
                        hdrNode.RemoveChild(nodeClientOfferId);
                    }

                    if (nodeChargebackDeptCode != null)
                    {
                        string chargebackDeptCode = nodeChargebackDeptCode.InnerText;
                        string bannerID = "0";
                        if (nodeBannersExtBannerID != null)
                        {
                            bannerID = nodeBannersExtBannerID.InnerText;
                        }
                        else if (nodeBannersBannerID != null)
                        {
                            bannerID = nodeBannersBannerID.InnerText;
                        }

                        commonInterface.Open_LogixRT();
                        if ((nodeBannersExtBannerID != null) && (!(bannerID.Equals("0"))))
                        {
                            commonInterface.QueryStr = "SELECT ChargeBackDepts.ChargeBackDeptID, ChargeBackDepts.BannerID from ChargeBackDepts INNER JOIN Banners ON ChargeBackDepts.BannerID=Banners.BannerID WHERE ChargeBackDepts.ExternalID=@ChargebackDeptCode and Banners.ExtBannerID=@BannerID";
                            commonInterface.DBParameters.Add("@BannerID", SqlDbType.NVarChar).Value = bannerID;
                        }
                        else
                        {
                            commonInterface.QueryStr = "SELECT ChargeBackDeptID from ChargeBackDepts WHERE ExternalID=@ChargebackDeptCode and BannerID=@BannerID";
                            commonInterface.DBParameters.Add("@BannerID", SqlDbType.Int).Value = Convert.ToInt32(bannerID);
                        }
                        commonInterface.DBParameters.Add("@ChargebackDeptCode", SqlDbType.NVarChar).Value = chargebackDeptCode;
                        DataTable matchingChargebackDeptCount = commonInterface.ExecuteQuery(Copient.DataBases.LogixRT);
                        if (matchingChargebackDeptCount.Rows.Count == 1)
                        {
                            int chargebackDeptID = ((int)matchingChargebackDeptCount.Rows[0]["ChargeBackDeptID"]);
                            if (nodeChargebackDeptID != null)
                            {
                                nodeChargebackDeptID.InnerText = chargebackDeptID.ToString();
                            }
                            else
                            {
                                XmlElement newNodeChargebackDeptID = offerDoc.CreateElement("ChargebackDeptID");
                                newNodeChargebackDeptID.SetAttribute("DataType", "Int32");
                                newNodeChargebackDeptID.InnerText = chargebackDeptID.ToString();
                                nodeDiscount.PrependChild(newNodeChargebackDeptID);
                            }

                            if ((nodeBannersExtBannerID != null) && (!(bannerID.Equals("0"))))
                            {
                                bannerID = matchingChargebackDeptCount.Rows[0]["BannerID"].ToString();
                                if (nodeBannersBannerID != null)
                                {
                                    nodeBannersBannerID.InnerText = bannerID;
                                }
                                else
                                {
                                    XmlElement newNodeBannersBannerID = offerDoc.CreateElement("BannerID");
                                    newNodeBannersBannerID.SetAttribute("DataType", "Int32");
                                    newNodeBannersBannerID.InnerText = bannerID;
                                    hdrnodeBannersBanner.PrependChild(newNodeBannersBannerID);
                                }
                            }
                        }
                    }

                    if (nodeBannersExtBannerID != null)
                    {
                        hdrnodeBannersBanner.RemoveChild(nodeBannersExtBannerID);
                    }

                    sXml = offerDoc.OuterXml;
                    //anOffer.m_offerXML.m_encoded_xml = System.Web.HttpUtility.HtmlEncode(sXml);
                }

                l.log(string.Format("Importing offer with ExtInterfaceID( {0} ), ExtOfferID( {1} ), OfferID( {2} ), EngineID( {3} )", eid.m_extInterfaceID, anOffer.m_extOfferID, anOffer.m_offerID, anOffer.m_engineID));
                if (anOffer.writeToFile(importFileTargetPath) && anOffer.associateExternalIDToOffer(commonInterface))
                {
                    bool m_outboundEnabled = eid.isOutboundEnabled(commonInterface); // this specifies if the CRMEngineID used for outbound is the same value as Inbound CRMEngineID
                    anOffer.updateOfferCRMEngineID(commonInterface, eid.m_extInterfaceID, m_outboundEnabled);
                    anOffer.addToImportQueue(commonInterface, eid.m_extInterfaceID,operation);
                    anOffer.addToImportFileQueue(commonInterface, eid.m_extInterfaceID);
                }

            } // end foreach
            return true;
        } // end method import

        private bool ProcessCustomerConditions(XmlNodeList nodeList, Copient.CommonInc commonInterface, long offerid, ref Dictionary<long, long> CGIds)
        {
            System.Collections.IEnumerator nodeEnum = default(IEnumerator);
            XmlNode node = default(XmlNode);
            long ROID = 0;
            long oldCG = 0;
            long newCG = 0;
            StringBuilder sqlBuf = new StringBuilder();
            Hashtable ExceptionTable = new Hashtable(1);
            Hashtable ReplacementTable = new Hashtable(4);
            string HHEnable = "";
            bool bstatus = true;

            try
            {

                if ((nodeList != null) && (nodeList.Count > 0))
                {
                    ROID = GetParentROID(commonInterface, offerid);
                    if ((ROID > 0))
                    {
                        nodeEnum = nodeList.GetEnumerator();

                        //Delete any existing records
                        commonInterface.QueryStr = "update CPE_IncentiveCustomerGroups set Deleted=1 where RewardOptionID=" + ROID + ";";
                        commonInterface.LRT_Execute();

                        while ((nodeEnum.MoveNext()))
                        {

                            node = (XmlNode)nodeEnum.Current;
                            if (node.ChildNodes.Count > 0)
                            {
                                oldCG = long.Parse(GetChildNodeValue(node, "CustomerGroupID"));
                                newCG = Convert.ToInt64(LookupByKey("CG", oldCG));
                                ExceptionTable.Clear();
                                ExceptionTable.Add("IncentiveCustomerID", "IncentiveCustomerID");
                                ExceptionTable.Add("HHEnable", "HHEnable");
                                ReplacementTable.Clear();
                                ReplacementTable.Add("RewardOptionID", ROID);
                                ReplacementTable.Add("CustomerGroupID", newCG);
                                ReplacementTable.Add("Deleted", 0);
                                ReplacementTable.Add("LastUpdate", System.DateTime.Now.ToString());

                                sqlBuf = new StringBuilder("insert into CPE_IncentiveCustomerGroups ");
                                sqlBuf.Append(GenerateInsertSQL(node, ExceptionTable, ReplacementTable, commonInterface));
                                commonInterface.QueryStr = sqlBuf.ToString();
                                commonInterface.LRT_Execute();

                                //Update the ROID for Householding enabled
                                HHEnable = GetChildNodeValue(node, "HHEnable");
                                if ((!string.IsNullOrEmpty(HHEnable)))
                                {
                                    commonInterface.QueryStr = "update CPE_RewardOptions set HHEnable = " + (HHEnable.ToUpper() == "TRUE" ? 1 : 0) + " where RewardOptionID=" + ROID;
                                    commonInterface.LRT_Execute();
                                }
                            }

                            CGIds.Add(offerid, newCG);
                        }
                    }
                }
            }
            catch (Exception ex)
            {

                bstatus = false;

            }


            return bstatus;
        }

        private string GenerateInsertSQL(XmlNode xmlNode, Hashtable ExceptionTable, Hashtable ReplacementTable, Copient.CommonInc commonInterface)
        {
            StringBuilder sqlBuf = new StringBuilder();
            XmlNode node = null;
            string DataType = "";
            int i = 0;
            int nodeCt = 0;
            object NodeValue = null;


            if (((xmlNode != null) && xmlNode.HasChildNodes))
            {
                nodeCt = xmlNode.ChildNodes.Count;
                ArrayList FieldList = new ArrayList(nodeCt);
                ArrayList ValueList = new ArrayList(nodeCt);
                ArrayList DataTypeList = new ArrayList(nodeCt);

                // first store the fields, values and datatypes for each tag
                for (i = 0; i <= (nodeCt - 1); i++)
                {
                    node = xmlNode.ChildNodes[i];
                    if ((((node != null)) && (!ExceptionTable.ContainsKey(node.Name))))
                    {
                        FieldList.Add(node.Name);
                        DataType = node.Attributes["DataType"].InnerText;
                        DataTypeList.Add(DataType);

                        // used for primary key replacement
                        if ((ReplacementTable.ContainsKey(node.Name)))
                        {
                            NodeValue = ReplacementTable[(node.Name)];
                        }
                        else
                        {
                            NodeValue = node.InnerText;
                        }
                        ValueList.Add(NodeValue);
                    }
                }

                // then add any replacement tags that weren't in the xml already
                if ((ReplacementTable != null))
                {
                    foreach (DictionaryEntry RTable in ReplacementTable)
                    {
                        if (!(FieldList.Contains(RTable.Key)))
                        {
                            FieldList.Add(RTable.Key);
                            DataTypeList.Add("STRING");
                            ValueList.Add(RTable.Value);
                        }
                    }
                }

                // now construct the statement based on the lists
                if ((FieldList.Count > 0))
                {
                    sqlBuf.Append("(");
                    for (i = 0; i <= FieldList.Count - 1; i++)
                    {
                        sqlBuf.Append(FieldList[(i)]);
                        if ((i < FieldList.Count - 1))
                        {
                            sqlBuf.Append(",");
                        }
                    }
                    sqlBuf.Append(") values (");
                    for (i = 0; i <= ValueList.Count - 1; i++)
                    {
                        switch (DataTypeList[i].ToString().ToUpper())
                        {
                            case "DATETIME":
                                sqlBuf.Append("convert(DATETIME,'" + ValueList[i].ToString() + "',121)");
                                break;
                            case "STRING":
                                if (DataTypeList[i].ToString().ToUpper() == "STRING")
                                    sqlBuf.Append("N");
                                sqlBuf.Append("'");
                                sqlBuf.Append(commonInterface.Parse_Quotes(ValueList[i].ToString()));
                                sqlBuf.Append("'");
                                break;
                            case "BOOLEAN":
                                sqlBuf.Append((ValueList[i].ToString().ToUpper() == "TRUE" ? 1 : 0));
                                break;
                            default:
                                sqlBuf.Append(ValueList[i].ToString());
                                break;
                        }
                        if ((i < ValueList.Count - 1))
                        {
                            sqlBuf.Append(",");
                        }
                    }
                    sqlBuf.Append(");");
                }
            }

            return sqlBuf.ToString();
        }



        private object LookupByKey(string KeyPrefix, object OldValue)
        {
            object obj = OldValue;

            if ((NamingTable.ContainsKey(KeyPrefix + OldValue.ToString())))
            {
                obj = NamingTable[(KeyPrefix + OldValue.ToString())];
            }

            return obj;
        }

        private string GetChildNodeValue(XmlNode ParentNode, string ChildNodeName)
        {
            string Value = "";
            XmlNode ChildNode = default(XmlNode);

            if ((ParentNode != null))
            {
                ChildNode = ParentNode.SelectSingleNode(ChildNodeName);
                if ((ChildNode != null))
                {
                    Value = ChildNode.InnerText;
                }
            }

            return CPEStrFilter(Value);
        }

        private long GetParentROID(Copient.CommonInc commonInterface, long offerid)
        {
            DataTable dt = default(DataTable);
            long ROID = -1;
            try
            {
                if ((ROID == -1))
                {
                    // get the roid for this offer
                    commonInterface.QueryStr = "select RewardOptionID from CPE_RewardOptions Where TouchResponse=0 and IncentiveId = " + offerid;
                    dt = commonInterface.LRT_Select();
                    if ((dt.Rows.Count > 0))
                    {
                        ROID = (long)dt.Rows[0]["RewardOptionID"];
                    }
                }
            }
            catch (Exception ex)
            {
                ROID = -1;
            }

            return ROID;
        }


        private bool ProcessCustomerGroups(XmlNodeList nodeList, Copient.CommonInc commonInterface)
        {
            bool functionReturnValue = false;
            System.Collections.IEnumerator nodeEnum = default(System.Collections.IEnumerator);
            XmlNode cgNode = default(XmlNode);
            XmlNode hdrNode = default(XmlNode);
            XmlNode node = default(XmlNode);
            long oldCGID = -1;
            long newCGID = 0;
            long newCustPK = 0;
            long newMembershipID = 0;
            string CustomerList = "";
            string FileName = "";
            string SQLBulkPath = "";
            string NewCgName = "";
            long AllCAMGroup = 0;
            long NewCardholdersID = 0;
            Copient.CAM MyCam = new Copient.CAM();
            bool IsAnyCAM = false;
            bool IsNewCardholder = false;
            bool IsOptInGroup = false;
            bool bstatus = true;
            const long ADMIN_USER_ID = 1;
            const int LANGUAGE_ID = 1;

            try
            {
                //sErrorSubMethod = "(ProcessCustomerGroups)";
                if ((commonInterface.LRTadoConn.State != ConnectionState.Open))
                {
                    commonInterface.Open_LogixRT();
                }
                if ((commonInterface.LXSadoConn.State != ConnectionState.Open))
                {
                    commonInterface.Open_LogixXS();
                }

                SQLBulkPath = commonInterface.Fetch_SystemOption(29);
                if (string.IsNullOrEmpty(SQLBulkPath))
                {
                    return functionReturnValue;
                    //MyCommon.Error_Processor(, , , , AppID)
                }

                if (!(SQLBulkPath.Substring(SQLBulkPath.Length - 1, 1) == "\\"))
                {
                    SQLBulkPath = SQLBulkPath + "\\";
                }

                AllCAMGroup = MyCam.GetAllCAMCustomerGroupID();
                NewCardholdersID = GetNewCardholdersID(commonInterface);

                nodeEnum = nodeList.GetEnumerator();
                cgNode = null;
                hdrNode = null;
                while ((nodeEnum.MoveNext()))
                {
                    // First create a new CustomerGroup from the XML Header Tag    
                    cgNode = (XmlNode)nodeEnum.Current;

                    hdrNode = cgNode.FirstChild;

                    // do not import  anycustomer(1) or anycardholder(2) groups
                    oldCGID = long.Parse(hdrNode.SelectSingleNode("CustomerGroupID").InnerText);
                    if ((hdrNode.SelectSingleNode("AnyCAMCardholder") != null))
                    {
                        bool.TryParse(hdrNode.SelectSingleNode("AnyCAMCardholder").InnerXml, out IsAnyCAM);
                    }
                    if ((hdrNode.SelectSingleNode("NewCardholders") != null))
                    {
                        bool.TryParse(hdrNode.SelectSingleNode("NewCardholders").InnerXml, out IsNewCardholder);
                    }

                    if ((IsAnyCAM && AllCAMGroup > 0))
                    {
                        StoreNodeValue(hdrNode, "CustomerGroupID", "CG", AllCAMGroup.ToString());
                    }
                    else if ((IsNewCardholder && NewCardholdersID > 0))
                    {
                        StoreNodeValue(hdrNode, "CustomerGroupID", "CG", NewCardholdersID.ToString());
                    }
                    if ((oldCGID >= 3))
                    {
                        if ((hdrNode.SelectSingleNode("Name") == null))
                        {
                            NewCgName = "Unknown!";
                        }
                        else
                        {
                            NewCgName = CPEStrFilter(hdrNode.SelectSingleNode("Name").InnerXml);
                        }
                        if ((hdrNode.SelectSingleNode("IsOptInGroup") == null))
                        {
                            IsOptInGroup = false;
                        }
                        else
                        {
                            bool.TryParse(hdrNode.SelectSingleNode("IsOptInGroup").InnerXml, out IsOptInGroup);
                        }

                        commonInterface.QueryStr = "dbo.pt_CustomerGroups_Insert";
                        commonInterface.Open_LRTsp();
                        commonInterface.LRTsp.Parameters.Add("@Name", SqlDbType.NVarChar, 200).Value = NewCgName;
                        commonInterface.LRTsp.Parameters.Add("@CustomerGroupID", SqlDbType.BigInt).Direction = ParameterDirection.Output;
                        commonInterface.LRTsp.Parameters.Add("@CAMCustomerGroup", SqlDbType.Bit).Value = 0;
                        commonInterface.LRTsp.Parameters.Add("@IsOptInGroup", SqlDbType.Bit).Value = IsOptInGroup;
                        commonInterface.LRTsp.ExecuteNonQuery();
                        newCGID = (long)commonInterface.LRTsp.Parameters["@CustomerGroupID"].Value;
                        commonInterface.Close_LRTsp();
                        StoreNodeValue(hdrNode, "CustomerGroupID", "CG", newCGID.ToString());

                        // log history for this group
                        commonInterface.Activity_Log(4, newCGID, ADMIN_USER_ID, Copient.PhraseLib.Lookup("history.cgroup-create", LANGUAGE_ID));

                        // Then create the customers associated to this customer group
                        node = hdrNode.NextSibling;
                        if ((node != null))
                        {
                            CustomerList = node.InnerText;
                            FileName = "CG" + newCGID + ".txt";
                            CustomerList = CustomerList.Replace("\n", "\r\n");
                            WriteListToFile(CustomerList, SQLBulkPath + FileName);
                            WriteToCustomerQueue(FileName, newCGID, "ProcessCustomerGroups", commonInterface);
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                bstatus = false;
            }

            return bstatus;
        }

        private void WriteListToFile(string IdList, string FileName)
        {
            FileStream fStream = null;
            System.Text.ASCIIEncoding encoding = new System.Text.ASCIIEncoding();
            byte[] b = null;

            try
            {
                fStream = File.Create(FileName);
                b = encoding.GetBytes(IdList);
                fStream.Write(b, 0, b.Length);

            }
            catch (Exception ex)
            {
            }
            finally
            {
                if ((fStream != null))
                {
                    fStream.Flush();
                    fStream.Close();
                    fStream.Dispose();
                    fStream = null;
                }
            }
        }

        private void WriteToCustomerQueue(string FileName, long CustomerGroupID, string ServiceName, Copient.CommonInc commonInterface)
        {
            try
            {

                if ((commonInterface.LXSadoConn.State != ConnectionState.Open))
                {
                    commonInterface.Open_LogixXS();
                }
                //set the date for yesterday so this will be the top of the Queue
                commonInterface.QueryStr = "Insert into GMInsertQueue (FileName, UploadTime, CustomerGroupID, CardTypeID, StatusFlag) " + "values (N'" + commonInterface.Parse_Quotes(FileName) + "', '" + System.DateTime.Now.AddDays(-1) + "', " + CustomerGroupID + "," + commonInterface.Fetch_SystemOption(30) + ", 0);";

                commonInterface.LXS_Execute();

            }
            catch (Exception ex)
            {
                throw ex;
            }

        }


        private bool StoreNodeValue(XmlNode ParentNode, string TagName, string KeyPrefix, string NewValue)
        {
            XmlNode OldIdNode = default(XmlNode);
            bool Stored = false;

            try
            {

                OldIdNode = ParentNode.SelectSingleNode(TagName);
                if ((OldIdNode != null) & (NewValue != null))
                {
                    if ((NamingTable.ContainsKey(KeyPrefix + OldIdNode.InnerXml) && NewValue != "-1"))
                    {
                        NamingTable[(KeyPrefix + OldIdNode.InnerXml)] = NewValue;
                    }
                    else if ((NewValue != "-1"))
                    {
                        NamingTable.Add(KeyPrefix + OldIdNode.InnerXml, NewValue);
                    }
                    else if ((NewValue == "-1"))
                    {
                        NamingTable.Add(KeyPrefix + OldIdNode.InnerXml, OldIdNode.InnerXml);
                    }
                    Stored = true;
                }
            }
            catch (Exception ex)
            {

                Stored = false;

            }


            return Stored;
        }


        private string CPEStrFilter(string a)
        {
            StringBuilder builder = new StringBuilder();
            foreach (char i in a)
            {
                if (i == Convert.ToChar(10))
                {
                    builder.Append(" ");
                }
                else if (i == Convert.ToChar(13))
                {
                    // Squelch it
                }
                else if (i < Convert.ToChar(32))
                {
                    builder.Append(" ");
                }
                else
                {
                    builder.Append(i);
                }
            }
            return builder.ToString();
        }

        private long GetNewCardholdersID(Copient.CommonInc commonInterface)
        {
            long CustomerGroupID = 0;
            DataTable dt = default(DataTable);

            try
            {
                if (commonInterface.LRTadoConn.State == ConnectionState.Closed)
                    commonInterface.Open_LogixRT();

                commonInterface.QueryStr = "select CustomerGroupID from CustomerGroups with (NoLock) where Deleted=0 and NewCardholders=1;";
                dt = commonInterface.LRT_Select();
                if (dt.Rows.Count > 0)
                {
                    CustomerGroupID = (long)commonInterface.NZ(dt.Rows[0]["CustomerGroupID"], 0);

                }
            }
            catch (Exception ex)
            {
                CustomerGroupID = -1;
            }

            return CustomerGroupID;
        }


        public OfferImmediateStatuses import_immediate(Copient.CommonInc commonInterface, logger l, ExternalInterfaceID eid, Offers os)
        {

            if (m_offers == null)
            {
                return null;
            }
            if (m_offers.GetLength(0) < 1)
            {
                return null;
            }
            Copient.ImportXml m_ImportXml = new Copient.ImportXml(ref commonInterface, false);
            OfferImmediateStatuses m_OfferImmediateStatuses = new OfferImmediateStatuses(commonInterface, eid, os);
            string sXml;
            int i = 0;
            bool bImportOk, bAutoDeployed, bAutoSendOutbound;
            int extInterfaceType = 0;

            // foreach offer in the list
            foreach (OfferDefinition anOffer in m_offers)
            {
                string operation = "add";
                m_OfferImmediateStatuses.m_offerImmediateStatuses[i].m_offerId = 0;
                m_OfferImmediateStatuses.m_offerImmediateStatuses[i].m_extOfferId = anOffer.m_extOfferID;
                if (anOffer == null || !anOffer.hasExternalID())
                {
                    anOffer.m_engineID = (int)Copient.CommonInc.InstalledEngines.CM;
                    m_OfferImmediateStatuses.m_offerImmediateStatuses[i].m_status = 1;
                    m_OfferImmediateStatuses.m_offerImmediateStatuses[i].m_status_msg = string.Format("Offer in slot ({0}) of Offer array is null or no ExternalId is defined!", i);
                    l.log(m_OfferImmediateStatuses.m_offerImmediateStatuses[i].m_status_msg);
                }
                else
                {
                    sXml = anOffer.m_offerXML.decode();
                    if (sXml.Contains("Engine=\"CM\""))
                    {
                        sXml = CM_SetElementInOfferHeader("ExtOfferID", anOffer.m_extOfferID.ToString(), true, sXml);
                        int iEffectiveInboundId = eid.m_extInterfaceID;
                        sXml = CM_SetXmlInboundId(ref iEffectiveInboundId, sXml, commonInterface, true, ref extInterfaceType);
                        anOffer.m_engineID = (int)Copient.CommonInc.InstalledEngines.CM;
                        bImportOk = m_ImportXml.ImportOfferCRM("", sXml, 1, 1);
                        if (bImportOk)
                        {
                            // get the Offer ID (populate anOffer.m_offerID)
                            if (anOffer.offerExists(commonInterface, iEffectiveInboundId))
                                operation = "edit";


                            // Add to import queue, since this is where UI gets CRM last received info
                            // CRM Import will set the "Deleted" flag since there is no corresponding entry in CRMFileImportQueue
                            anOffer.addToImportQueue(commonInterface, eid.m_extInterfaceID,operation);

                            m_OfferImmediateStatuses.m_offerImmediateStatuses[i].m_offerId = anOffer.m_offerID;
                            m_OfferImmediateStatuses.m_offerImmediateStatuses[i].m_status = 0;
                            bAutoDeployed = CM_AutoDeploy(anOffer.m_offerID, eid.m_extInterfaceID, commonInterface);
                            if (commonInterface.Fetch_CM_SystemOption(114) == "1")
                            {
                                bAutoSendOutbound = CM_AutoSendOutbound(anOffer.m_offerID, eid.m_extInterfaceID, commonInterface);
                            }
                            if (bAutoDeployed)
                            {
                                m_OfferImmediateStatuses.m_offerImmediateStatuses[i].m_status_msg = string.Format("Imported and deployed offer with ExtInterfaceID( {0} ), ExtOfferID( {1} ), OfferID( {2} ), EngineID( {3} )", eid.m_extInterfaceID, anOffer.m_extOfferID, anOffer.m_offerID, anOffer.m_engineID);
                            }
                            else
                            {
                                m_OfferImmediateStatuses.m_offerImmediateStatuses[i].m_status_msg = string.Format("Imported offer with ExtInterfaceID( {0} ), ExtOfferID( {1} ), OfferID( {2} ), EngineID( {3} )", eid.m_extInterfaceID, anOffer.m_extOfferID, anOffer.m_offerID, anOffer.m_engineID);
                            }
                            if (extInterfaceType == 6)
                            {
                                update_offer_translation_column(commonInterface, l, anOffer.m_offerID);
                            }
                        }
                        else
                        {
                            anOffer.offerExists(commonInterface, iEffectiveInboundId);
                            m_OfferImmediateStatuses.m_offerImmediateStatuses[i].m_offerId = anOffer.m_offerID;
                            m_OfferImmediateStatuses.m_offerImmediateStatuses[i].m_status = 1;
                            m_OfferImmediateStatuses.m_offerImmediateStatuses[i].m_status_msg = m_ImportXml.GetErrorMsg();
                            if (m_OfferImmediateStatuses.m_offerImmediateStatuses[i].m_status_msg.Length == 0)
                            {
                                m_OfferImmediateStatuses.m_offerImmediateStatuses[i].m_status_msg = m_ImportXml.GetStatusMsg();
                            }
                        }
                        l.log(m_OfferImmediateStatuses.m_offerImmediateStatuses[i].m_status_msg);
                    }
                    else
                    {
                        if (sXml.Contains("EngineID>9")) { anOffer.m_engineID = (int)Copient.CommonInc.InstalledEngines.UE; }
                        else { anOffer.m_engineID = (int)Copient.CommonInc.InstalledEngines.CPE; }
                        m_OfferImmediateStatuses.m_offerImmediateStatuses[i].m_status = 1;
                        m_OfferImmediateStatuses.m_offerImmediateStatuses[i].m_status_msg = string.Format("Offer with ID( {0} )failed because offers for EngineID( {1} ) can not be imported via this method!", anOffer.m_extOfferID, anOffer.m_engineID);
                        l.log(m_OfferImmediateStatuses.m_offerImmediateStatuses[i].m_status_msg);
                    }
                }
                i++;
            } // end foreach
            return m_OfferImmediateStatuses;
        } // end method import_immediate

        private void update_offer_translation_column(Copient.CommonInc commonInterface, logger l, long lOfferId)
        {
            commonInterface.QueryStr = "update Offers set AutoTranslateToUE=1 where Deleted=0 and OfferID=" + lOfferId + ";";
            commonInterface.LRT_Execute();
            l.log("Updated CM Offer '" + lOfferId + "' to be translated to UE");
        }

        public OfferImmediateStatuses transfer_cm_immediate(Copient.CommonInc commonInterface, logger l, ExternalInterfaceID eid, Offers os)
        {
            if (m_offers == null)
            {
                return null;
            }
            if (m_offers.GetLength(0) < 1)
            {
                return null;
            }
            Copient.ImportXml m_ImportXml = new Copient.ImportXml(ref commonInterface, false);
            OfferImmediateStatuses m_OfferImmediateStatuses = new OfferImmediateStatuses(commonInterface, eid, os);
            string sXml;
            int i = 0;
            bool bImportOk, bAutoDeployed;

            // foreach offer in the list
            foreach (OfferDefinition anOffer in m_offers)
            {
                m_OfferImmediateStatuses.m_offerImmediateStatuses[i].m_offerId = 0;
                m_OfferImmediateStatuses.m_offerImmediateStatuses[i].m_extOfferId = anOffer.m_extOfferID;
                if (anOffer == null)
                {
                    anOffer.m_engineID = (int)Copient.CommonInc.InstalledEngines.CM;
                    m_OfferImmediateStatuses.m_offerImmediateStatuses[i].m_status = 1;
                    m_OfferImmediateStatuses.m_offerImmediateStatuses[i].m_status_msg = string.Format("Offer in slot ({0}) of Offer array is null!", i);
                    l.log(m_OfferImmediateStatuses.m_offerImmediateStatuses[i].m_status_msg);
                }
                else
                {
                    sXml = anOffer.m_offerXML.decode();
                    if (sXml.Contains("Engine=\"CM\""))
                    {
                        anOffer.m_engineID = (int)Copient.CommonInc.InstalledEngines.CM;
                        bImportOk = m_ImportXml.ImportOfferTransfer(sXml, 1, 1, eid.m_extInterfaceID);
                        if (bImportOk)
                        {
                            // get the Offer ID (populate anOffer.m_offerID)
                            anOffer.transferCmOfferExists(commonInterface);

                            m_OfferImmediateStatuses.m_offerImmediateStatuses[i].m_offerId = anOffer.m_offerID;
                            m_OfferImmediateStatuses.m_offerImmediateStatuses[i].m_status = 0;
                            bAutoDeployed = CM_AutoDeploy(anOffer.m_offerID, eid.m_extInterfaceID, commonInterface);
                            if (bAutoDeployed)
                            {
                                m_OfferImmediateStatuses.m_offerImmediateStatuses[i].m_status_msg = string.Format("Imported and deployed offer with ExtInterfaceID( {0} ), ExtOfferID( {1} ), OfferID( {2} ), EngineID( {3} )", eid.m_extInterfaceID, anOffer.m_extOfferID, anOffer.m_offerID, anOffer.m_engineID);
                            }
                            else
                            {
                                m_OfferImmediateStatuses.m_offerImmediateStatuses[i].m_status_msg = string.Format("Imported offer with ExtInterfaceID( {0} ), ExtOfferID( {1} ), OfferID( {2} ), EngineID( {3} )", eid.m_extInterfaceID, anOffer.m_extOfferID, anOffer.m_offerID, anOffer.m_engineID);
                            }
                        }
                        else
                        {
                            anOffer.transferCmOfferExists(commonInterface);
                            m_OfferImmediateStatuses.m_offerImmediateStatuses[i].m_offerId = anOffer.m_offerID;
                            m_OfferImmediateStatuses.m_offerImmediateStatuses[i].m_status = 1;
                            m_OfferImmediateStatuses.m_offerImmediateStatuses[i].m_status_msg = m_ImportXml.GetErrorMsg();
                            if (m_OfferImmediateStatuses.m_offerImmediateStatuses[i].m_status_msg.Length == 0)
                            {
                                m_OfferImmediateStatuses.m_offerImmediateStatuses[i].m_status_msg = m_ImportXml.GetStatusMsg();
                            }
                        }
                        l.log(m_OfferImmediateStatuses.m_offerImmediateStatuses[i].m_status_msg);
                    }
                    else
                    {
                        if (sXml.Contains("EngineID>9")) { anOffer.m_engineID = (int)Copient.CommonInc.InstalledEngines.UE; }
                        else { anOffer.m_engineID = (int)Copient.CommonInc.InstalledEngines.CPE; }
                        m_OfferImmediateStatuses.m_offerImmediateStatuses[i].m_status = 1;
                        m_OfferImmediateStatuses.m_offerImmediateStatuses[i].m_status_msg = string.Format("Offer with ID( {0} )failed because offers for EngineID( {1} ) can not be imported via this method!", anOffer.m_extOfferID, anOffer.m_engineID);
                        l.log(m_OfferImmediateStatuses.m_offerImmediateStatuses[i].m_status_msg);
                    }
                }
                i++;
            } // end foreach
            return m_OfferImmediateStatuses;
        } // end method transfer_cm_immediate

        public OfferTranslateStatuses translate_cm_to_ue(Copient.CommonInc commonInterface, logger l, ExternalInterfaceID eid, Offers os, bool bEditPreviousLocationGroupAssignments)
        {
            if (m_offers == null)
            {
                return null;
            }
            if (m_offers.GetLength(0) < 1)
            {
                return null;
            }
            Copient.ImportXml m_ImportXml = new Copient.ImportXml(ref commonInterface, false);
            OfferTranslateStatuses m_OfferTranslateStatuses = new OfferTranslateStatuses(commonInterface, eid, os);
            string sXml;
            string sXmlUe;
            string sMsg;
            string sMsgTemp;
            int i = 0;
            int iStatus;
            bool bStatus;
            bool bAutoDeployed;
            long lOfferId = 0;

            // foreach offer in the list
            foreach (OfferDefinition anOffer in m_offers)
            {
                if (anOffer == null)
                {
                    //anOffer.m_engineID = (int)Copient.CommonInc.InstalledEngines.UE;
                    m_OfferTranslateStatuses.m_offerTranslateStatuses[i].m_status = 1;
                    sMsg = string.Format("Offer in slot ({0}) of Offer array is null!", i);
                }
                else
                {
                    m_OfferTranslateStatuses.m_offerTranslateStatuses[i].m_ueOfferId = 0;
                    m_OfferTranslateStatuses.m_offerTranslateStatuses[i].m_cmOfferId = 0;
                    sXml = anOffer.m_offerXML.decode();
                    if (sXml.Contains("PromoMaint"))
                    {
                        sMsgTemp = "";
                        anOffer.m_engineID = (int)Copient.CommonInc.InstalledEngines.UE;

                        iStatus = m_ImportXml.TranslateCmToUe(sXml, 1, 1, eid.m_extInterfaceID, anOffer.m_locations, bEditPreviousLocationGroupAssignments);
                        l.log("Status: " + iStatus.ToString());
                        anOffer.m_extOfferID = m_ImportXml.GetOfferId();
                        if (anOffer.m_extOfferID == null)
                        {
                            anOffer.m_extOfferID = "*";
                        }

                        if (!long.TryParse(anOffer.m_extOfferID, out lOfferId)) lOfferId = 0;
                        anOffer.translateUeOfferExists(eid.m_extInterfaceID, commonInterface);
                        m_OfferTranslateStatuses.m_offerTranslateStatuses[i].m_cmOfferId = lOfferId;
                        m_OfferTranslateStatuses.m_offerTranslateStatuses[i].m_ueOfferId = anOffer.m_offerID;

                        if (iStatus > 0)
                        {
                            sMsg = m_ImportXml.GetErrorMsg();
                            if (sMsg.Length == 0)
                            {
                                sMsg = m_ImportXml.GetStatusMsg();
                            }
                        }
                        else
                        {
                            //  0 -> translated ok; -1 -> translated with warning;
                            // -3 -> unassignined locations ok; -4 -> unassigned locations with warning
                            if (iStatus == 0 || iStatus == -1 || iStatus == -3 || iStatus == -4)
                            {
                                if (iStatus == -1 || iStatus == -4)
                                {
                                    sMsgTemp = m_ImportXml.GetErrorMsg();
                                    if (sMsgTemp.Length == 0)
                                    {
                                        sMsgTemp = m_ImportXml.GetStatusMsg();
                                    }
                                }
                                if (iStatus == 0 || iStatus == -1)
                                {
                                    sXmlUe = m_ImportXml.GetTransLatedUeXml();
                                    CMS.AMS.CurrentRequest.Resolver.AppName = "CRMOfferConnector.asmx";
                                    Copient.ImportXMLUE m_ImportXmlUe = CMS.AMS.CurrentRequest.Resolver.Resolve<Copient.ImportXMLUE>();
                                    bStatus = m_ImportXmlUe.ImportTranslatedOffer("", sXmlUe, 1, 1, anOffer.m_locations, bEditPreviousLocationGroupAssignments);
                                    if (bStatus)
                                    {
                                        // get the Offer ID (populate anOffer.m_offerID)
                                        if (anOffer.translateUeOfferExists(eid.m_extInterfaceID, commonInterface))
                                        {
                                            m_OfferTranslateStatuses.m_offerTranslateStatuses[i].m_ueOfferId = anOffer.m_offerID;
                                            bAutoDeployed = UE_AutoDeploy(anOffer.m_offerID, eid.m_extInterfaceID, commonInterface);
                                            if (bAutoDeployed)
                                            {
                                                sMsg = string.Format("Translated CM Offer '{0}' and deployed UE Offer '{1}'.", anOffer.m_extOfferID, anOffer.m_offerID);
                                            }
                                            else
                                            {
                                                sMsg = string.Format("Translated CM Offer '{0}' to UE Offer '{1}'.", anOffer.m_extOfferID, anOffer.m_offerID);
                                            }
                                            if (anOffer.m_locations.Length > 0)
                                            {
                                                l.log(string.Format("Locations: {0}", anOffer.m_locations));
                                            }

                                            UE_SetupDistributionMigration(anOffer.m_extOfferID, anOffer.m_offerID, commonInterface, l);
                                        }
                                        else
                                        {
                                            iStatus = 1;
                                            sMsg = string.Format("Translated CM Offer '{0}' to UE Offer, but UE Offer does not exist in DB.", anOffer.m_extOfferID);
                                        }
                                    }
                                    else
                                    {
                                        iStatus = 1;
                                        sMsg = m_ImportXmlUe.GetErrorMsg();
                                        if (sMsg.Length == 0)
                                        {
                                            sMsg = m_ImportXmlUe.GetStatusMsg();
                                        }
                                    }
                                }
                                else
                                {
                                    // Un-assign locations
                                    // get the Offer ID (populate anOffer.m_offerID)
                                    if (anOffer.translateUeOfferExists(eid.m_extInterfaceID, commonInterface))
                                    {
                                        m_OfferTranslateStatuses.m_offerTranslateStatuses[i].m_ueOfferId = anOffer.m_offerID;
                                        bAutoDeployed = UE_AutoDeploy(anOffer.m_offerID, eid.m_extInterfaceID, commonInterface);
                                        if (bAutoDeployed)
                                        {
                                            sMsg = string.Format("Un-assigned locations for CM Offer '{0}' and re-deployed UE Offer '{1}'.", anOffer.m_extOfferID, anOffer.m_offerID);
                                        }
                                        else
                                        {
                                            sMsg = string.Format("Un-assigned locations for CM Offer '{0}' (UE Offer '{1}').", anOffer.m_extOfferID, anOffer.m_offerID);
                                        }
                                        if (anOffer.m_locations.Length > 0)
                                        {
                                            l.log(string.Format("Locations: {0}", anOffer.m_locations));
                                        }
                                    }
                                    else
                                    {
                                        iStatus = 1;
                                        sMsg = string.Format("Un-assigned locations for CM Offer '{0}' to UE Offer, but UE Offer does not exist in DB.", anOffer.m_extOfferID);
                                    }
                                }
                            }
                            else
                            {
                                sMsg = m_ImportXml.GetErrorMsg();
                                if (sMsg.Length == 0)
                                {
                                    sMsg = m_ImportXml.GetStatusMsg();
                                }
                            }
                        }
                        m_OfferTranslateStatuses.m_offerTranslateStatuses[i].m_status = iStatus;
                        if (sMsgTemp.Length > 0)
                        {
                            sMsg = sMsg + " - " + sMsgTemp;
                        }
                    }
                    else
                    {
                        m_OfferTranslateStatuses.m_offerTranslateStatuses[i].m_status = 1;
                        sMsg = string.Format("Invalid XML in slot ({0}) of Offer array!", i);
                    }
                }
                m_OfferTranslateStatuses.m_offerTranslateStatuses[i].m_status_msg = sMsg;
                l.log(sMsg);
                i++;
            } // end foreach
            return m_OfferTranslateStatuses;
        } // end method transfer_cm_to_ue

        public OfferTranslateStatuses delete_translated_offer(Copient.CommonInc commonInterface, logger l, ExternalInterfaceID eid, Offers os)
        {
            if (m_offers == null)
            {
                return null;
            }
            if (m_offers.GetLength(0) < 1)
            {
                return null;
            }
            Copient.ImportXml m_ImportXml = new Copient.ImportXml(ref commonInterface, false);
            OfferTranslateStatuses m_OfferTranslateStatuses = new OfferTranslateStatuses(commonInterface, eid, os);
            string sMsg;
            int i = 0;
            int iStatus;
            DataTable dt;
            long UeOfferId = 0;
            long CmOfferId = 0;

            // foreach offer in the list
            foreach (OfferDefinition anOffer in m_offers)
            {
                sMsg = "Unknown Error!";
                if (anOffer == null)
                {
                    m_OfferTranslateStatuses.m_offerTranslateStatuses[i].m_status = 1;
                    sMsg = string.Format("Offer in slot ({0}) of Offer array is null!", i);
                }
                else
                {
                    switch (anOffer.m_engineID)
                    {
                        case (int)Copient.CommonInc.InstalledEngines.UE:
                            if (anOffer.m_offerID > 0)
                            {
                                commonInterface.QueryStr = "select IncentiveID, ClientOfferID from CPE_Incentives with (NoLock) where IncentiveID = @IncentiveID and InboundCRMEngineID = @InboundCRMEngineID and Deleted=0";
                                commonInterface.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = anOffer.m_offerID;
                                commonInterface.DBParameters.Add("@InboundCRMEngineID", SqlDbType.Int).Value = eid.m_extInterfaceID;
                                dt = commonInterface.ExecuteQuery(Copient.DataBases.LogixRT);
                                if (dt.Rows.Count > 0)
                                {
                                    UeOfferId = anOffer.m_offerID;
                                    DataRow dr = dt.Rows[0];
                                    CmOfferId = int.Parse(dr["ClientOfferID"].ToString());
                                }
                                else
                                {
                                    UeOfferId = 0;
                                    CmOfferId = 0;
                                    sMsg = string.Format("UE Offer with IncentiveID '{0}' and External Source '{1}' does not exist!", anOffer.m_offerID, eid.m_extInterfaceID);
                                }
                            }
                            else
                            {
                                UeOfferId = 0;
                                CmOfferId = 0;
                                sMsg = string.Format("Offer in slot ({0}) of Offer array has no Offer ID specified!", i);
                            }
                            break;
                        case (int)Copient.CommonInc.InstalledEngines.CM:
                            if (anOffer.m_offerID > 0)
                            {
                                commonInterface.QueryStr = "select IncentiveID, ClientOfferID from CPE_Incentives with (NoLock) where ClientOfferID = @ClientOfferID and InboundCRMEngineID = @InboundCRMEngineID and Deleted=0";
                                commonInterface.DBParameters.Add("@ClientOfferID", SqlDbType.NVarChar).Value = anOffer.m_offerID.ToString();
                                commonInterface.DBParameters.Add("@InboundCRMEngineID", SqlDbType.Int).Value = eid.m_extInterfaceID;
                                dt = commonInterface.ExecuteQuery(Copient.DataBases.LogixRT);
                                if (dt.Rows.Count > 0)
                                {
                                    CmOfferId = anOffer.m_offerID;
                                    DataRow dr = dt.Rows[0];
                                    UeOfferId = int.Parse(dr["IncentiveID"].ToString());
                                }
                                else
                                {
                                    UeOfferId = 0;
                                    CmOfferId = 0;
                                    sMsg = string.Format("UE Offer with ClientOfferID (CM OfferID) '{0}' and External Source '{1}' does not exist!", anOffer.m_offerID, eid.m_extInterfaceID);
                                }
                            }
                            else
                            {
                                UeOfferId = 0;
                                CmOfferId = 0;
                                sMsg = string.Format("Offer in slot ({0}) of Offer array has no Offer ID specified!", i);
                            }
                            break;
                        default:
                            UeOfferId = 0;
                            CmOfferId = 0;
                            sMsg = string.Format("Offer in slot ({0}) of Offer array has invalid Engine ID ({1}) (0 => CM; 9 => UE)", i, anOffer.m_engineID);
                            break;
                    }

                    if (UeOfferId > 0)
                    {
                        commonInterface.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=2,Deleted=1,UpdateLevel=UpdateLevel+1,LastUpdate=getdate() where IncentiveID=@IncentiveID;";
                        commonInterface.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = UeOfferId;
                        commonInterface.ExecuteNonQuery(Copient.DataBases.LogixRT);
                        if (commonInterface.RowsAffected > 0)
                        {
                            // Mark the shadow table offer as deleted as well.
                            commonInterface.QueryStr = "update CPE_ST_Incentives with (RowLock) set Deleted=1, UpdateLevel = (select UpdateLevel from CPE_Incentives with (NoLock) where IncentiveID = @IncentiveID) where IncentiveID = @IncentiveID";
                            commonInterface.DBParameters.Add("@IncentiveID", SqlDbType.BigInt).Value = UeOfferId;
                            commonInterface.ExecuteNonQuery(Copient.DataBases.LogixRT);

                            // remove the banners assigned to this offer
                            if (commonInterface.Fetch_SystemOption(66) == "1")
                            {
                                commonInterface.QueryStr = "delete from BannerOffers with (RowLock) where OfferID = @OfferID";
                                commonInterface.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = UeOfferId;
                                commonInterface.ExecuteNonQuery(Copient.DataBases.LogixRT);
                            }
                            m_OfferTranslateStatuses.m_offerTranslateStatuses[i].m_status = 0;
                            sMsg = string.Format("Deleted UE Offer (UE Offer ID = '{0}')  (CM Offer ID = '{1}')", UeOfferId, CmOfferId);
                            commonInterface.Activity_Log(3, UeOfferId, 1, Copient.PhraseLib.Lookup("history.offer-delete", 1));
                        }
                        else
                        {
                            m_OfferTranslateStatuses.m_offerTranslateStatuses[i].m_status = 1;
                            sMsg = string.Format("Update of 'Deleted' flag for UE Offer (UE Offer ID = '{0}')  (CM Offer ID = '{1}') failed!", UeOfferId, CmOfferId);
                        }
                    }
                    else
                    {
                        m_OfferTranslateStatuses.m_offerTranslateStatuses[i].m_status = 1;
                    }
                }
                m_OfferTranslateStatuses.m_offerTranslateStatuses[i].m_cmOfferId = CmOfferId;
                m_OfferTranslateStatuses.m_offerTranslateStatuses[i].m_ueOfferId = UeOfferId;
                m_OfferTranslateStatuses.m_offerTranslateStatuses[i].m_status_msg = sMsg;
                l.log(sMsg);
                i++;
            } // end foreach
            return m_OfferTranslateStatuses;
        } // end method transfer_cm_to_ue


        private string CM_SetXmlInboundId(ref int iInBoundId, string sXML, Copient.CommonInc commonInterface, bool bAllowInternal, ref int extInterfaceType)
        {
            bool bAddElement = true;
            string sNewXml = "";
            int iSrcType;

            extInterfaceType = 0;
            if (sXML != null)
            {
                commonInterface.QueryStr = "select ExtInterfaceTypeID from ExtCRMInterfaces with (NoLock) where ExtInterfaceID=@ExtInterfaceID;";
                commonInterface.DBParameters.Add("@ExtInterfaceID", SqlDbType.Int).Value = iInBoundId;
                DataTable dt = commonInterface.ExecuteQuery(Copient.DataBases.LogixRT);
                if (dt.Rows.Count > 0)
                {
                    DataRow dr = dt.Rows[0];
                    iSrcType = int.Parse(dr["ExtInterfaceTypeID"].ToString());
                    extInterfaceType = iSrcType;
                    switch (iSrcType)
                    {
                        case 0:
                            break;
                        case 2:
                            if (bAllowInternal)
                            {
                                iInBoundId = 0;
                                bAddElement = false;
                            }
                            else
                            {
                                // this external source should not be posting offers via CRM!
                                throw new CRMOfferConnectorException(string.Format("External Interface ID '{0}' has invalid External Interface Type ID '{1}' for posting offers via CRM Import agent!", iInBoundId, iSrcType));
                            }
                            break;
                        // Allow import of type 6, which imports and translates
                        case 6:
                            break;

                        default:
                            // this external source should not be posting offers via CRM!
                            throw new CRMOfferConnectorException(string.Format("External Interface ID '{0}' has invalid External Interface Type ID '{1}' for posting offers!", iInBoundId, iSrcType));
                    }
                    sNewXml = CM_SetElementInOfferHeader("InboundCRMEngineID", iInBoundId.ToString(), bAddElement, sXML);
                }
                else
                {
                    // how did this source pass validation?
                    throw new CRMOfferConnectorException(string.Format("External Interface ID '{0}' is either invalid or inactive!", iInBoundId));
                }
            }
            return sNewXml;
        }

        private string CM_SetElementInOfferHeader(string sName, string sValue, bool bAddElement, string sXML)
        {
            int startPos, endPos;
            // allow for DataType attribute in element
            string startTag = "<" + sName;
            string endTag = "</" + sName + ">";
            string sNewXml = "";

            if (sXML != null)
            {
                startPos = sXML.IndexOf(startTag);
                if (startPos >= 0)
                {
                    startPos += startTag.Length;
                    // allow for DataType attribute in element
                    startPos = sXML.IndexOf(">", startPos) + 1;
                    endPos = sXML.IndexOf(endTag, startPos);
                    if (endPos > startPos)
                    {
                        sNewXml = sXML.Substring(0, startPos) + sValue + sXML.Substring(endPos);
                    }
                    else
                    {
                        // invalid XML - start element, but no end element
                        throw new CRMOfferConnectorException(string.Format("The tag '{0}' is not present in the XML!", endTag));
                    }
                }
                else
                {
                    if (bAddElement)
                    {
                        // insert new element row at end of header data
                        startPos = sXML.IndexOf("</Header>");
                        if (startPos >= 0)
                        {
                            sNewXml = sXML.Substring(0, startPos) + startTag + ">" + sValue + endTag + sXML.Substring(startPos);
                        }
                        else
                        {
                            // invalid XML - must have Header element!
                            throw new CRMOfferConnectorException(string.Format("The tag </Header> is not in the XML!"));
                        }
                    }
                    else
                    {
                        sNewXml = sXML;
                    }
                }
            }
            return sNewXml;
        }

        private bool CM_AutoDeploy(long lOfferID, int iExtInterfaceID, Copient.CommonInc commonInterface)
        {
            DataTable dt;
            bool bAutoDeploy = false;

            commonInterface.QueryStr = "select AutoDeploy from ExtCRMInterfaces with (NoLock) where ExtInterfaceID=@ExtInterfaceID and Deleted=0;";
            commonInterface.DBParameters.Add("@ExtInterfaceID", SqlDbType.Int).Value = iExtInterfaceID;
            dt = commonInterface.ExecuteQuery(Copient.DataBases.LogixRT);
            if (dt.Rows.Count > 0)
            {
                DataRow dr = dt.Rows[0];
                bAutoDeploy = bool.Parse(dr["AutoDeploy"].ToString());
                if (bAutoDeploy)
                {
                    commonInterface.QueryStr = "update Offers with (RowLock) set StatusFlag=2,UpdateLevel=UpdateLevel+1,CRMAutoDeployed=1 where OfferID=@OfferID;";
                    commonInterface.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = lOfferID;
                    commonInterface.ExecuteNonQuery(Copient.DataBases.LogixRT);
                    if (commonInterface.RowsAffected > 0)
                    {
                        commonInterface.Activity_Log(3, lOfferID, 1, Copient.PhraseLib.Lookup("history.offer-deploy", 1));
                    }
                }
            }
            return bAutoDeploy;
        }

        private bool UE_AutoDeploy(long lOfferID, int iExtInterfaceID, Copient.CommonInc commonInterface)
        {
            DataTable dt;
            bool bAutoDeploy = false;

            commonInterface.QueryStr = "select AutoDeploy from ExtCRMInterfaces with (NoLock) where ExtInterfaceID=@ExtInterfaceID and Deleted=0;";
            commonInterface.DBParameters.Add("@ExtInterfaceID", SqlDbType.Int).Value = iExtInterfaceID;
            dt = commonInterface.ExecuteQuery(Copient.DataBases.LogixRT);
            if (dt.Rows.Count > 0)
            {
                DataRow dr = dt.Rows[0];
                bAutoDeploy = bool.Parse(dr["AutoDeploy"].ToString());
                if (bAutoDeploy)
                {
                    commonInterface.QueryStr = "update CPE_Incentives with (RowLock) set StatusFlag=2,UpdateLevel=UpdateLevel+1, LastDeployClick=getdate(), DeployDeferred=0, LastUpdate=getdate(), LastUpdatedByAdminID=1 where IncentiveID=@OfferID and Deleted=0;";
                    commonInterface.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = lOfferID;
                    commonInterface.ExecuteNonQuery(Copient.DataBases.LogixRT);
                    if (commonInterface.RowsAffected > 0)
                    {
                        commonInterface.Activity_Log(3, lOfferID, 1, Copient.PhraseLib.Lookup("history.offer-deploy", 1));
                    }
                }
            }
            return bAutoDeploy;
        }

        private bool CM_AutoSendOutbound(long lOfferID, int iExtInterfaceID, Copient.CommonInc commonInterface)
        {
            DataTable dt;
            bool bAutoSendOutbound = false;

            commonInterface.QueryStr = "select AutoSendOutbound from ExtCRMInterfaces with (NoLock) where ExtInterfaceID=@ExtInterfaceID and OutboundEnabled=1 and Deleted=0;";
            commonInterface.DBParameters.Add("@ExtInterfaceID", SqlDbType.Int).Value = iExtInterfaceID;
            dt = commonInterface.ExecuteQuery(Copient.DataBases.LogixRT);
            if (dt.Rows.Count > 0)
            {
                DataRow dr = dt.Rows[0];
                bAutoSendOutbound = Convert.ToBoolean(commonInterface.NZ(dr["AutoSendOutbound"], false));
                if (bAutoSendOutbound)
                {
                    commonInterface.QueryStr = "update Offers with (RowLock) set LastCRMSendDate=getdate(),CRMEngineUpdateLevel=CRMEngineUpdateLevel+1,CRMSendToExport=1,CRMSendStatus=1 where OfferID=@OfferID;";
                    commonInterface.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = lOfferID;
                    commonInterface.ExecuteNonQuery(Copient.DataBases.LogixRT);
                    if (commonInterface.RowsAffected > 0)
                    {
                        commonInterface.Activity_Log(3, lOfferID, 1, Copient.PhraseLib.Lookup("history.offer-sendoutbound", 1));
                    }
                }
            }
            return bAutoSendOutbound;
        }

        private int GetBuyerID(string ExternalBuyerId, Copient.CommonInc commonInterface)
        {
            DataTable dt;
            int BuyerID = -1;
            commonInterface.QueryStr = "select BuyerID from Buyers with (NoLock) where ExternalBuyerId=@ExternalBuyerId;";
            commonInterface.DBParameters.Add("@ExternalBuyerId", SqlDbType.VarChar, 60).Value = ExternalBuyerId;
            dt = commonInterface.ExecuteQuery(Copient.DataBases.LogixRT);
            if (dt.Rows.Count > 0)
            {
                BuyerID = int.Parse(dt.Rows[0][0].ToString());
            }
            return BuyerID;
        }

        private void UE_SetupDistributionMigration(string sOfferID, long lIncentiveID, Copient.CommonInc commonInterface, logger l)
        {
            commonInterface.QueryStr = "insert into RewardDistributionMigration (CMOfferID, UEIncentiveID, RewardOptionID, StatusDate, StatusFlag, NumDistributionsMigrated) select @CMOfferID, @UEIncentiveID, RewardOptionID, GetDate(), 1, 0  from cpe_rewardoptions where incentiveid = @UEIncentiveID";
            commonInterface.DBParameters.Add("@CMOfferID", SqlDbType.NVarChar).Value = sOfferID;
            commonInterface.DBParameters.Add("@UEIncentiveID", SqlDbType.BigInt).Value = lIncentiveID;
            commonInterface.ExecuteNonQuery(Copient.DataBases.LogixRT);
            if (commonInterface.RowsAffected > 0)
            {
                l.log("UE_SetupDistributionMigration(Added migration data for offer " + sOfferID + ")");
            }
            else
            {
                l.log("UE_SetupDistributionMigration(Error: failed to add migration data for offer " + sOfferID + ")");
            }
        }

    } // end class Offers

    private string WriteOfferXmlResponse(Dictionary<string, long> offerids, string ErrorMsg, bool ResponseFlag, Dictionary<long, long> CGIds)
    {
        StringWriter sw = null;
        XmlWriterSettings Settings = default(XmlWriterSettings);
        XmlWriter Writer = default(XmlWriter);
        string xmlStr = "";


        sw = new StringWriter();
        Settings = new XmlWriterSettings();
        Settings.Encoding = Encoding.UTF8;
        Settings.Indent = true;
        Settings.IndentChars = ControlCharacters.Tab;

        Settings.NewLineChars = ControlCharacters.CrLf;
        Settings.NewLineHandling = NewLineHandling.Replace;

        Writer = XmlWriter.Create(sw, Settings);

        Writer.WriteStartDocument();
        Writer.WriteStartElement("CRMOfferConnector");


        foreach (KeyValuePair<string, long> offer in offerids)
        {
            Writer.WriteStartElement("Offer");
            Writer.WriteAttributeString("ClientOfferId", (offer.Key));
            Writer.WriteAttributeString("LogixOfferID", offer.Value == 0 ? "-1" : offer.Value.ToString());
            foreach (KeyValuePair<long, long> cgroup in CGIds)
            {
                if (cgroup.Key == offer.Value)
                {
                    Writer.WriteAttributeString("CustomerGroupId", (cgroup.Value.ToString()));
                }
            }
            Writer.WriteAttributeString("operation", "PostOffersEnhancedReponse");
            Writer.WriteAttributeString("success", ResponseFlag.ToString().ToLower());
            Writer.WriteEndElement();
        }


        //Offer

        if ((!string.IsNullOrEmpty(ErrorMsg.Trim())))
        {
            Writer.WriteStartElement("Error");
            Writer.WriteAttributeString("message", ErrorMsg);
            Writer.WriteEndElement();
            //Error
        }

        Writer.WriteEndElement();
        //ExternalOfferConnector
        Writer.WriteEndDocument();

        Writer.Flush();
        Writer.Close();

        // workaround for problem where encoding is always set to utf-16 no matter
        // what you set for the encoding in the XMLWriterSettings.Encoding 
        xmlStr = sw.ToString();

        if ((xmlStr != null))
        {
            xmlStr = xmlStr.Replace("encoding=\"utf-16\"", "encoding=\"utf-8\"");
        }

        return xmlStr;
    }
    [WebMethod] //------------------------------------------------------------------------------------------------------------------   
    public string postOffersEnhancedResponse(ExternalInterfaceID ei, Offers os)
    {
        // accept a set of offers from the external interface and queue them for importation
        Dictionary<string, long> offerids = new Dictionary<string, long>();
        Dictionary<long, long> cgids = new Dictionary<long, long>();
        log(string.Format("postOffersEnhancedResponse( {0} )", ei));
        ei.verify(m_commonRoutines);
        if (os == null)
            return WriteOfferXmlResponse(offerids, "No Offers to import", false, cgids);

        if (os.import(ei, m_logger, m_commonRoutines, ref offerids, ref cgids) == true)

            return WriteOfferXmlResponse(offerids, "", true, cgids);

        else
            return WriteOfferXmlResponse(offerids, "An error occured. Please see CRM Offer Connetor log", false, cgids);
    }

    [WebMethod] //------------------------------------------------------------------------------------------------------------------
    public bool postOffers(ExternalInterfaceID ei, Offers os)
    {
        bool importStatus;

        // accept a set of offers from the external interface and queue them for importation
        log(string.Format("postOffers( {0} )", ei));
        if (os == null)
        {
            log(string.Format("postOffers(Error: null value sent for Offers)"));
            return false;
        }

        //filename validation
        foreach (Offers.OfferDefinition anOffer in os.m_offers)
        {
            if (anOffer.m_filename != null && anOffer.m_filename != "")
            {
                //then validate
                if(anOffer.m_filename.IndexOfAny(invalidFileNameChars) != -1)
                {
                    //log the error
                    log(string.Format("postOffers(Error: Filename is invalid and is not allowed)"));
                    //return false
                    return false;
                }
            }
        }
        // This fails silently without generating any notification in import() where it continues past any offers not having external OfferIDs, not processing them,
        //     but processing other offers having external OfferIDs.  By adding this check, all processing of the given request is halted until the issue is resolved
        //     and a message is logged to the CRMOfferConnector Copient log file to notify the user what is wrong.
        foreach (Offers.OfferDefinition anOffer in os.m_offers)
        {
            if (anOffer.m_extOfferID == null)
            {
                log(string.Format("postOffers(Error: Required external OfferID not specified ExtOfferID( {0} ), OfferID( {1} ), EngineID( {2} ))", anOffer.m_extOfferID, anOffer.m_offerID, anOffer.m_engineID));
                return false;
            }
        }

        try
        {
            ei.verify(m_commonRoutines);
            importStatus = os.import(ei, m_logger, m_commonRoutines);
        }
        catch (ApplicationException e)
        {
            log(string.Format("Error: {0}", e.Message));
            return false;
        }
        return importStatus;

    } // end method postOffers()



    [WebMethod] //------------------------------------------------------------------------------------------------------------------
    public Offers getOffers(ExternalInterfaceID ei)
    {
        // find the offers that are available for the external interface
        log(string.Format("getOffers( {0} )", ei));
        ei.verify(m_commonRoutines);
        //log( Offers.offerFilesToExportQuery( ei ) );
        try
        {
            Offers o = new Offers(ei, m_commonRoutines);
            log(string.Format("Found {0} offers", o.m_offers.Length));
            return o;
        }
        catch (ApplicationException e)
        {
            log(e.Message);
        }
        return new Offers();
    } // end method getOffers()


    /* This class encapsulates a list of offerid/ExtOfferID pairs */
    public class OfferList
    {
        public struct OfferMapping
        {
            public long m_offerID;
            public string m_extOfferID;
            public string m_filename;

            public string as_str()
            {
                return string.Format("OfferID: {0}; ExtOfferID: {1}; file: {2}", m_offerID, m_extOfferID, m_filename);
            }

            public int isImported(Copient.CommonInc commonInterface, int eid)
            {
                commonInterface.QueryStr = string.Format(@"
                    SELECT TOP 1 * FROM
                    (SELECT CRMFileImportQueue.CRMStatusFlag, CPE_Incentives.EngineID, CRMFileImportQueue.LastUpdate
                    FROM CRMFileImportQueue WITH ( NoLock )
                    INNER JOIN CRMImportQueue ON CRMFileImportQueue.CRMImportQueuePKID = CRMImportQueue.PKID
                    INNER JOIN CPE_Incentives ON CPE_Incentives.IncentiveID = CRMImportQueue.OfferID
                        AND CPE_Incentives.InboundCRMEngineID = {0} AND CPE_incentives.ClientOfferID = {1}
                    UNION
                    SELECT CRMFileImportQueue.CRMStatusFlag, Offers.EngineId, CRMFileImportQueue.LastUpdate
                    FROM CRMFileImportQueue WITH ( NoLock )
                    INNER JOIN CRMImportQueue ON CRMFileImportQueue.CRMImportQueuePKID = CRMImportQueue.PKID
                    INNER JOIN Offers ON CRMImportQueue.OfferID = Offers.OfferID
                        AND Offers.InboundCRMEngineID = {0} AND Offers.ExtOfferID = {1}) as inLineView
                    ORDER BY LastUpdate DESC
                ", "@EngineId", "@ExtOfferId");
                commonInterface.DBParameters.Add("@ExtOfferId", SqlDbType.NVarChar).Value = m_extOfferID;
                commonInterface.DBParameters.Add("@EngineId", SqlDbType.Int).Value = eid;

                DataTable status = commonInterface.ExecuteQuery(Copient.DataBases.LogixRT);
                bool hasRows = status.Rows.Count > 0;
                if (hasRows)
                {
                    DataRow r = status.Rows[0];
                    if (r["CRMStatusFlag"] != null && !(r["CRMStatusFlag"] is DBNull))
                    {
                        int flag = int.Parse(r["CRMStatusFlag"].ToString());
                        return flag;
                    }
                }
                return 0;
            }

            public int getEngineID(Copient.CommonInc commonInterface)
            {
                commonInterface.QueryStr = string.Format(@"
                    SELECT EngineID FROM OfferIDs WITH ( NoLock ) where OfferID = {0};
                ", "@OfferID");
                commonInterface.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = m_offerID;
                DataTable status = commonInterface.ExecuteQuery(Copient.DataBases.LogixRT);
                bool hasRows = status.Rows.Count > 0;
                if (hasRows)
                {
                    DataRow r = status.Rows[0];
                    if (r["EngineID"] != null && !(r["EngineID"] is DBNull))
                    {
                        int engineID = int.Parse(r["EngineID"].ToString());
                        return engineID;
                    }
                }
                return (int)Copient.CommonInc.InstalledEngines.CPE;
            }

            public const string EXPORTED_OFFER_ID_UPDATE_QUERY =
            @"
                UPDATE CPE_Incentives WITH ( RowLock )
                    SET CPE_Incentives.ClientOfferID = {0}, CPE_Incentives.LastUpdate = GETDATE()
                    WHERE CPE_Incentives.IncentiveID = {1} AND CPE_Incentives.CRMEngineID = {2}
            ";
            public string OfferExportExternalIDUpdateQuerey()
            {
                return string.Format(EXPORTED_OFFER_ID_UPDATE_QUERY, "@ExtOfferID", "@OfferID", "@EngineID");
            }

            // CM
            public const string EXPORTED_OFFER_ID_UPDATE_QUERY_CM =
            @"
                UPDATE Offers WITH ( RowLock )
                    SET CRMSendStatus = 3
                    WHERE OfferId = {0} AND CRMEngineID = {1}
            ";
            public string OfferExportExternalIDUpdateQuereyCm()
            {
                return string.Format(EXPORTED_OFFER_ID_UPDATE_QUERY_CM, "@OfferID", "@EngineID");
            }

            // arguments are ( filename )
            public const string EXPORT_UPDATE_QUERY =
            @"
                UPDATE FEQ WITH ( RowLock )
                    SET FEQ.CRMStatusFlag = 1, FEQ.LastUpdate = GETDATE(), FEQ.Deleted = 'true'
                    FROM CRMFileExportQueue AS FEQ
                    WHERE FEQ.FileName = {0} AND EXISTS( select PKID FROM CRMExportQueue AS EQ WITH ( NoLock )
                                     WHERE EQ.PKID = FEQ.CRMExportQueuePKID AND EQ.ExtInterfaceID = {1})
            ";
            public string OfferExportAcknowledgeUpdateQuery()
            {
                return string.Format(EXPORT_UPDATE_QUERY, "@Filename", "@EngineID");
            }

            public const string EXPORT_UPDATE_DELETED_QUERY =
            @"
                UPDATE EQ WITH ( RowLock )
                    SET EQ.LastUpdate = GETDATE(), EQ.Deleted = 1
                    FROM CRMExportQueue AS EQ
                    WHERE EQ.Deleted = 0 AND EQ.OfferID = {0} AND EQ.ExtInterfaceID = {1}
                    AND NOT EXISTS ( SELECT CRMExportQueuePKID FROM CRMFileExportQueue AS FEQ WITH ( NoLock )
                                     WHERE FEQ.CRMExportQueuePKID = EQ.PKID AND FEQ.DELETED = 0 )
            ";
            // set CrmExportQueue for offer to "deleted", only if all files for offer have been set to "deleted"
            public string OfferExportAcknowledgeUpdateDeletedQuery()
            {
                return string.Format(EXPORT_UPDATE_DELETED_QUERY, "@OfferID", "@EngineID");
            }

            public void deleteMatchingOfferFile(Copient.CommonInc commonInterface, string filePath)
            {
                commonInterface.QueryStr = string.Format(@"
                    SELECT PKID FROM CRMFileExportQueue WITH ( NoLock ) WHERE FileName = {0} AND Deleted=0;
                ", "@Filename");
                commonInterface.DBParameters.Add("@Filename", SqlDbType.NVarChar).Value = m_filename;

                DataTable status = commonInterface.ExecuteQuery(Copient.DataBases.LogixRT);
                bool hasRows = status.Rows.Count > 0;
                if (!hasRows)
                {
                    File.Delete(filePath + m_filename);
                }
            } // end method deleteMatchingOfferFile()

            public void logOfferRetrieval(Copient.CommonInc commonInterface, int ExtInterfaceId)
            {
                const long ADMIN_USER_ID = 1;
                const int ACTIVITY_TYPE_OFFER = 3;
                const int SYS_OPTION_DEFAULT_LANGUAGE_ID = 1;
                int defaultLanguageID = int.Parse(commonInterface.Fetch_SystemOption(SYS_OPTION_DEFAULT_LANGUAGE_ID));
                string msg = Copient.PhraseLib.Lookup("history.offer-outboundretrieved", defaultLanguageID) + " (" + ExtInterfaceId.ToString() + ")";
                commonInterface.Activity_Log(ACTIVITY_TYPE_OFFER, m_offerID, ADMIN_USER_ID, msg);
            }

        } // end struct OfferMapping

        public OfferMapping[] m_offers; // map an extOfferID to an offer id

        public string as_str()
        {
            string s = "Offers:\n";
            foreach (OfferMapping anOffer in m_offers)
            {
                s += string.Format("  {0}\n", anOffer.as_str());
            }
            return s;
        } // as_str()

        public void acknowledge(Copient.CommonInc commonInterface, logger l, ExternalInterfaceID eid)
        {
            string exportedOffersPath = exportFilePath(commonInterface);

            commonInterface.Open_LogixRT();
            foreach (OfferMapping anOffer in m_offers)
            {
                // update the external offer id for the offer (CPE only)
                if ((anOffer.getEngineID(commonInterface) == (int)Copient.CommonInc.InstalledEngines.CPE) || (anOffer.getEngineID(commonInterface) == (int)Copient.CommonInc.InstalledEngines.UE))
                {
                    commonInterface.QueryStr = anOffer.OfferExportExternalIDUpdateQuerey();
                    commonInterface.DBParameters.Add("@ExtOfferID", SqlDbType.NVarChar).Value = anOffer.m_extOfferID;
                    commonInterface.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = anOffer.m_offerID;
                    commonInterface.DBParameters.Add("@EngineID", SqlDbType.Int).Value = eid.m_extInterfaceID;
                    //l.log( commonInterface.QueryStr );
                    commonInterface.ExecuteNonQuery(Copient.DataBases.LogixRT);
                }
                else if (anOffer.getEngineID(commonInterface) == (int)Copient.CommonInc.InstalledEngines.CM)
                {
                    commonInterface.QueryStr = anOffer.OfferExportExternalIDUpdateQuereyCm();
                    commonInterface.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = anOffer.m_offerID;
                    commonInterface.DBParameters.Add("@EngineID", SqlDbType.Int).Value = eid.m_extInterfaceID;
                    commonInterface.ExecuteNonQuery(Copient.DataBases.LogixRT);
                }

                // update the status of the exported file
                commonInterface.QueryStr = anOffer.OfferExportAcknowledgeUpdateQuery();
                commonInterface.DBParameters.Add("@Filename", SqlDbType.NVarChar).Value = anOffer.m_filename;
                commonInterface.DBParameters.Add("@EngineID", SqlDbType.Int).Value = eid.m_extInterfaceID;
                commonInterface.ExecuteNonQuery(Copient.DataBases.LogixRT);

                // update the status of the export queue
                commonInterface.QueryStr = anOffer.OfferExportAcknowledgeUpdateDeletedQuery();
                commonInterface.DBParameters.Add("@OfferID", SqlDbType.BigInt).Value = anOffer.m_offerID;
                commonInterface.DBParameters.Add("@EngineID", SqlDbType.Int).Value = eid.m_extInterfaceID;
                commonInterface.ExecuteNonQuery(Copient.DataBases.LogixRT);

                anOffer.logOfferRetrieval(commonInterface, eid.m_extInterfaceID);

                // delete the exported offer file
                try
                {
                    anOffer.deleteMatchingOfferFile(commonInterface, exportedOffersPath);
                }
                catch (ApplicationException e)
                {
                    l.log(string.Format("Failed to delete exported offer file {0} : {1}", anOffer.m_filename, e.Message));
                }

            } // end foreach

        } // end method acknowledge()

    } // end class OfferList



    [WebMethod] //------------------------------------------------------------------------------------------------------------------
    public bool ackOffers(ExternalInterfaceID ei, OfferList ol)
    {
        // take the list of offers provided by the external interface and mark them as having been successfully received by it
        log(string.Format("ackOffers( {0} )", ei));
        ei.verify(m_commonRoutines);
        ol.acknowledge(m_commonRoutines, m_logger, ei);
        return true;
    } // end ackOffers() WebMethod


    public class OfferStatuses
    {

        public struct OfferStatus
        {
            public string m_extOfferId;
            public int m_imported; // 0 - not yet; 1 - success; 2 - error
        }

        public OfferStatus[] m_offerImportStatuses;

        public OfferStatuses() { } /* default empty set */

        /* find the import statuses for each of the given offers that are available to the given external interface */
        public OfferStatuses(Copient.CommonInc commonInterface, ExternalInterfaceID ei, OfferList ol)
        {
            int numOffersToLookup = ol.m_offers.GetLength(0);
            m_offerImportStatuses = new OfferStatus[numOffersToLookup];
            int i = 0;
            foreach (OfferList.OfferMapping mapping in ol.m_offers)
            {
                m_offerImportStatuses[i].m_extOfferId = mapping.m_extOfferID;
                m_offerImportStatuses[i].m_imported = mapping.isImported(commonInterface, ei.m_extInterfaceID);
                ++i;
            }
        }

    } // end class OfferStatuses


    [WebMethod] //------------------------------------------------------------------------------------------------------------------
    public OfferStatuses getOfferStatus(ExternalInterfaceID ei, OfferList ol)
    {
        // return the import status for each of the given list of offers
        log(string.Format("offerStatus( {0} )", ei));
        ei.verify(m_commonRoutines);
        return new OfferStatuses(m_commonRoutines, ei, ol);
    }

    [WebMethod]
    public string aboutThisService()
    {
        return "CRMOfferConnector - " + sVersion;
    }

    public class OfferImmediateStatuses
    {

        public struct OfferImmediateStatusStruct
        {
            public string m_extOfferId;
            public long m_offerId;
            public int m_status; // 0 -> ok; 1 -> error
            public string m_status_msg;
        }

        public OfferImmediateStatusStruct[] m_offerImmediateStatuses;

        public OfferImmediateStatuses() { } /* default empty set */

        /* create array for import statuses for each of the given offers */
        public OfferImmediateStatuses(Copient.CommonInc commonInterface, ExternalInterfaceID ei, Offers os)
        {
            int numOffersToLookup = os.m_offers.GetLength(0);
            m_offerImmediateStatuses = new OfferImmediateStatusStruct[numOffersToLookup];
        }
    } // end class OfferImmediateStatuses

    public sealed class ControlCharacters
    {
        public const char Back = '\b';
        public const char Cr = '\r';
        public const string CrLf = "\r\n";
        public const char FormFeed = '\f';
        public const char Lf = '\n';
        public const string NewLine = "\r\n";
        public const char NullChar = '\0';
        public const char Quote = '"';
        public const string Tab = "\t";
        public const char VerticalTab = '\v';
    }

    [WebMethod] //------------------------------------------------------------------------------------------------------------------
    public OfferImmediateStatuses postOffersImmediate(ExternalInterfaceID ei, Offers os)
    {
        OfferImmediateStatuses objOfferStatuses;

        log(string.Format("postOffersImmediate( {0} )", ei));
        if (os == null)
        {
            log(string.Format("postOffersImmediate(Error: null value sent for Offers)"));
            return null;
        }

        // This fails silently without generating any notification in import() where it continues past any offers not having external OfferIDs, not processing them,
        //     but processing other offers having external OfferIDs.  By adding this check, all processing of the given request is halted until the issue is resolved
        //     and a message is logged to the CRMOfferConnector Copient log file to notify the user what is wrong.
        foreach (Offers.OfferDefinition anOffer in os.m_offers)
        {
            if (anOffer.m_extOfferID == null)
            {
                log(string.Format("postOffersImmediate(Error: Required external OfferID not specified ExtOfferID( {0} ), OfferID( {1} ), EngineID( {2} ))", anOffer.m_extOfferID, anOffer.m_offerID, anOffer.m_engineID));
                return null;
            }
        }

        try
        {
            // accept a set of CM offers from the external interface and import them in real time
            ei.verify(m_commonRoutines);
            objOfferStatuses = os.import_immediate(m_commonRoutines, m_logger, ei, os);
        }
        catch (ApplicationException e)
        {
            log(string.Format("Error: {0}", e.Message));
            return null;
        }
        return objOfferStatuses;

    } // end method postOffersImmediate()

    [WebMethod] //------------------------------------------------------------------------------------------------------------------
    public OfferImmediateStatuses transferCmOffersImmediate(ExternalInterfaceID ei, Offers os)
    {
        OfferImmediateStatuses objOfferStatuses;

        log(string.Format("transferCmOffersImmediate( {0} )", ei));
        if (os == null)
        {
            log(string.Format("Error: null value sent for Offers"));
            return null;
        }

        try
        {
            // accept a set of CM offers from the external interface and import them in real time
            ei.verify(m_commonRoutines);
            objOfferStatuses = os.transfer_cm_immediate(m_commonRoutines, m_logger, ei, os);
        }
        catch (ApplicationException e)
        {
            log(string.Format("Error: {0}", e.Message));
            return null;
        }
        return objOfferStatuses;

    } // end method transferCmOffersImmediate()

    public class OfferTranslateStatuses
    {

        public struct OfferTranslateStatusStruct
        {
            public long m_cmOfferId;
            public long m_ueOfferId;
            public int m_status; // 0 -> ok; 1 -> error; -1 -> warning ; -2 -> skipped 
            public string m_status_msg;
        }

        public OfferTranslateStatusStruct[] m_offerTranslateStatuses;

        public OfferTranslateStatuses() { } /* default empty set */

        /* create array for import statuses for each of the given offers */
        public OfferTranslateStatuses(Copient.CommonInc commonInterface, ExternalInterfaceID ei, Offers os)
        {
            int numOffersToLookup = os.m_offers.GetLength(0);
            m_offerTranslateStatuses = new OfferTranslateStatusStruct[numOffersToLookup];
        }
    } // end class OfferImmediateStatuses

    [WebMethod] //------------------------------------------------------------------------------------------------------------------   
    public OfferTranslateStatuses translateCmToUe(ExternalInterfaceID ei, Offers os)
    {
        OfferTranslateStatuses objOfferStatuses;

        log(string.Format("translateCmToUe( {0} )", ei));
        if (os == null)
        {
            log(string.Format("Error: null value sent for Offers"));
            return null;
        }

        try
        {
            // accept a set of CM offers from the external interface and import them in real time
            ei.verify(m_commonRoutines);
            objOfferStatuses = os.translate_cm_to_ue(m_commonRoutines, m_logger, ei, os, false);
        }
        catch (ApplicationException e)
        {
            log(string.Format("Error: {0}", e.Message));
            return null;
        }
        return objOfferStatuses;

    } // end method translateCmToUe()

    [WebMethod] //------------------------------------------------------------------------------------------------------------------   
    public OfferTranslateStatuses translateCmToUeEx(ExternalInterfaceID ei, Offers os)
    {
        OfferTranslateStatuses objOfferStatuses;

        log(string.Format("translateCmToUeEx( {0} )", ei));
        if (os == null)
        {
            log(string.Format("Error: null value sent for Offers"));
            return null;
        }

        try
        {
            // accept a set of CM offers from the external interface and import them in real time
            ei.verify(m_commonRoutines);
            objOfferStatuses = os.translate_cm_to_ue(m_commonRoutines, m_logger, ei, os, true);
        }
        catch (ApplicationException e)
        {
            log(string.Format("Error: {0}", e.Message));
            return null;
        }
        return objOfferStatuses;

    } // end method translateCmToUe()

    [WebMethod] //------------------------------------------------------------------------------------------------------------------   
    public OfferTranslateStatuses deleteTranslatedOffer(ExternalInterfaceID ei, Offers os)
    {
        OfferTranslateStatuses objOfferStatuses;

        log(string.Format("deleteTranslatedOffer( {0} )", ei));
        if (os == null)
        {
            log(string.Format("Error: null value sent for Offers"));
            return null;
        }

        try
        {
            // accept a set of CM offers from the external interface and import them in real time
            ei.verify(m_commonRoutines);
            objOfferStatuses = os.delete_translated_offer(m_commonRoutines, m_logger, ei, os);
        }
        catch (ApplicationException e)
        {
            log(string.Format("Error: {0}", e.Message));
            return null;
        }
        return objOfferStatuses;

    } // end method deleteTranslatedOffer()
} // end class CRMOfferConnector //
