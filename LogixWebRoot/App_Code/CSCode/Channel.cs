using System;
using System.Collections.Generic;
using System.Web;
using System.Web.Services;
using System.Data;
using System.Xml;
using System.IO;
using Copient;
using CMS;
using CMS.AMS;
using CMS.AMS.Contract;
using CMS.Contract;
using CMS.AMS.Models;


[WebService(Namespace = "http://ncr.cms.ams.com/")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
// To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
// [System.Web.Script.Services.ScriptService]

public class ChannelService : System.Web.Services.WebService
{
  // version:7.3.1.138972.Official Build (SUSDAY10202)

  private CMS.AMS.Common m_common;
  private CMS.AMS.ConnectorInc m_connectorInc;
  private Copient.ChannelWebSvc m_channelInc;
  private CMS.AMS.AuthLib m_authInc;
  private CMS.AMS.CustomerCommon m_custCommon;
  private CMS.AMS.CustomerLogin m_custLogin;
  private IOffer m_Offers;
  private ILogger m_logger;
  private IErrorHandler m_errHandler;
  private const int CHANNELID = 2;
  private Copient.CommonInc m_legacyCommon; //LEGACY CODE TO BE REMOVED
  private IPhraseLib m_phraseLib;
  Copient.CommonInc MyCommon = new Copient.CommonInc();
  private CMS.AMS.Customer m_Customer; 

  // -------------------------------------------------------------------------------------------------------------------------------------

  [WebMethod]
  public XmlDocument GetLogonIdentifierTypes(string GUID) {
    string methodName = "GetLogonIdentifierTypes";
    XmlDocument ResponseXML = new XmlDocument();
    StringWriter sw = null;
    XmlWriter Writer = null;
    int ChannelID = 0;

    try {
      Startup();
      m_connectorInc.Init_ResponseXML(methodName, ref sw, ref Writer);
      ChannelID = m_authInc.Is_Valid_ChannelGUID(GUID);
      if (ChannelID == 0) {
        m_connectorInc.Generate_Status_XML(ref Writer, methodName, CMS.CMSException.StatusCodes.INVALID_GUID, m_phraseLib.Detokenize("term.guidnotvalid", GUID, m_common.Get_AppInfo().AppName), false);
      }
      else {
        FetchLogonIdentifierTypes(ref Writer);
        m_connectorInc.Generate_Status_XML(ref Writer, methodName, CMS.CMSException.StatusCodes.SUCCESS, m_phraseLib.Lookup("term.success"), true);
      }
    }
    catch (Exception ex) {
      ProcessException(ex, methodName, ref Writer, ref sw);
    }
    Close_ResponseXML(ref Writer, ref sw, ref ResponseXML);
    Shutdown();
    return ResponseXML;
  }

  // -------------------------------------------------------------------------------------------------------------------------------------

  [WebMethod]
  public XmlDocument Logon(string GUID, string ExtIdentifier, string ExtIDType, string Password) {
    string methodName = "Logon";
    XmlDocument ResponseXML = new XmlDocument();
    StringWriter sw = null;
    XmlWriter Writer = null;
    CMS.AMS.Customer Customer;
    Int64 CustomerPK;
    CMS.AMS.CustIDTypes custIDTypes;
    int ChannelID = 0;

    try {
      Startup();
      m_connectorInc.Init_ResponseXML(methodName, ref sw, ref Writer);
      ChannelID = m_authInc.Is_Valid_ChannelGUID(GUID);
      if (ChannelID == 0) {
        m_connectorInc.Generate_Status_XML(ref Writer, methodName, CMS.CMSException.StatusCodes.INVALID_GUID, m_phraseLib.Detokenize("term.guidnotvalid", GUID, m_common.Get_AppInfo().AppName), false);
      }
      else {
        custIDTypes = new CMS.AMS.CustIDTypes(m_common, ExtIDType);
        CustomerPK = m_custLogin.CustomerLogin(ExtIdentifier, custIDTypes.Get_CustIDType().CardTypeID, Password);
        if (CustomerPK > 0) {
          Customer = new CMS.AMS.Customer(m_common, CustomerPK);
          Writer.WriteStartElement("Customer");
          Writer.WriteElementString("AuthToken", m_common.Get_Requestor().Customer.AuthToken);
          Writer.WriteElementString("CustomerPK", Customer.CustomerPK.ToString());
          Writer.WriteElementString("ExtIdentifier", ExtIdentifier);
          Writer.WriteElementString("ExtIDType", ExtIDType);
          Writer.WriteElementString("IdentifierType", custIDTypes.Get_CustIDType().CardTypeID.ToString());
          Writer.WriteEndElement();
          m_connectorInc.Generate_Status_XML(ref Writer, methodName, CMS.CMSException.StatusCodes.SUCCESS, m_phraseLib.Lookup("term.success"), true);
        }
      }
    }
    catch (Exception ex) {
      ProcessException(ex, methodName, ref Writer, ref sw);
    }
    Close_ResponseXML(ref Writer, ref sw, ref ResponseXML);
    Shutdown();
    return ResponseXML;
  }

  // -------------------------------------------------------------------------------------------------------------------------------------

  [WebMethod]
  public XmlDocument UpdatePassword(string GUID, string AuthToken, string ExtIdentifier, string ExtIDType, string NewPassword) {
    string methodName = "UpdatePassword";
    XmlDocument ResponseXML = new XmlDocument();
    StringWriter sw = null;
    XmlWriter Writer = null;
    CMSException.StatusCodes StatusCode = CMSException.StatusCodes.GENERAL_ERROR;
    CMS.AMS.Customer Customer;
    Int64 CustomerPK = -1;
    int index;
    CMS.CustomerIdentifier CustID = new CMS.CustomerIdentifier(m_common);
    int ChannelID = 0;

    try {
      Startup();
      m_connectorInc.Init_ResponseXML(methodName, ref sw, ref Writer);
      ChannelID = m_authInc.Is_Valid_ChannelGUID(GUID);
      if (ChannelID == 0) {
        m_connectorInc.Generate_Status_XML(ref Writer, methodName, CMS.CMSException.StatusCodes.INVALID_GUID, m_phraseLib.Detokenize("term.guidnotvalid", GUID, m_common.Get_AppInfo().AppName), false);
      }
      else {
        CustomerPK = m_custCommon.LookupCustomerFromAuthToken(AuthToken);
        if (IsAuthTokenValid(Writer, methodName, AuthToken, CustomerPK)) {
          Customer = new CMS.AMS.Customer(m_common, CustomerPK);
          ExtIdentifier = m_custCommon.Pad_ExtCardID(ExtIdentifier, ExtIDType);
          CMS.AMS.CustIDTypes IDType = new CMS.AMS.CustIDTypes(m_common, ExtIDType);
          if (!IDType.Get_CustIDType().CustomerCanUpdatePIN) {
            StatusCode = (CMSException.StatusCodes)CMS.CMSException.CUST_ERROR_CODES.PASSWORD_CAN_NOT_BE_UPDATED;
            m_connectorInc.Generate_Status_XML(ref Writer, methodName, StatusCode, m_phraseLib.Lookup("term.idtypepasswordnotupdateable"), false);  // Passwords for this identifier type can not be updated by customers
          }
          else {
            for (index = 0; index < Customer.Identifiers.Count; index++) {
              if ((Customer.Identifiers[index].ExtCardID == ExtIdentifier.ToLower()) && (Customer.Identifiers[index].IdentifierTypeID == IDType.Get_CustIDType().CardTypeID)) {
                // we found the identifier/IDType for which the password should be updated
                CustID = Customer.Identifiers[index];
                break;
              }
            }
            if (CustID.CardPK != 0) {
              // we found a matching customer identifier 
              CustID.Password = NewPassword;
              Customer.UpdateIdentifier(CustID);
              m_connectorInc.Generate_Status_XML(ref Writer, methodName, CMS.CMSException.StatusCodes.SUCCESS, m_phraseLib.Detokenize("term.success"), true);
            }
            else {
              // invalid customerID/IDType
              StatusCode = (CMSException.StatusCodes)CMS.CMSException.CUST_ERROR_CODES.IDENTIFIER_NOT_FOUND;
              m_connectorInc.Generate_Status_XML(ref Writer, methodName, StatusCode, m_phraseLib.Lookup("term.custidnotfound"), false);  // Customer identifier not found
            }
          } // if CustomerCanUpdatePIN
        } // if AuthToken is valid
      }  // if GUID is valid
    }
    catch (Exception ex) {
      ProcessException(ex, methodName, ref Writer, ref sw);
    }
    Close_ResponseXML(ref Writer, ref sw, ref ResponseXML);
    Shutdown();
    return ResponseXML;
  }

  // -------------------------------------------------------------------------------------------------------------------------------------

  [WebMethod]
  public XmlDocument GetCustomerDetails(string GUID, string AuthToken) {
    string methodName = "GetCustomerDetails";
    XmlDocument ResponseXML = new XmlDocument();
    StringWriter sw = null;
    XmlWriter Writer = null;
    Int64 CustomerPK = -1;
    int ChannelID = 0;

    try {
      Startup();
      m_connectorInc.Init_ResponseXML(methodName, ref sw, ref Writer);
      ChannelID = m_authInc.Is_Valid_ChannelGUID(GUID);
      if (ChannelID == 0) {
        m_connectorInc.Generate_Status_XML(ref Writer, methodName, CMS.CMSException.StatusCodes.INVALID_GUID, m_phraseLib.Detokenize("term.guidnotvalid", GUID, m_common.Get_AppInfo().AppName), false);  // "GUID '" + GUID + "' is not valid for " + m_common.Get_AppInfo().AppName
      }
      else {
        CustomerPK = m_custCommon.LookupCustomerFromAuthToken(AuthToken);
        if (IsAuthTokenValid(Writer, methodName, AuthToken, CustomerPK)) {
          FetchCustomerDetails(ref Writer, CustomerPK);
        } // if AuthToken is valid
      }  // if GUID is valid
    }
    catch (Exception ex) {
      ProcessException(ex, methodName, ref Writer, ref sw);
    }
    Close_ResponseXML(ref Writer, ref sw, ref ResponseXML);
    Shutdown();
    return ResponseXML;
  }

  // -------------------------------------------------------------------------------------------------------------------------------------

  /// <summary>
  /// Updates an existing customer’s general information example of which includes name, address, and contact information
  /// </summary>
  /// <param name="GUID">Global Unique Identifier used to authenticate the web service caller and provide means of identification of the caller</param>
  /// <param name="AuthToken">Authentication Token used to authenticate the specific caller as customer</param>
  /// <param name="CustomerDetailXML">XML string containing the tags for the customer details record as it should be saved</param>
  /// <returns>Status of operation as XmlDocument</returns>
  [WebMethod]
  public XmlDocument SaveCustomerDetails(string GUID, string AuthToken, string CustomerDetailXML) {

    string methodName = "SaveCustomerDetails";
    XmlWriter writer = null;
    StringWriter sw = null;
    XmlDocument responseXML = new XmlDocument();
    XmlDocument xmlInput = null;
    string errMsg = "";
    long CustomerPK = 0;

    try {
      Startup();
      m_connectorInc.Init_ResponseXML(methodName, ref sw, ref writer);

      // check validity of input arguments
      if (m_authInc.Is_Valid_ChannelGUID(GUID) == 0) {
        m_connectorInc.Generate_Status_XML(ref writer, methodName, CMS.CMSException.StatusCodes.INVALID_GUID, m_phraseLib.Detokenize("term.guidnotvalid", GUID, m_common.Get_AppInfo().AppName), false);
      }
      else if (!m_connectorInc.ConvertStringToXML(CustomerDetailXML, ref xmlInput)) {
        m_connectorInc.Generate_Status_XML(ref writer, methodName, CMS.CMSException.StatusCodes.INVALID_XML_DOCUMENT, m_phraseLib.Detokenize("term.invalidparameterxmldocument", "CustomerDetailXML"), false);
      }
      else if (!m_connectorInc.IsValidXmlDocument("ChannelSaveCustomerDetailsRequest.xsd", xmlInput, out errMsg)) {
        m_connectorInc.Generate_Status_XML(ref writer, methodName, CMS.CMSException.StatusCodes.INVALID_XML_DOCUMENT, m_phraseLib.Detokenize("term.invalidxmlconformxsd", "CustomerDetailXML", "ChannelSaveCustomerDetailsRequest.xsd: " + errMsg), false);
      }
      else {

        CustomerPK = m_custCommon.LookupCustomerFromAuthToken(AuthToken);
        if (IsAuthTokenValid(writer, methodName, AuthToken, CustomerPK)) {

          // all validation tests passed - create new customer and save details
          CMS.AMS.Customer tempCustomer = ParseCustomerDetails(xmlInput, CustomerPK);
          tempCustomer.SaveCustomer();

          m_connectorInc.Generate_Status_XML(ref writer, methodName, CMS.CMSException.StatusCodes.SUCCESS, m_phraseLib.Lookup("term.success"), true);
        }
      }
    }
    catch (Exception ex) {
      ProcessException(ex, methodName, ref writer, ref sw);
    }
    Close_ResponseXML(ref writer, ref sw, ref responseXML);
    Shutdown();

    return responseXML;
  }

  // -------------------------------------------------------------------------------------------------------------------------------------

  [WebMethod]
  public XmlDocument GetOfferList(string GUID, string AuthToken, int PageNum) {
    string methodName = "GetOfferList";
    XmlDocument ResponseXML = new XmlDocument();
    StringWriter sw = null;
    XmlWriter Writer = null;
    Int64 CustomerPK = -1;
    int ChannelID = 0;

    try {
      Startup();
      m_connectorInc.Init_ResponseXML(methodName, ref sw, ref Writer);
      ChannelID = m_authInc.Is_Valid_ChannelGUID(GUID);
      if (ChannelID == 0) {
        m_connectorInc.Generate_Status_XML(ref Writer, methodName, CMS.CMSException.StatusCodes.INVALID_GUID, m_phraseLib.Detokenize("term.guidnotvalid", GUID, m_common.Get_AppInfo().AppName), false);  // "GUID '" + GUID + "' is not valid for " + m_common.Get_AppInfo().AppName
      }
      else {
        CustomerPK = m_custCommon.LookupCustomerFromAuthToken(AuthToken);
        if (IsAuthTokenValid(Writer, methodName, AuthToken, CustomerPK)) {
          FetchOfferList(ref Writer, ChannelID, CustomerPK, PageNum, GUID, AuthToken);
        } // if AuthToken is valid
      }  // if GUID is valid
    }
    catch (Exception ex) {
      ProcessException(ex, methodName, ref Writer, ref sw);
    }
    Close_ResponseXML(ref Writer, ref sw, ref ResponseXML);
    Shutdown();
    return ResponseXML;
  }

  // -------------------------------------------------------------------------------------------------------------------------------------

  /// <summary>
  /// Retrieves details for all channels for passed-in offer IDs
  /// </summary>
  /// <param name="GUID">Global Unique Identifier used to authenticate the web service caller and provide means of identification of the caller</param>
  /// <param name="AuthToken">Authentication Token used to authenticate the specific caller as customer</param>
  /// <param name="OfferIDsXML">XML string containing the offer IDs for the offer channel records that should be retrieved</param>
  /// <returns>Status of operation as XmlDocument</returns>
  [WebMethod]
  public XmlDocument GetChannelOfferDetails(string GUID, string AuthToken, string OfferIDsXML) {

    string methodName = "GetChannelOfferDetails";
    XmlWriter writer = null;
    StringWriter sw = null;
    XmlDocument responseXML = new XmlDocument();
    XmlDocument xmlInput = null;
    string errMsg = "";
    int channelID = 0;
    long CustomerPK = 0;

    try {
      Startup();
      m_connectorInc.Init_ResponseXML(methodName, ref sw, ref writer);

      // check validity of input arguments
      channelID = m_authInc.Is_Valid_ChannelGUID(GUID);
      if (channelID == 0) {
        m_connectorInc.Generate_Status_XML(ref writer, methodName, CMS.CMSException.StatusCodes.INVALID_GUID, m_phraseLib.Detokenize("term.guidnotvalid", GUID, m_common.Get_AppInfo().AppName), false);
      }
      else if (!m_connectorInc.ConvertStringToXML(OfferIDsXML, ref xmlInput)) {
        m_connectorInc.Generate_Status_XML(ref writer, methodName, CMS.CMSException.StatusCodes.INVALID_XML_DOCUMENT, m_phraseLib.Detokenize("term.invalidparameterxmldocument", "OfferIDsXML"), false);
      }
      else if (!m_connectorInc.IsValidXmlDocument("ChannelGetChannelOfferDetailsRequest.xsd", xmlInput, out errMsg)) {
        m_connectorInc.Generate_Status_XML(ref writer, methodName, CMS.CMSException.StatusCodes.INVALID_XML_DOCUMENT, m_phraseLib.Detokenize("term.invalidxmlconformxsd", "OfferIDsXML", "ChannelGetChannelOfferDetailsRequest.xsd: " + errMsg), false);
      }
      else {
        CustomerPK = m_custCommon.LookupCustomerFromAuthToken(AuthToken);
        if (IsAuthTokenValid(writer, methodName, AuthToken, CustomerPK)) {

          // all validation tests passed - get channel details for all offer IDs
          FetchChannelOfferDetails(ref writer, xmlInput, channelID);
          m_connectorInc.Generate_Status_XML(ref writer, methodName, CMS.CMSException.StatusCodes.SUCCESS, m_phraseLib.Lookup("term.success"), true);
        }
      }
    }
    catch (Exception ex) {
      ProcessException(ex, methodName, ref writer, ref sw);
    }
    Close_ResponseXML(ref writer, ref sw, ref responseXML);
    Shutdown();

    return responseXML;
  }

  // -------------------------------------------------------------------------------------------------------------------------------------

  /// <summary>
  /// Links a Customer to an Offer
  /// </summary>
  /// <param name="GUID">Global Unique Identifier used to authenticate the web service caller and provide means of identification of the caller</param>
  /// <param name="AuthToken">Authentication Token used to authenticate the specific caller as customer</param>
  /// <param name="OfferID">String containing the OfferID that will the Customer wants to Opt-in to</param>
  /// <returns>Status of operation as XmlDocument</returns>
  [WebMethod]
  public XmlDocument OptInToOffer(string GUID, string AuthToken, long OfferID) {

    string methodName = "OptInToOffer";
    XmlWriter writer = null;
    StringWriter sw = null;
    XmlDocument responseXML = new XmlDocument();
    int channelID = 0;
    long CustomerPK = 0;
    int EngineID = -1;
    string OfferName = string.Empty;
    try {
      Startup();
      m_logger.WriteInfo(String.Format("Begin {0} for GUID :  {1}  AuthToken : {2}  Offer ID : {3}", methodName, GUID, AuthToken, OfferID));
      m_connectorInc.Init_ResponseXML(methodName, ref sw, ref writer);

      // check validity of input arguments
      channelID = m_authInc.Is_Valid_ChannelGUID(GUID);
      if (channelID == 0) {
        GetXMLResponse(ref writer, methodName, CMS.CMSException.StatusCodes.INVALID_GUID, m_phraseLib.Detokenize("term.guidnotvalid", GUID, m_common.Get_AppInfo().AppName), false);
      }
      else {
        CustomerPK = m_custCommon.LookupCustomerFromAuthToken(AuthToken);
        if (IsAuthTokenValid(writer, methodName, AuthToken, CustomerPK)) {

          m_logger.WriteInfo(String.Format("Customer ID : {0} is trying to Opt-In into Offer ID : {1}", CustomerPK, OfferID));

          if (!m_Offers.IsDeployedOfferExists(OfferID, ref EngineID, ref OfferName)) {
            GetXMLResponse(ref writer, methodName, CMS.CMSException.StatusCodes.OFFER_DOES_NOT_EXIST, m_phraseLib.Detokenize("term.offerexpiredordeleted", OfferID), false);
          }
          else if (EngineID != 0 && EngineID != 2 && EngineID != 9) {
            GetXMLResponse(ref writer, methodName, CMS.CMSException.StatusCodes.OFFER_ENGINE_ID_INVALID, m_phraseLib.Detokenize("term.invalidofferengine", EngineID), false);
          }
          else if (!m_Offers.IsDeployedOfferOptable(OfferID, EngineID)) {
            GetXMLResponse(ref writer, methodName, CMS.CMSException.StatusCodes.OFFER_NOT_OPTABLE, m_phraseLib.Detokenize("term.offernotoptable", OfferID), false);
          }
          else if (m_Offers.IsCustomerOptedInToDeployedOffer(OfferID, EngineID, CustomerPK)) {
            m_connectorInc.Generate_Status_XML(ref writer, methodName, CMS.CMSException.StatusCodes.SUCCESS, m_phraseLib.Lookup("term.success"), true);
            m_logger.WriteInfo(String.Format("Customer ID : {0} already opted in to Offer ID :  {1}", CustomerPK, OfferID));
          }
          else if (!m_Offers.IsCustomerEligibleToOptInToDeployedOffer(OfferID, EngineID, CustomerPK)) {
            GetXMLResponse(ref writer, methodName, CMS.CMSException.StatusCodes.CUSTOMER_NOT_ELIGIBLE_TO_OPTIN_TO_OFFER, m_phraseLib.Detokenize("term.customernoteligibletooptintooffer", OfferID), false);
          }
          else if (!m_Offers.OptInToOffer(OfferID, EngineID, CustomerPK)) {
            GetXMLResponse(ref writer, methodName, CMS.CMSException.StatusCodes.OFFER_OPTIN_FAILED, m_phraseLib.Detokenize("term.optinfailed", OfferID), false);
          }
          else {
            m_Customer = new CMS.AMS.Customer(m_common, CustomerPK);
            string InitialCard = m_Customer.InitialCardID;
            string CardTypeDesc = ((commonShared.CardTypes)m_Customer.CustomerTypeID).ToString();
            CMS.AMS.Models.CustomerGroup objCustomerGroup = m_Offers.GetOfferDefaultCustomerGroup(OfferID, EngineID);
            long CustomerGroupID = objCustomerGroup.CustomerGroupID;

            WriteToActivityLog(25, CustomerPK, -1, m_phraseLib.Detokenize("term.customeroptedin", OfferID, OfferName));
            WriteToActivityLog(4, CustomerGroupID, -1, m_phraseLib.Detokenize("history.cgroup-optedin", InitialCard, CardTypeDesc));

            m_logger.WriteInfo(String.Format("Customer ID : {0} Opted in to offer ID : {1}", CustomerPK, OfferID));
            GetXMLResponse(ref writer, methodName, CMS.CMSException.StatusCodes.SUCCESS, m_phraseLib.Lookup("term.success"), true);
          }
        }
      }
      m_logger.WriteInfo(String.Format("End {0} for GUID :  {1}  AuthToken : {2}  Offer ID : {3}", methodName, GUID, AuthToken, OfferID));
    }
    catch (Exception ex) {
      ProcessException(ex, methodName, ref writer, ref sw);
      m_logger.WriteInfo(ex.Message ); 
    }
    Close_ResponseXML(ref writer, ref sw, ref responseXML);
    Shutdown();
    return responseXML;
  }

  // -------------------------------------------------------------------------------------------------------------------------------------

  /// <summary>
  /// Removes a Customer from an Offer
  /// </summary>
  /// <param name="GUID">Global Unique Identifier used to authenticate the web service caller and provide means of identification of the caller</param>
  /// <param name="AuthToken">Authentication Token used to authenticate the specific caller as customer</param>
  /// <param name="OfferID">String containing the OfferID that will the Customer wants to Opt-out of</param>
  /// <returns>Status of operation as XmlDocument</returns>
  [WebMethod]
  public XmlDocument OptOutOfOffer(string GUID, string AuthToken, long OfferID) {

    string methodName = "OptOutOfOffer";
    XmlWriter writer = null;
    StringWriter sw = null;
    XmlDocument responseXML = new XmlDocument();
    int channelID = 0;
    long CustomerPK = 0;
    int EngineID = 0;
    string OfferName = string.Empty;
    try {
      Startup();
      m_logger.WriteInfo(String.Format("Begin {0} for GUID :  {1}  AuthToken : {2}  Offer ID : {3}", methodName, GUID, AuthToken, OfferID));
      m_connectorInc.Init_ResponseXML(methodName, ref sw, ref writer);

      // check validity of input arguments
      channelID = m_authInc.Is_Valid_ChannelGUID(GUID);
      if (channelID == 0) {
        GetXMLResponse(ref writer, methodName, CMS.CMSException.StatusCodes.INVALID_GUID, m_phraseLib.Detokenize("term.guidnotvalid", GUID, m_common.Get_AppInfo().AppName), false);
      }
      else {
        CustomerPK = m_custCommon.LookupCustomerFromAuthToken(AuthToken);
        if (IsAuthTokenValid(writer, methodName, AuthToken, CustomerPK)) {

          m_logger.WriteInfo(String.Format("Customer ID : {0} try to Opt-out from offer ID : {1}", CustomerPK, OfferID));

          if (!m_Offers.IsDeployedOfferExists(OfferID, ref EngineID, ref OfferName)) {
            GetXMLResponse(ref writer, methodName, CMS.CMSException.StatusCodes.OFFER_DOES_NOT_EXIST, m_phraseLib.Detokenize("term.offerexpiredordeleted", OfferID), false);
          }
          else if (EngineID != 0 && EngineID != 2 && EngineID != 9) {
            GetXMLResponse(ref writer, methodName, CMS.CMSException.StatusCodes.OFFER_ENGINE_ID_INVALID, m_phraseLib.Detokenize("term.invalidofferengine", EngineID), false);
          }
          else if (!m_Offers.IsCustomerOptedInToDeployedOffer(OfferID, EngineID, CustomerPK)) {
            m_connectorInc.Generate_Status_XML(ref writer, methodName, CMS.CMSException.StatusCodes.CUSTOMER_NOT_OPTED_IN, m_phraseLib.Lookup("term.customermustoptin"), false);
            m_logger.WriteInfo(String.Format("Customer ID: {0} must Opt-In before Opting Out from offer ID : {1}", CustomerPK, OfferID)); 
          }
          else if (!m_Offers.OptOutOfOffer(OfferID, EngineID, CustomerPK)) {
            GetXMLResponse(ref writer, methodName, CMS.CMSException.StatusCodes.OFFER_OPTOUT_FAILED, m_phraseLib.Detokenize("term.optoutfailed", OfferID), false);
          }
          else {
            m_Customer = new CMS.AMS.Customer(m_common, CustomerPK);
            string InitialCard = m_Customer.InitialCardID;
            string CardTypeDesc = ((commonShared.CardTypes)m_Customer.CustomerTypeID).ToString();
            CMS.AMS.Models.CustomerGroup objCustomerGroup = m_Offers.GetOfferDefaultCustomerGroup(OfferID, EngineID);
            long CustomerGroupID = objCustomerGroup.CustomerGroupID;

            WriteToActivityLog(25, CustomerPK, -1, m_phraseLib.Detokenize("term.customeroptedout", OfferID, OfferName));
            WriteToActivityLog(4, CustomerGroupID, -1, m_phraseLib.Detokenize("history.cgroup-optedout", InitialCard, CardTypeDesc));

            m_logger.WriteInfo(String.Format("Customer ID : {0} successfully Opt-out from offer ID : {1}", CustomerPK, OfferID)); 
            GetXMLResponse(ref writer, methodName, CMS.CMSException.StatusCodes.SUCCESS, m_phraseLib.Lookup("term.success"), true);
          }
        }
      }
      m_logger.WriteInfo(String.Format("End {0} for GUID :  {1}  AuthToken : {2}  Offer ID : {3}", methodName, GUID, AuthToken, OfferID));
    }
    catch (Exception ex) {
      ProcessException(ex, methodName, ref writer, ref sw);
      m_logger.WriteInfo(ex.Message); 
    }
    Close_ResponseXML(ref writer, ref sw, ref responseXML);
    Shutdown();
    return responseXML;
  }
  // -------------------------------------------------------------------------------------------------------------------------------------

  [WebMethod]
  public XmlDocument Logout(string GUID, string AuthToken) {

    string methodName = "Logout";
    XmlWriter writer = null;
    StringWriter sw = null;
    XmlDocument responseXML = new XmlDocument();
    int channelID = 0;
    int DeleteStatus = 0;
    try {
      Startup();
      m_connectorInc.Init_ResponseXML(methodName, ref sw, ref writer);
      // check validity of input arguments
      channelID = m_authInc.Is_Valid_ChannelGUID(GUID);
      if (channelID == 0) {
        m_connectorInc.Generate_Status_XML(ref writer, methodName, CMS.CMSException.StatusCodes.INVALID_GUID, m_phraseLib.Detokenize("term.guidnotvalid", GUID, m_common.Get_AppInfo().AppName), false);
      }
      else {
        DeleteStatus = m_custCommon.CustomerLogout(AuthToken);
        m_connectorInc.Generate_Status_XML(ref writer, methodName, CMS.CMSException.StatusCodes.SUCCESS, m_phraseLib.Lookup("term.success"), true);
      }
    }
    catch (Exception ex) {
      ProcessException(ex, methodName, ref writer, ref sw);
    }
    Close_ResponseXML(ref writer, ref sw, ref responseXML);
    Shutdown();

    return responseXML;
  }

  // -------------------------------------------------------------------------------------------------------------------------------------

  [WebMethod]
  public XmlDocument GetNumOffers(string GUID, string AuthToken) {
    string methodName = "GetNumOffers";
    XmlDocument ResponseXML = new XmlDocument();
    StringWriter sw = null;
    XmlWriter Writer = null;
    Int64 CustomerPK = -1;
    int ChannelID = 0;

    try {
      Startup();
      m_connectorInc.Init_ResponseXML(methodName, ref sw, ref Writer);
      ChannelID = m_authInc.Is_Valid_ChannelGUID(GUID);
      if (ChannelID == 0) {
        m_connectorInc.Generate_Status_XML(ref Writer, methodName, CMS.CMSException.StatusCodes.INVALID_GUID, m_phraseLib.Detokenize("term.guidnotvalid", GUID, m_common.Get_AppInfo().AppName), false);  // "GUID '" + GUID + "' is not valid for " + m_common.Get_AppInfo().AppName
      }
      else {
        CustomerPK = m_custCommon.LookupCustomerFromAuthToken(AuthToken);
        if (IsAuthTokenValid(Writer, methodName, AuthToken, CustomerPK)) {
          Writer.WriteStartElement("TotalOffers");
          Writer.WriteAttributeString("TotalOffers", m_channelInc.NumChannelOffers(ChannelID, CustomerPK).ToString());
          Writer.WriteEndElement();
        } // if AuthToken is valid
      }  // if GUID is valid
    }
    catch (Exception ex) {
      ProcessException(ex, methodName, ref Writer, ref sw);
    }
    Close_ResponseXML(ref Writer, ref sw, ref ResponseXML);
    Shutdown();
    return ResponseXML;
  }

  // =====================================================================================================================================

  private void Startup() {
    CurrentRequest.Resolver.AppName = "ChannelWS";
    m_common = CurrentRequest.Resolver.Resolve<CMS.AMS.Common>();
    if (m_common.LRT_Connection_State() == ConnectionState.Closed) { m_common.Open_LogixRT(); }
    if (m_common.LXS_Connection_State() == ConnectionState.Closed) { m_common.Open_LogixXS(); }

    m_legacyCommon = new CommonInc();
    if (m_legacyCommon.LRTadoConn.State == ConnectionState.Closed) { m_legacyCommon.Open_LogixRT(); }
    if (m_legacyCommon.LXSadoConn.State == ConnectionState.Closed) { m_legacyCommon.Open_LogixXS(); }
    m_connectorInc = new CMS.AMS.ConnectorInc(m_common);

    m_common.Set_AppInfo();
    if (m_common.Is_Integration_Installed(CMS.CommonBase.Integrations.PREFERENCE_MANAGER)) {
      if (m_common.PMRT_Connection_State() == ConnectionState.Closed) { m_common.Open_PrefManRT(); }
    }
    m_channelInc = new Copient.ChannelWebSvc(m_common);
    m_authInc = new CMS.AMS.AuthLib(m_common);
    m_custCommon = new CMS.AMS.CustomerCommon(ref m_common);
    m_custLogin = new CMS.AMS.CustomerLogin(m_common);
    CurrentRequest.Resolver.RegisterInstance<CMS.CommonBase>(m_common);
    CurrentRequest.Resolver.RegisterInstance<Copient.CommonInc>(m_legacyCommon);
    m_logger = CurrentRequest.Resolver.Resolve<ILogger>();
    m_errHandler = CurrentRequest.Resolver.Resolve<IErrorHandler>();
    m_Offers = CurrentRequest.Resolver.Resolve<IOffer>();
    m_phraseLib = CurrentRequest.Resolver.Resolve<IPhraseLib>();
  }

  // -------------------------------------------------------------------------------------------------------------------------------------

  private void Shutdown() {
    if (m_common.Is_Integration_Installed(CMS.CommonBase.Integrations.PREFERENCE_MANAGER)) {
      if (m_common.PMRT_Connection_State() != ConnectionState.Closed) { m_common.Close_PrefManRT(); }
    }
    if (m_common.LXS_Connection_State() != ConnectionState.Closed) { m_common.Close_LogixXS(); }
    if (m_common.LRT_Connection_State() != ConnectionState.Closed) { m_common.Close_LogixRT(); }


    if (m_legacyCommon.LRTadoConn.State != ConnectionState.Closed) { m_legacyCommon.Close_LogixRT(); }
    if (m_legacyCommon.LXSadoConn.State != ConnectionState.Closed) { m_legacyCommon.Close_LogixXS(); }

    m_connectorInc = null;
    m_channelInc = null;
    m_authInc = null;
    m_custCommon = null;
    m_common = null;
    m_legacyCommon = null;
    m_logger = null;
    m_errHandler = null;
    m_Offers = null;
    m_phraseLib = null;
  }

  // -------------------------------------------------------------------------------------------------------------------------------------

  /// <summary>
  /// Processes exception catch for other methods
  /// </summary>
  /// <param name="ex">Exception that occurred</param>
  /// <param name="methodName">Name of the method where the exception occurred</param>
  /// <param name="xmlWriter">XmlWriter from the method where the exception occurred</param>
  /// <param name="stringWriter">StringWriter from the method where the exception occurred</param>
  private void ProcessException(Exception ex, string methodName, ref XmlWriter xmlWriter, ref StringWriter stringWriter) {
    m_connectorInc.Init_ResponseXML(methodName, ref stringWriter, ref xmlWriter);
    if (ex is CMSException) {
      CMSException.StatusCodes statusCode = CMSException.StatusCodes.GENERAL_ERROR;
      if (ex.Data != null) {
        statusCode = (CMSException.StatusCodes)ex.Data["StatusCode"];
      }
      m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, statusCode, ex.Message, false);
    }
    else {
      m_common.Error_Processor(ex.ToString());
      m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, CMS.CMSException.StatusCodes.APPLICATION_EXCEPTION, m_phraseLib.Lookup("term.errorprocessingseelog"), false);
      m_common.Write_Log("An error occurred while processing - please see the error log!");
    }
  }

  // -------------------------------------------------------------------------------------------------------------------------------------

  private bool IsAuthTokenValid(XmlWriter xmlWriter, string methodName, string authToken, long customerPK) {
    string msg = String.Empty;
    if (customerPK < 0) {
      msg=m_phraseLib.Detokenize("term.authtokennotvalid", authToken);
      m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, CMS.CMSException.StatusCodes.INVALID_CUSTOMER_AUTH, msg, false);
      m_logger.WriteInfo(msg); 
    }
    else if (customerPK == 0) {
      msg =m_phraseLib.Detokenize("term.authtokenexpired", authToken);
      m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, CMS.CMSException.StatusCodes.EXPIRED_CUSTOMER_AUTH, msg, false);
      m_logger.WriteInfo(msg); 
    }
    return customerPK > 0;
  }

  // -------------------------------------------------------------------------------------------------------------------------------------

  /// <summary>
  /// Retrieves CMS.AMS.Customer from database based on customer details XML argument and CustomerPK
  /// </summary>
  /// <param name="customerDetails">XML with customer's details to retrieve from database</param>
  /// <param name="customerPK">Number used to identify the customer</param>
  /// <returns>Loaded CMS.AMS.Customer from database that matches the CustomerPK in the argument</returns>
  private CMS.AMS.Customer ParseCustomerDetails(XmlDocument customerDetails, long customerPK) {

    XmlNode detailsNode = customerDetails.SelectSingleNode("//Customer");
    CMS.AMS.Customer customerToRet = new CMS.AMS.Customer(m_common, customerPK);

    if (detailsNode != null) {

      // fill customer with data
      customerToRet.Prefix = ParseNodeValue(detailsNode, "Prefix");
      customerToRet.FirstName = ParseNodeValue(detailsNode, "FirstName");
      customerToRet.MiddleName = ParseNodeValue(detailsNode, "MiddleName");
      customerToRet.LastName = ParseNodeValue(detailsNode, "LastName");
      customerToRet.Suffix = ParseNodeValue(detailsNode, "Suffix");
      customerToRet.Address = ParseNodeValue(detailsNode, "Address");
      customerToRet.City = ParseNodeValue(detailsNode, "City");
      customerToRet.State = ParseNodeValue(detailsNode, "State");
      customerToRet.Zip = ParseNodeValue(detailsNode, "Zip");
      customerToRet.Country = ParseNodeValue(detailsNode, "Country");
      customerToRet.Phone = ParseNodeValue(detailsNode, "Phone");
      customerToRet.MobilePhone = ParseNodeValue(detailsNode, "MobilePhone");
      customerToRet.Email = ParseNodeValue(detailsNode, "Email");
      //customerToRet.DateOfBirth = DateTime.Parse(ParseNodeValue(detailsNode, "DateOfBirth"));
      customerToRet.DateOfBirth = ParseNodeValue(detailsNode, "DateOfBirth");
      customerToRet.AltIDOptOut = Int32.Parse(ParseNodeValue(detailsNode, "AltIDOptOut"));
    }

    return customerToRet;
  }

  // -------------------------------------------------------------------------------------------------------------------------------------

  /// <summary>
  /// Parses XML document contains OfferID numbers and returns them in a list
  /// </summary>
  /// <param name="offerIDsXML">XML containing one or more offer ID in ChannelGetChannelOfferDetailsRequest.xsd format</param>
  /// <returns>List of long integers representing offer IDs</returns>
  private List<long> ParseOfferIDs(XmlDocument offerIDsXML) {
    List<long> offerIDsList = new List<long>();

    XmlNodeList nodeList = offerIDsXML.SelectNodes("//OfferID");

    foreach (XmlNode node in nodeList) {
      long parsedValue = 0;

      long.TryParse(node.InnerText, out parsedValue);
      offerIDsList.Add(parsedValue);
    }

    return offerIDsList;
  }

  // -------------------------------------------------------------------------------------------------------------------------------------

  /// <summary>
  /// Parses XmlNode based on name passed in as argument
  /// </summary>
  /// <param name="parentNode">XML node to parse</param>
  /// <param name="childNodeName">Name of the node to parse</param>
  /// <returns>Value of the node parsed from arguments</returns>
  private string ParseNodeValue(XmlNode parentNode, string childNodeName) {

    string nodeValue = "";
    XmlNode tempNode;

    try {
      tempNode = parentNode.SelectSingleNode(childNodeName);

      if (tempNode != null) {
        nodeValue = tempNode.InnerText;
      }
    }
    catch {
      nodeValue = "";
    }

    return nodeValue;
  }

  // -------------------------------------------------------------------------------------------------------------------------------------

  private void FetchLogonIdentifierTypes(ref XmlWriter Writer) {
    List<Copient.ChannelWebSvc.IdentifierTypesRec> IdentifierList;

    Writer.WriteStartElement("IdentifierTypes");
    IdentifierList = m_channelInc.Get_LogonIdentifierTypes(Copient.ChannelWebSvc.IDENTIFIER_USAGE.LOGON, CHANNELID);
    foreach (Copient.ChannelWebSvc.IdentifierTypesRec IdentifierRec in IdentifierList) {
      Writer.WriteStartElement("IdentifierType");
      Writer.WriteElementString("TypeID", IdentifierRec.TypeID.ToString());
      Writer.WriteElementString("TypeName", IdentifierRec.TypeName);
      Writer.WriteElementString("MaxFieldLength", IdentifierRec.MaxFieldLength.ToString());
      Writer.WriteElementString("ExtIDType", IdentifierRec.ExtIDType);
      Writer.WriteEndElement();  // IdentifierType
    }
    Writer.WriteEndElement(); // IdentifierTypes
  }

  // -------------------------------------------------------------------------------------------------------------------------------------

  private void FetchCustomerDetails(ref XmlWriter Writer, Int64 CustomerPK) {

    CMS.AMS.Customer Customer;

    m_common.Write_Log("Inside FetchCustomerDetails - CustomerPK=" + CustomerPK.ToString());
    Customer = new CMS.AMS.Customer(m_common, CustomerPK);
    m_common.Write_Log("Customer.CustomerPK=" + Customer.CustomerPK.ToString());

    Writer.WriteStartElement("Customer");
    Writer.WriteElementString("Prefix", Customer.Prefix);
    Writer.WriteElementString("FirstName", Customer.FirstName);
    Writer.WriteElementString("MiddleName", Customer.MiddleName);
    Writer.WriteElementString("LastName", Customer.LastName);
    Writer.WriteElementString("Suffix", Customer.Suffix);
    Writer.WriteElementString("Address", Customer.Address);
    Writer.WriteElementString("City", Customer.City);
    Writer.WriteElementString("State", Customer.State);
    Writer.WriteElementString("Zip", Customer.Zip);
    Writer.WriteElementString("Phone", Customer.Phone);
    Writer.WriteElementString("MobilePhone", Customer.MobilePhone);
    Writer.WriteElementString("Country", Customer.Country);
    Writer.WriteElementString("Email", Customer.Email);
    Writer.WriteElementString("Employee", Customer.Employee.ToString());
    Writer.WriteElementString("EmployeeID", Customer.EmployeeID);
    Writer.WriteEndElement(); // Customer
  }

  // -------------------------------------------------------------------------------------------------------------------------------------

  private void FetchOfferList(ref XmlWriter Writer, int ChannelID, Int64 CustomerPK, int PageNum, string GUID, string AuthToken) {

    Copient.ChannelWebSvc.OfferListRec Offers;
    Offers = m_channelInc.GetChannelOfferList(ChannelID, CustomerPK, PageNum);
    if (Offers.OfferList.Count > 0) {
      Writer.WriteStartElement("Offers");
      foreach (Copient.ChannelWebSvc.ChannelOfferRec OfferRec in Offers.OfferList) {
        Writer.WriteStartElement("Offer");
        Writer.WriteElementString("OfferID", OfferRec.OfferID.ToString());
        Writer.WriteElementString("OfferStartDate", OfferRec.OfferStartDate.ToString());
        Writer.WriteElementString("OfferEndDate", OfferRec.OfferEndDate.ToString());
        Writer.WriteElementString("OfferOptable", OfferRec.IsOfferOptable.ToString().ToLower());
        Writer.WriteElementString("OfferOptedIn", OfferRec.IsOfferOptedIn.ToString().ToLower());
        Writer.WriteElementString("MediaTypeID", OfferRec.MediaTypeID.ToString());
        Writer.WriteElementString("MediaFormatID", OfferRec.MediaFormatID.ToString());
        Writer.WriteElementString("MediaData", OfferRec.MediaData);
        Writer.WriteElementString("OfferDetailLink", "channel.asmx/GetChannelOfferDetails?GUID=" + GUID + "&AuthToken=" + AuthToken + "&OfferIDsXML=<GetChannelOfferDetails><OfferIDs><OfferID>" + OfferRec.OfferID.ToString() + "</OfferID></OfferIDs></GetChannelOfferDetails>");
        Writer.WriteEndElement();
      }
      Writer.WriteStartElement("Navigation");
      Writer.WriteElementString("TotalNumOffers", Offers.TotalNumOffers.ToString());
      Writer.WriteElementString("TotalNumPages", Offers.TotalPages.ToString());
      Writer.WriteElementString("PageStartOfferNum", Offers.PageStartRec.ToString());
      Writer.WriteElementString("PageEndOfferNum", Offers.PageEndRec.ToString());
      Writer.WriteElementString("CurrentPageNum", Offers.PageNum.ToString());
      if (Offers.PageEndRec < Offers.TotalNumOffers) {
        Writer.WriteElementString("NextPage", "channel.asmx/GetOfferList?GUID=" + GUID + "&AuthToken=" + AuthToken + "&PageNum=" + (PageNum + 1).ToString());
      }
      if (PageNum > 1) {
        Writer.WriteElementString("PrevPage", "channel.asmx/GetOfferList?GUID=" + GUID + "&AuthToken=" + AuthToken + "&PageNum=" + (PageNum - 1).ToString());
      }
      Writer.WriteEndElement();
      Writer.WriteEndElement();
    }
  }

  // -------------------------------------------------------------------------------------------------------------------------------------

  /// <summary>
  /// Generates XML for channel offer details using validated offer IDs. 
  /// </summary>
  /// <param name="writer">XmlWriter used to generate XML</param>
  /// <param name="xmlInput">XML containing offer IDs verified for validity</param>
  /// <param name="channelID">Channel ID verified for validity</param>
  private void FetchChannelOfferDetails(ref XmlWriter writer, XmlDocument xmlInput, int channelID) {

    writer.WriteStartElement("Offers");

    // get OfferIDs from input XML
    foreach (long offerID in ParseOfferIDs(xmlInput)) {

      writer.WriteStartElement("Offer");
      writer.WriteElementString("OfferID", offerID != 0 ? offerID.ToString() : "Invalid OfferID");

      if (offerID != 0) {
        writer.WriteStartElement("Media");

        // get MediaAssetRec structures from database
        foreach (ChannelWebSvc.MediaAssetRec mediaAsset in m_channelInc.GetChannelOfferDetails(offerID, channelID)) {
          writer.WriteStartElement("MediaAsset");
          writer.WriteElementString("MediaTypeID", mediaAsset.MediaTypeID.ToString());
          writer.WriteElementString("MediaType", mediaAsset.MediaType);
          writer.WriteElementString("MediaFormatID", mediaAsset.MediaFormatID.ToString());
          writer.WriteElementString("MediaFormat", mediaAsset.MediaFormat);
          writer.WriteElementString("MediaData", mediaAsset.MediaData);
          writer.WriteEndElement();  // MediaAsset
        }
        writer.WriteEndElement();  // Media  
      }
      writer.WriteEndElement();  // Offer
    }
    writer.WriteEndElement();  // Offers
  }

  // -------------------------------------------------------------------------------------------------------------------------------------

  private void Close_ResponseXML(ref XmlWriter xmlWriter, ref StringWriter stringWriter, ref XmlDocument xmlResponse) {
    m_connectorInc.Close_ResponseXML(ref xmlWriter);
    xmlWriter.Flush();
    xmlWriter.Close();
    xmlResponse.LoadXml(stringWriter.ToString());
  }

  // -------------------------------------------------------------------------------------------------------------------------------------

  private void GetXMLResponse(ref XmlWriter writer, string methodName, CMSException.StatusCodes statusCode, string statusDescription, bool success)
  {
    m_connectorInc.Generate_Status_XML(ref writer, methodName, statusCode, statusDescription, success);
    m_logger.WriteInfo(statusDescription); 
  }
  private void WriteToActivityLog(int ActivityType, long OfferId, long UserID, string strMessage)
  {
    if (MyCommon.LRTadoConn.State == ConnectionState.Closed)
      MyCommon.Open_LogixRT();
    MyCommon.Activity_Log(ActivityType, OfferId, UserID, strMessage);
    MyCommon.Close_LogixRT();
  }

}
