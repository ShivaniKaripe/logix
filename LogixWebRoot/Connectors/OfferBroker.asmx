<%@ WebService Language="C#" Class="OfferBrokerService" %>

using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Text;
using System.Web;
using System.Web.Services;
using System.Web.Services.Protocols;
using System.Web.Script.Serialization;
using System.Web.Script.Services;
using System.Xml;
using Copient;

[WebService(Namespace = "http://ams.ncr.com/OfferBroker/")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]

// To allow this Web Service to be called from script, using ASP.NET AJAX 
 [System.Web.Script.Services.ScriptService]

public class OfferBrokerService : System.Web.Services.WebService
{
  private CommonInc m_common;
  private BrokerChannel m_broker;
  private ConnectorInc m_connInc;
  
  private string m_logfile  = "OfferBrokerWSLog." + DateTime.Now.ToString("yyyyMMdd") + ".txt";
  private const int BROKER_CONNECTOR_ID = 57;
  
  public enum StatusCodes
  {
    SUCCESS = 0,
    INVALID_GUID = 1,
    INVALID_PARAMETER = 2,
    APPLICATION_EXCEPTION = 9999 
  }

  // constructor
  public OfferBrokerService()
  {
    m_common = new CommonInc();
    m_common.Open_LogixRT();
    m_common.Open_LogixXS();
    
    m_broker = new BrokerChannel(ref m_common);
    m_connInc = new ConnectorInc();
  }

  // destructor
  ~ OfferBrokerService()
  {
    if (m_common.LRTadoConn.State != ConnectionState.Closed) m_common.Close_LogixRT();
    if (m_common.LXSadoConn.State != ConnectionState.Closed) m_common.Close_LogixXS();
    
    m_common = null;
    m_broker = null;
  }

  [WebMethod]
  [ScriptMethod(ResponseFormat=ResponseFormat.Json)]
  public string  GetAllOffers(string accessToken, string merchantId)
  {
    return new JavaScriptSerializer().Serialize(GetOffers(accessToken, merchantId, "GetAllOffers"));
  }

  
  [WebMethod]
  [ScriptMethod(ResponseFormat=ResponseFormat.Json)]
  public string GetChangedOffers(string accessToken, string merchantId)
  {
    return new JavaScriptSerializer().Serialize(GetOffers(accessToken, merchantId, "GetChangedOffers"));
  }

  
  [WebMethod]
  [ScriptMethod(ResponseFormat=ResponseFormat.Json)]
  public string  GetAllLocationGroups(string accessToken)
  {
    return new JavaScriptSerializer().Serialize(GetLocationGroups(accessToken, "GetAllLocationGroups"));
  }

  [WebMethod]
  [ScriptMethod(ResponseFormat=ResponseFormat.Json)]
  public string GetChangedLocationGroups(string accessToken)
  {
    return new JavaScriptSerializer().Serialize(GetLocationGroups(accessToken, "GetChangedLocationGroups"));
  }

  [WebMethod]
  [ScriptMethod(ResponseFormat=ResponseFormat.Json)]
  public string GetAllLocations(string accessToken)
  {
    return new JavaScriptSerializer().Serialize(GetLocations(accessToken, "GetAllLocations"));
  }

    
  [WebMethod]
  [ScriptMethod(ResponseFormat=ResponseFormat.Json)]
  public string GetChangedLocations(string accessToken)
  {
    return new JavaScriptSerializer().Serialize(GetLocations(accessToken, "GetChangedLocations"));
  }
  
  
  [WebMethod]
  [ScriptMethod(ResponseFormat=ResponseFormat.Json)]
  public string GetIdentifiers(string accessToken)
  {
    List<BrokerChannel.CardTypeRec> idTypeRecs = null;
    
    if (IsValidGUID(accessToken, "GetIdentifiers"))
    {
      idTypeRecs = m_broker.GetCardTypes();
    }

    return new JavaScriptSerializer().Serialize(idTypeRecs);
  }


  [WebMethod]
  [ScriptMethod(ResponseFormat=ResponseFormat.Json)]
  public string GetBanners(string accessToken)
  {
    List<BrokerChannel.BannerRec> bannerRecs = null;
    
    if (IsValidGUID(accessToken, "GetBanners"))
    {
      bannerRecs = m_broker.GetBanners();
    }
    
    return new JavaScriptSerializer().Serialize(bannerRecs);
  }

    
  [WebMethod]
  public bool SetCoalitionData(string accessToken, string coalitionData) 
  {
    bool didSet = false;

    try 
    {
      if (IsValidGUID(accessToken, "SetCoalitionData"))
      {
        List<BrokerChannel.SubChannelRec> channelRecs = DeserializeJson<List<BrokerChannel.SubChannelRec>>(coalitionData);
        m_broker.SetSubChannels(channelRecs);
        didSet = true;      
      }
    }
    catch (SoapException soapEx) 
    {
      throw soapEx; 
    }
    catch (Exception ex)
    {
      throw new SoapException(ex.ToString(), 
                              SoapException.ClientFaultCode, 
                              Context.Request.Url.AbsoluteUri,
                              CreateSoapExceptionDetail(StatusCodes.APPLICATION_EXCEPTION, null));      
    }
    
    return didSet;
    
  }

  [WebMethod]
  [ScriptMethod(ResponseFormat=ResponseFormat.Json)]
  public bool AcknowledgeOfferUpdates(string accessToken, string jsonOffers)
  {
    bool wasAcked = false;
    List<BrokerChannel.OfferACKRec> ackRecs = new List<BrokerChannel.OfferACKRec>();
    
    if (!string.IsNullOrWhiteSpace(jsonOffers)) 
    {
      try
      {
        if (IsValidGUID(accessToken, "AcknowledgeOfferUpdates"))
        {
          List<OfferUpdateResponse> ackResponses = new JavaScriptSerializer().Deserialize<List<OfferUpdateResponse>>(jsonOffers);
          foreach (OfferUpdateResponse r in ackResponses) 
          {
            if (r.offerState == OfferState.Updated) 
            {
              ackRecs.Add(new BrokerChannel.OfferACKRec() { OfferID = r.offerId,
                                                          DeployLevel = r.offerVersionNum});   
            }
            else
            {
              m_common.Write_Log(m_logfile, 
                                 string.Format("Offer {0} for merchant {1} failed to update at the broker due to the following exception: {2}", r.offerId, r.merchantId, r.message),
                                 true);
            }
          }
        
          m_broker.ACKDeployedOffers(ackRecs);
          wasAcked = true;
        }
      } 
      catch (SoapException soapEx) 
      {
        throw soapEx; 
      }
      catch(Exception ex)
      {
        m_common.Write_Log(m_logfile, 
                           string.Format("Failed to Acknowledge response due to the following exception: {0}", ex.Message),
                           true);
      }  
    }
    
    //m_broker.ACKDeployedOffers();
    return wasAcked;
  }

  
  private T DeserializeJson<T>(string data)
  {
    object retObj = default(T);
    
    try
    {
      retObj = new JavaScriptSerializer().Deserialize<T>(data);
      retObj = retObj ?? default(T);
    }
    catch (Exception ex)
    {
      Dictionary<string, string> tokens = new Dictionary<string,string>();
      retObj = default(T);
      tokens.Add("Format", new JavaScriptSerializer().Serialize(retObj));
      tokens.Add("ErrorMessage", ex.ToString());
      
      SoapException se = new SoapException("Parameter value error.", SoapException.ClientFaultCode, Context.Request.Url.AbsoluteUri,
                                           CreateSoapExceptionDetail(StatusCodes.INVALID_PARAMETER, tokens));
      throw se;
    }  
    
    return (T) retObj;
    
  }
  
  private List<Copient.BrokerChannel.BrokerOfferRec> GetOffers(string accessToken, string merchantId, string methodName) 
  {
    List<BrokerChannel.BrokerOfferRec> brokerOfferRecs = null;
    
    if (IsValidGUID(accessToken, methodName))
    {
      m_broker.MerchantID = merchantId;
      brokerOfferRecs = (methodName=="GetChangedOffers") ? m_broker.GetDeployableOffers() : m_broker.IPLOffers();
    }
    
    return brokerOfferRecs;
  }

  
  private List<BrokerChannel.LocationGroupRec> GetLocationGroups(string accessToken, string methodName) 
  {
    List<BrokerChannel.LocationGroupRec> lgRecs = null;
    
    if (IsValidGUID(accessToken, methodName))
    {
      lgRecs = (methodName=="GetChangedLocationGroups") ? m_broker.GetDeployableLocationGroups() : m_broker.IPLLocationGroups();
    }
    
    return lgRecs;
  }
  

  private List<BrokerChannel.LocationRec> GetLocations(string accessToken, string methodName) 
  {
    List<Copient.BrokerChannel.LocationRec> locRecs = null;
    
    if (IsValidGUID(accessToken, methodName))
    {
      locRecs = (methodName=="GetChangedLocations") ? m_broker.GetDeployableLocations() : m_broker.IPLLocations();
    }

    return locRecs;
  }
  
    
  private bool IsValidGUID(string accessToken, string methodName)
  {
    bool valid = false;
    StringBuilder msgBuf = new StringBuilder();
    
    try {
      valid = m_connInc.IsValidConnectorGUID(ref m_common, BROKER_CONNECTOR_ID, accessToken);
    } catch {
      valid = false; 
    }

    // log the call
    try {
      msgBuf.Append((valid) ? "Validated call to " :  "Invalid call to ");
      msgBuf.Append(methodName + " from access token: " + accessToken);
      msgBuf.Append(" and IP: " + HttpContext.Current.Request.UserHostAddress);
      m_common.Write_Log(m_logfile, msgBuf.ToString(), true);
    } catch {
      // ignore
    }
        

    if (!valid)
    {
      Dictionary<string, string> tokens = new Dictionary<string,string>();
      tokens.Add("AccessToken", accessToken);
      tokens.Add("MethodName", methodName);
      
      SoapException se = new SoapException("AccessToken " + accessToken + " is invalid.", SoapException.ClientFaultCode, Context.Request.Url.AbsoluteUri,
                                           CreateSoapExceptionDetail(StatusCodes.INVALID_GUID, tokens));
      throw se;
    }

    return valid; 
  }  

  private XmlNode CreateSoapExceptionDetail(StatusCodes errorType, Dictionary<string, string> errorTokens) 
  {
    System.Xml.XmlNode detailsChild;
    
    // Build the detail element of the SOAP fault.
    System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
    System.Xml.XmlNode node = doc.CreateNode(XmlNodeType.Element, SoapException.DetailElementName.Name, SoapException.DetailElementName.Namespace);

    // Build specific details for the SoapException.
    // Add first child of detail XML element.
    System.Xml.XmlNode details = doc.CreateNode(XmlNodeType.Element, "OfferBrokerException", "http://ams.ncr.com/OfferBroker/");
    XmlAttribute attr = doc.CreateAttribute("code");
    attr.Value = errorType.ToString();
    details.Attributes.Append(attr);
    
    if (errorTokens != null)
    {
      foreach (string key in errorTokens.Keys)
      {
        detailsChild = doc.CreateNode(XmlNodeType.Element, key, "http://ams.ncr.com/OfferBroker/");
        detailsChild.InnerText = errorTokens[key];
        details.AppendChild(detailsChild);
      }
    }
            
    // Append the two child elements to the detail node.
    node.AppendChild(details);
    
    return node;
  }

}

  internal enum OfferState 
  {
      Failed = -1,
      Unchanged = 0,
      Updated = 1
  }

  internal class OfferUpdateResponse
  {
      public long offerId { get; set; }
      public string merchantId { get; set; }
      public OfferState offerState { get; set; }
      public int offerVersionNum { get; set; }
      public string message { get; set; }
  }
