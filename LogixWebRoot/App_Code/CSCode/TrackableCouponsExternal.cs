using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Web;
using System.Web.Services;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using CMS;
using CMS.AMS;
using CMS.AMS.Contract;
using CMS.AMS.Models;
using CMS.Contract;
using CMS.DB;

/// <summary>
/// Summary description for TrackableCouponsExternal
/// </summary>

[WebService(Namespace = "http://www.copienttech.com/TrackableCouponsExternal/")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
public class TrackableCouponsExternal : WebService
{
    public enum StatusCodes
    {
    }

    private CMS.AMS.Common m_common;
    private ConnectorInc m_connectorInc;
    private ILogger m_logger;
    private IPhraseLib m_phraseLib;
    private IErrorHandler m_errHandler;
    private ITrackableCouponService m_trackableCoupon;
    private IDBAccess m_dbAccess;
	private CMS.AMS.AuthLib m_authInc = null;

    private int m_UpdateCap = 100;
    private int m_CreateCap = 100;
    private int m_LookupCap = 1000;

    /// <summary>
    /// Adds a trackable coupon to the database and returns its current status.
    /// </summary>
    /// <param name="requestXml">XML containing create request that follows TrackableCouponsExternalCreateCouponsRequest.xsd</param>
    /// <returns>XML containing create response that follows TrackableCouponsExternalCreateCouponsResponse.xsd</returns>
    [WebMethod]
    public XmlDocument CreateCoupons(string GUID, string requestXml)
    {
        const string methodName = "CreateCoupons";
        XmlWriter xmlWriter = null;
        StringWriter strWriter = null;
        var xmlResponse = new XmlDocument();
        XmlDocument xmlInput = null;
		string errMsg = "";
	
        try
        {
            Startup();
            m_connectorInc.Init_ResponseXML(ref strWriter, ref xmlWriter);
			if (!m_authInc.Is_Valid_GUID(GUID))
			{
				m_logger.WriteError( m_phraseLib.Detokenize("term.guidnotvalid", GUID, m_common.Get_AppInfo().AppName));
				m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, CMS.CMSException.StatusCodes.INVALID_GUID, m_phraseLib.Detokenize("term.guidnotvalid", GUID, m_common.Get_AppInfo().AppName), false);
			}
			else if (!m_connectorInc.ConvertStringToXML(requestXml, ref xmlInput))
			{
				m_logger.WriteError(m_phraseLib.Detokenize("term.invalidparameterxmldocument", "requestXml"));
				m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, CMS.CMSException.StatusCodes.INVALID_XML_DOCUMENT, m_phraseLib.Detokenize("term.invalidparameterxmldocument", "requestXml"), false);
			}
			else if (!m_connectorInc.IsValidXmlDocument("TrackableCouponsExternalCreateCouponsRequest.xsd", xmlInput, out errMsg))
			{
				m_logger.WriteError(m_phraseLib.Detokenize("term.invalidxmlconformxsd", "requestXml", "TrackableCouponsExternalCreateCouponsRequest.xsd: " + errMsg));
				m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, CMS.CMSException.StatusCodes.INVALID_XML_DOCUMENT, m_phraseLib.Detokenize("term.invalidxmlconformxsd", "requestXml", "TrackableCouponsExternalCreateCouponsRequest.xsd: " + errMsg), false);
			}			
            else
            {
                var couponList = ConvertXmlToCreateCouponListRequest(requestXml, "Coupon");
                if (RequestItemsAreUnderCap(xmlWriter, methodName, m_CreateCap, couponList.Count))
                {
                    var response = m_trackableCoupon.AddTrackableCouponsExternal(couponList);
                    new XmlSerializer(response.Result.GetType(), new XmlRootAttribute(methodName)).Serialize(xmlWriter, response.Result);
                }
            }
        }
        catch (Exception ex)
        {
            //MsgOut.PrintLine(ex.StackTrace);
            ProcessException(ex, methodName, xmlWriter, strWriter);
        }
        CloseResponseXml(xmlWriter, strWriter, xmlResponse);
        Shutdown();

        return xmlResponse;
    }

    private List<TrackableCouponCreateRequest> ConvertXmlToCreateCouponListRequest(string xml, string root)
    {
        var doc = XDocument.Parse(xml);

        return doc.Descendants(root).Select(d =>
        new TrackableCouponCreateRequest
        {
            CouponCode = (string)d.Element("CouponCode").Value,
            InitialUses = Convert.ToByte(d.Element("NumUses").Value),
            RemainingUses =  Convert.ToInt16(d.Element("NumUses").Value),
            ExtProgramId = d.Element("ExtProgramID").Value
        }).ToList();
    }

    /// <summary>
    /// Queries trackable coupon from database and returns its current status.
    /// </summary>
    /// <param name="requestXml">XML containing query request that follows TrackableCouponsExternalLookupCouponsRequest.xsd</param>
    /// <returns>XML containing query response that follows TrackableCouponsExternalLookupCouponsResponse.xsd</returns>
    [WebMethod]
    public XmlDocument LookupCoupons(string GUID, string requestXml)
    {
        const string methodName = "LookupCoupons";
        XmlWriter xmlWriter = null;
        StringWriter strWriter = null;
        var xmlResponse = new XmlDocument();
        XmlDocument xmlInput = null;
		string errMsg = "";
		
        try
        {
            Startup();
            m_connectorInc.Init_ResponseXML(ref strWriter, ref xmlWriter);
			if (!m_authInc.Is_Valid_GUID(GUID))
			{
				m_logger.WriteError( m_phraseLib.Detokenize("term.guidnotvalid", GUID, m_common.Get_AppInfo().AppName));
				m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, CMS.CMSException.StatusCodes.INVALID_GUID, m_phraseLib.Detokenize("term.guidnotvalid", GUID, m_common.Get_AppInfo().AppName), false);
			}
			else if (!m_connectorInc.ConvertStringToXML(requestXml, ref xmlInput))
			{
				m_logger.WriteError(m_phraseLib.Detokenize("term.invalidparameterxmldocument", "requestXml"));
				m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, CMS.CMSException.StatusCodes.INVALID_XML_DOCUMENT, m_phraseLib.Detokenize("term.invalidparameterxmldocument", "requestXml"), false);
			}
			else if (!m_connectorInc.IsValidXmlDocument("TrackableCouponsExternalLookupCouponsRequest.xsd", xmlInput, out errMsg))
			{
				m_logger.WriteError(m_phraseLib.Detokenize("term.invalidxmlconformxsd", "requestXml", "TrackableCouponsExternalLookupCouponsRequest.xsd: " + errMsg));
				m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, CMS.CMSException.StatusCodes.INVALID_XML_DOCUMENT, m_phraseLib.Detokenize("term.invalidxmlconformxsd", "requestXml", "TrackableCouponsExternalLookupCouponsRequest.xsd: " + errMsg), false);
			}			
            else
            {
                var request = ConvertXmlToCouponCodeStringList(requestXml, methodName);

                if (RequestItemsAreUnderCap(xmlWriter, methodName, m_LookupCap, request.Count))
                {
                    var response = m_trackableCoupon.QueryCouponExternal(request);
                    new XmlSerializer(response.Result.GetType(), new XmlRootAttribute(methodName)).Serialize(xmlWriter, response.Result);
                }
            }
        }
        catch (Exception ex)
        {
            ProcessException(ex, methodName, xmlWriter, strWriter);
        }
        CloseResponseXml(xmlWriter, strWriter, xmlResponse);
        Shutdown();

        return xmlResponse;
    }

    private List<string> ConvertXmlToCouponCodeStringList(string xml, string root)
    {
        var doc = XDocument.Parse(xml);

        return doc.Descendants("Coupon")
                          .Select(node => node.Attribute("code").Value)
                          .ToList(); 
    }

    /// <summary>
    /// Updates the number of uses on a trackable coupon to the database and returns its current status.
    /// </summary>
    /// <param name="requestXml">XML containing update request that follows TrackableCouponsExternalAdjustNumCouponUsesRequest.xsd</param>
    /// <returns>XML containing update response that follows TrackableCouponsExternalAdjustNumCouponUsesReponse.xsd</returns>
    [WebMethod]
    public XmlDocument AdjustNumCouponUses(string GUID, string requestXml)
    {
        const string methodName = "AdjustNumCouponUses";
        XmlWriter xmlWriter = null;
        StringWriter strWriter = null;
        var xmlResponse = new XmlDocument();
        XmlDocument xmlInput = null;
		string errMsg = "";
		
        try
        {
            Startup();
            m_connectorInc.Init_ResponseXML(ref strWriter, ref xmlWriter); 
			if (!m_authInc.Is_Valid_GUID(GUID))
			{
				m_logger.WriteError( m_phraseLib.Detokenize("term.guidnotvalid", GUID, m_common.Get_AppInfo().AppName));
				m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, CMS.CMSException.StatusCodes.INVALID_GUID, m_phraseLib.Detokenize("term.guidnotvalid", GUID, m_common.Get_AppInfo().AppName), false);
			}
			else if (!m_connectorInc.ConvertStringToXML(requestXml, ref xmlInput))
			{
				m_logger.WriteError(m_phraseLib.Detokenize("term.invalidparameterxmldocument", "requestXml"));
				m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, CMS.CMSException.StatusCodes.INVALID_XML_DOCUMENT, m_phraseLib.Detokenize("term.invalidparameterxmldocument", "requestXml"), false);
			}
			else if (!m_connectorInc.IsValidXmlDocument("TrackableCouponsExternalAdjustNumCouponUsesRequest.xsd", xmlInput, out errMsg))
			{
				m_logger.WriteError(m_phraseLib.Detokenize("term.invalidxmlconformxsd", "requestXml", "TrackableCouponsExternalAdjustNumCouponUsesRequest.xsd: " + errMsg));
				m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, CMS.CMSException.StatusCodes.INVALID_XML_DOCUMENT, m_phraseLib.Detokenize("term.invalidxmlconformxsd", "requestXml", "TrackableCouponsExternalAdjustNumCouponUsesRequest.xsd: " + errMsg), false);
			}			
            else
            {
                var request = ConvertXmlToUpdateList(requestXml, "Coupon");
                if (RequestItemsAreUnderCap(xmlWriter, methodName, m_UpdateCap, request.Count))
                {
                    var response = m_trackableCoupon.UpdateTrackableCouponRemainingUses(request);
                    new XmlSerializer(response.Result.GetType(), new XmlRootAttribute(methodName)).Serialize(xmlWriter, response.Result);
                }
            }
        }
        catch (Exception ex)
        {
            ProcessException(ex, methodName, xmlWriter, strWriter);
        }
        CloseResponseXml(xmlWriter, strWriter, xmlResponse);
        Shutdown();

        return xmlResponse;
    }

    private List<TrackableCouponAdjustNumUsesRequest> ConvertXmlToUpdateList(string xml, string root)
    {
        var doc = XDocument.Parse(xml);

        return doc.Descendants(root).Select(d =>
        new TrackableCouponAdjustNumUsesRequest
        {
            CouponCode = (string)d.Element("code"),
            RemainingUses = (short)d.Element("RemainingUses"),
        }).ToList();
    }

    /// <summary>
    /// Processes exception catch for other methods
    /// </summary>
    /// <param name="ex">Exception that occurred</param>
    /// <param name="methodName">Name of the method where the exception occurred</param>
    /// <param name="xmlWriter">XmlWriter from the method where the exception occurred</param>
    /// <param name="strWriter">StringWriter from the method where the exception occurred</param>
    private void ProcessException(Exception ex, string methodName, XmlWriter xmlWriter, StringWriter strWriter)
    {
        m_connectorInc.Init_ResponseXML(methodName, ref strWriter, ref xmlWriter);
        if (ex is CMSException)
        {
            var statusCode = CMSException.StatusCodes.GENERAL_ERROR;
            if (ex.Data != null)
            {
                statusCode = (CMSException.StatusCodes)ex.Data["StatusCode"];
            }
            GenerateStatusXml(xmlWriter, methodName, statusCode, ex.Message);
        }
        else
        {
            m_errHandler.ProcessError(ex);
            GenerateStatusXml(xmlWriter, methodName, CMSException.StatusCodes.APPLICATION_EXCEPTION, m_phraseLib.Lookup("term.errorprocessingseelog"));
            m_logger.WriteError("An error occurred while processing - please see the error log!");
        }
        m_connectorInc.Close_ResponseXML(ref xmlWriter);
    }

    private bool IsValidXmlInput(XmlDocument xmlInput, XmlWriter xmlWriter, string methodName, string requestXml)
    {
        var isValid = false;

        if (!m_connectorInc.ConvertStringToXML(requestXml, ref xmlInput))
        {
            GenerateStatusXml(xmlWriter, methodName, CMSException.StatusCodes.INVALID_XML_DOCUMENT,
              m_phraseLib.Detokenize("term.invalidparameterxmldocument", "requestXml"));
        }
        else
        {
            string errMsg;
            var xsdFileName = "TrackableCouponsExternal" + methodName + "Request.xsd";
            if (!m_connectorInc.IsValidXmlDocument(xsdFileName, xmlInput, out errMsg))
            {
                GenerateStatusXml(xmlWriter, methodName, CMSException.StatusCodes.INVALID_XML_DOCUMENT,
                  m_phraseLib.Detokenize("term.invalidxmlconformxsd", "requestXml", xsdFileName + ": " + errMsg));
            }
            else
            {
                isValid = true;
            }
        }
        return isValid;
    }

    private bool RequestItemsAreUnderCap(XmlWriter xmlWriter, string methodName, int cap, int itemCount)
    {
        var isValid = false;
        if (itemCount > cap)
            GenerateStatusXml(xmlWriter, methodName, CMSException.StatusCodes.FAILED_UPDATE, string.Format("Only {0} objects may be entered at a time", cap));
        else
            isValid = true;
        return isValid;
    }

    private void GenerateStatusXml(XmlWriter xmlWriter, string methodName, CMSException.StatusCodes exception, string description)
    {
        xmlWriter.WriteStartElement(methodName);
        m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, exception, description, false);
    }

    private void CloseResponseXml(XmlWriter xmlWriter, StringWriter strWriter, XmlDocument xmlResponse)
    {
        xmlWriter.WriteEndDocument();
        xmlWriter.Flush();
        xmlWriter.Close();
        xmlResponse.LoadXml(strWriter.ToString());
    }
    private void Shutdown()
    {
        if (m_common.LXS_Connection_State() != ConnectionState.Closed) { m_common.Close_LogixXS(); }
        if (m_common.LRT_Connection_State() != ConnectionState.Closed) { m_common.Close_LogixRT(); }

        m_connectorInc = null;
        m_common = null;
        m_logger = null;
        m_phraseLib = null;
        m_errHandler = null;
        m_trackableCoupon = null;
    }

    private void Startup()
    {
        CurrentRequest.Resolver.AppName = "TrackableCouponsExternalWS";
        m_common = CurrentRequest.Resolver.Resolve<CMS.AMS.Common>();
        if (m_common.LRT_Connection_State() == ConnectionState.Closed) { m_common.Open_LogixRT(); }
        if (m_common.LXS_Connection_State() == ConnectionState.Closed) { m_common.Open_LogixXS(); }
        CurrentRequest.Resolver.RegisterInstance<CommonBase>(m_common);
        m_common.Set_AppInfo();

        m_connectorInc = new ConnectorInc(m_common);
		m_authInc = new CMS.AMS.AuthLib(m_common);		
        m_logger = CurrentRequest.Resolver.Resolve<ILogger>();
        m_phraseLib = CurrentRequest.Resolver.Resolve<IPhraseLib>();
        m_errHandler = CurrentRequest.Resolver.Resolve<IErrorHandler>();
        m_trackableCoupon = CurrentRequest.Resolver.Resolve<ITrackableCouponService>();
        m_dbAccess = CurrentRequest.Resolver.Resolve<IDBAccess>();
    }

  
   

    
}


