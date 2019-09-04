using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web.Services;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using CMS;
using CMS.AMS;
using CMS.AMS.Contract;
using CMS.AMS.Models;
using CMS.Contract;
using System.Globalization;

/// <summary>
/// Internal WebService for Trackable Coupons that allows coupons to be scanned, redeemed, unlocked, and queried.
/// </summary>
[WebService(Namespace = "http://ncr.cms.ams.com/")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
// To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line.
// [System.Web.Script.Services.ScriptService]

public class TrackableCoupons : WebService
{
  // version:7.3.1.138972.Official Build (SUSDAY10202)

  private CMS.AMS.Common m_common;
  private ConnectorInc m_connectorInc;
  private ILogger m_logger;
  private IPhraseLib m_phraseLib;
  private IErrorHandler m_errHandler;
  private ITrackableCouponService m_trackableCoupon;

  /// <summary>
  /// Queries trackable coupon from database and returns its current status.
  /// </summary>
  /// <param name="requestXml">XML containing query request that follows TrackableCouponsQueryCouponRequest.xsd</param>
  /// <returns>XML containing query response that follows TrackableCouponsQueryCouponResponse.xsd</returns>
  [WebMethod]
  public XmlDocument QueryCoupon(string requestXml)
  {
    const string methodName = "QueryCoupon";
    XmlWriter xmlWriter = null;
    StringWriter strWriter = null;
    var xmlResponse = new XmlDocument();
    XmlDocument xmlInput = null;

    try
    {
      Startup();
      m_connectorInc.Init_ResponseXML(ref strWriter, ref xmlWriter);

      if (IsValidXmlInput(xmlInput, xmlWriter, methodName, requestXml))
      {
        var request = m_connectorInc.ConvertXmlToList(requestXml, "Code");
        var response = m_trackableCoupon.QueryCoupon(request);
		  
		  XmlSerializerCache.Create(response.Result.GetType(), new XmlRootAttribute(methodName)).Serialize(xmlWriter, response.Result);

        //m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, CMSException.StatusCodes.SUCCESS, m_phraseLib.Lookup("term.success"), true);
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

  /// <summary>
  /// Redeems trackable coupon by subtracting the number of times the coupon was used from the remaining uses.
  /// </summary>
  /// <param name="requestXml">XML containing redeem request that follows TrackableCouponsRedeemCouponRequest.xsd</param>
  /// <returns>XML containing redeem response that follows TrackableCouponsRedeemCouponResponse.xsd</returns>
  [WebMethod]
  public XmlDocument RedeemCoupon(string requestXml)
  {
    const string methodName = "RedeemCoupon";
    XmlWriter xmlWriter = null;
    StringWriter strWriter = null;
    var xmlResponse = new XmlDocument();
    XmlDocument xmlInput = null;

    try
    {
      Startup();
      m_connectorInc.Init_ResponseXML(ref strWriter, ref xmlWriter);

      if (IsValidXmlInput(xmlInput, xmlWriter, methodName, requestXml))
      {
        var request = ConvertXmlToRedeemList(requestXml, "Coupon");
        var response = m_trackableCoupon.RedeemCoupon(request);
		XmlSerializerCache.Create(response.Result.GetType(), new XmlRootAttribute(methodName)).Serialize(xmlWriter, response.Result);

        //m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, CMSException.StatusCodes.SUCCESS, m_phraseLib.Lookup("term.success"), true);
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

  //This is implemented to support force redeem of Trackable coupon even the Coupon key is not supplied.
  //But with the change in UE requirements, not exposing this webmethod and accomodated the same functionality in RedeemCoupon webmethod.
  //[WebMethod]
  public XmlDocument ForceRedeemCoupon(string requestXml)
  {
    const string methodName = "ForceRedeemCoupon";
    XmlWriter xmlWriter = null;
    StringWriter strWriter = null;
    var xmlResponse = new XmlDocument();
    XmlDocument xmlInput = null;

    try
    {
      Startup();
      m_connectorInc.Init_ResponseXML(ref strWriter, ref xmlWriter);

      if (IsValidXmlInput(xmlInput, xmlWriter, methodName, requestXml))
      {
        var request = ConvertXmlToRedeemList(requestXml, "Coupon");
        var response = m_trackableCoupon.ForceRedeemCoupon(request);
		XmlSerializerCache.Create(response.Result.GetType(), new XmlRootAttribute(methodName)).Serialize(xmlWriter, response.Result);
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

  /// <summary>
  /// Scans trackable coupons by locking their use and returning their status.
  /// </summary>
  /// <param name="requestXml">XML containing scan request that follows TrackableCouponsScanCouponRequest.xsd</param>
  /// <returns>XML containing scan response that follows TrackableCouponsScanCouponResponse.xsd</returns>
  [WebMethod]
  public XmlDocument ScanCoupon(string requestXml)
  {
    const string methodName = "ScanCoupon";
    XmlWriter xmlWriter = null;
    StringWriter strWriter = null;
    var xmlResponse = new XmlDocument();
    XmlDocument xmlInput = null;

    try
    {
      Startup();
      m_connectorInc.Init_ResponseXML(ref strWriter, ref xmlWriter);

      if (IsValidXmlInput(xmlInput, xmlWriter, methodName, requestXml))
      {
        var request = ConvertXmlToScanList(requestXml, "Coupon");
        var response = m_trackableCoupon.ScanCoupon(request);
        XmlSerializerCache.Create(response.Result.GetType(), new XmlRootAttribute(methodName)).Serialize(xmlWriter, response.Result);

        //m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, CMSException.StatusCodes.SUCCESS, m_phraseLib.Lookup("term.success"), true);
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

  /// <summary>
  /// Unlocks trackable coupon if it was previously locked by a scan.
  /// </summary>
  /// <param name="requestXml">XML containing unlock request that follows TrackableCouponsUnlockCouponRequest.xsd</param>
  /// <returns>XML containing unlock response that follows TrackableCouponsUnlockCouponResponse.xsd</returns>
  [WebMethod]
  public XmlDocument UnlockCoupon(string requestXml)
  {
    const string methodName = "UnlockCoupon";
    XmlWriter xmlWriter = null;
    StringWriter strWriter = null;
    var xmlResponse = new XmlDocument();
    XmlDocument xmlInput = null;

    try
    {
      Startup();
      m_connectorInc.Init_ResponseXML(ref strWriter, ref xmlWriter);

      if (IsValidXmlInput(xmlInput, xmlWriter, methodName, requestXml))
      {
        var request = m_connectorInc.ConvertXmlToList(requestXml, "Code");
        var response = m_trackableCoupon.UnlockCoupon(request);
		XmlSerializerCache.Create(response.Result.GetType(), new XmlRootAttribute(methodName)).Serialize(xmlWriter, response.Result);

        //m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, CMSException.StatusCodes.SUCCESS, m_phraseLib.Lookup("term.success"), true);
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

  private List<TrackableCouponRedeemRequest> ConvertXmlToRedeemList(string xml, string root)
  {
    var doc = XDocument.Parse(xml);

    return doc.Descendants(root).Select(d =>
    new TrackableCouponRedeemRequest
    {
      Code = (string) d.Element("Code"),
      CouponsUsed = ((int) d.Element("CouponsUsed")).ConvertToByte(),
      LockKey = Utilities.NZ((long?)d.Element("LockKey"), 0L),
      StoreId = (long) d.Element("StoreId"),
      CustomerId = Utilities.NZ((long?) d.Element("CustomerId"), 0L),
      LogixTransNum = (string) d.Element("LogixTransNum") ?? "0",
      ForceRedeem= Utilities.NZ((bool?)d.Element("ForceRedeem"),false)
    }).ToList();
  }

  private List<TrackableCouponScanRequest> ConvertXmlToScanList(string xml, string root)
  {
    var doc = XDocument.Parse(xml);

    return doc.Descendants(root).Select(d =>
    new TrackableCouponScanRequest
    {
      Code = (string) d.Element("Code"),
      LogixTransNum = (string) d.Element("LogixTransNum") ?? "0"
    }).ToList();
  }

  private void CloseResponseXml(XmlWriter xmlWriter, StringWriter strWriter, XmlDocument xmlResponse)
  {
    xmlWriter.WriteEndDocument();
    xmlWriter.Flush();
    xmlWriter.Close();
	strWriter.Flush();
	strWriter.Close();
    xmlResponse.LoadXml(strWriter.ToString());
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
      var xsdFileName = "TrackableCoupons" + methodName + "Request.xsd";
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

  private void GenerateStatusXml(XmlWriter xmlWriter, string methodName, CMSException.StatusCodes exception, string description)
  {
    xmlWriter.WriteStartElement(methodName);
    m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, exception, description, false);
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

  private void Shutdown()
  {
    m_connectorInc = null;
    m_common = null;
    m_logger = null;
    m_phraseLib = null;
    m_errHandler = null;
    m_trackableCoupon = null;
  }

  private void Startup()
  {
    CurrentRequest.Resolver.AppName = "TrackableCouponsWS";
    m_common = CurrentRequest.Resolver.Resolve<CMS.AMS.Common>();
    CurrentRequest.Resolver.RegisterInstance<CommonBase>(m_common);
    m_common.Set_AppInfo();

    m_connectorInc = new ConnectorInc(m_common);
    m_logger = CurrentRequest.Resolver.Resolve<ILogger>();
    m_phraseLib = CurrentRequest.Resolver.Resolve<IPhraseLib>();
    m_errHandler = CurrentRequest.Resolver.Resolve<IErrorHandler>();
    m_trackableCoupon = CurrentRequest.Resolver.Resolve<ITrackableCouponService>();
  }
}