using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Web.Services;
using System.Xml;
using CMS;
using CMS.AMS;
using CMS.AMS.Contract;
using CMS.AMS.Models;
using CMS.Contract;
using CMS.Models;

[WebService(Namespace = "http://ncr.cms.ams.com/")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
public class BuyerRoleDataPopulation : System.Web.Services.WebService
{
  #region Private Variables
  private CMS.AMS.Common m_common = null;
  private ConnectorInc m_connectorInc = null;
  private ILogger m_logger = null;
  private IErrorHandler m_errHandler = null;
  private CMS.AMS.AuthLib m_authInc = null;
  private IPhraseLib m_phraseLib = null;
  private IAdminUserData m_adminUser = null;
  private IBuyerRoleData m_buyerRole = null;
  private IDepartment m_DepartmentService = null;
  #endregion Private Variables

  #region Web Methods

  [WebMethod]
  public XmlDocument CreateBuyerRole(string GUID, string requestXML)
  {
    string methodName = "CreateBuyerRole";
    XmlWriter xmlWriter = null;
    StringWriter strWriter = null;
    XmlDocument xmlResponse = new XmlDocument();
    XmlDocument xmlInput = null;
    string errMsg = "";

    try
    {
      Startup();
      m_connectorInc.Init_ResponseXML(methodName, ref strWriter, ref xmlWriter);
      if (!m_authInc.Is_Valid_GUID(GUID))
      {
        m_logger.WriteError(m_phraseLib.Detokenize("term.guidnotvalid", GUID, m_common.Get_AppInfo().AppName));
        m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, CMS.CMSException.StatusCodes.INVALID_GUID, m_phraseLib.Detokenize("term.guidnotvalid", GUID, m_common.Get_AppInfo().AppName), false);
      }
      else if (!m_connectorInc.ConvertStringToXML(requestXML, ref xmlInput))
      {
        m_logger.WriteError(m_phraseLib.Detokenize("term.invalidparameterxmldocument", "requestXML"));
        m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, CMS.CMSException.StatusCodes.INVALID_XML_DOCUMENT, m_phraseLib.Detokenize("term.invalidparameterxmldocument", "requestXML"), false);
      }
      else if (!m_connectorInc.IsValidXmlDocument("BuyerRoleData.xsd", xmlInput, out errMsg))
      {
        m_logger.WriteError(m_phraseLib.Detokenize("term.invalidxmlconformxsd", "requestXML", "BuyerRoleData.xsd: " + errMsg));
        m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, CMS.CMSException.StatusCodes.INVALID_XML_DOCUMENT, m_phraseLib.Detokenize("term.invalidxmlconformxsd", "requestXML", "BuyerRoleData.xsd: " + errMsg), false);
      }
      else
      {
        var buyerRole = ParseBuyerRoleDetails(xmlInput);
        var response = m_buyerRole.CreateBuyerRole(buyerRole);
        CMS.CMSException.StatusCodes statusCode = CMS.CMSException.StatusCodes.GENERAL_ERROR;
        if (response.ResultType.CompareTo(AMSResultType.Success) == 0)
        {
          statusCode = CMS.CMSException.StatusCodes.SUCCESS;
          xmlWriter.WriteStartElement("Buyer");
          xmlWriter.WriteElementString("ID", Convert.ToString(response.Result));
          xmlWriter.WriteEndElement();
          m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, statusCode, response.MessageString, true);
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
      m_logger.WriteError("Failed to create Buyer roler- please see the error log!");
      m_errHandler.ProcessError(ex);
      ProcessException(ex, methodName, ref xmlWriter, ref strWriter);
    }

    CloseResponseXml(ref xmlWriter, ref strWriter, ref xmlResponse);
    Shutdown();

    return xmlResponse;
  }

  [WebMethod]
  public XmlDocument UpdateBuyerRoleByInternalId(string GUID, int id, string requestXML)
  {
    string methodName = "UpdateBuyerRoleByInternalId";
    XmlWriter xmlWriter = null;
    StringWriter strWriter = null;
    XmlDocument xmlResponse = new XmlDocument();
    XmlDocument xmlInput = null;
    string errMsg = "";

    try
    {
      Startup();
      m_connectorInc.Init_ResponseXML(methodName, ref strWriter, ref xmlWriter);
      if (!m_authInc.Is_Valid_GUID(GUID))
      {
        m_logger.WriteError(m_phraseLib.Detokenize("term.guidnotvalid", GUID, m_common.Get_AppInfo().AppName));
        m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, CMS.CMSException.StatusCodes.INVALID_GUID, m_phraseLib.Detokenize("term.guidnotvalid", GUID, m_common.Get_AppInfo().AppName), false);
      }
      else if (!m_connectorInc.ConvertStringToXML(requestXML, ref xmlInput))
      {
        m_logger.WriteError(m_phraseLib.Detokenize("term.invalidparameterxmldocument", "requestXML"));
        m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, CMS.CMSException.StatusCodes.INVALID_XML_DOCUMENT, m_phraseLib.Detokenize("term.invalidparameterxmldocument", "requestXML"), false);
      }
      else if (!m_connectorInc.IsValidXmlDocument("BuyerRoleData.xsd", xmlInput, out errMsg))
      {
        m_logger.WriteError(m_phraseLib.Detokenize("term.invalidxmlconformxsd", "requestXML", "BuyerRoleData.xsd: " + errMsg));
        m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, CMS.CMSException.StatusCodes.INVALID_XML_DOCUMENT, m_phraseLib.Detokenize("term.invalidxmlconformxsd", "requestXML", "BuyerRoleData.xsd: " + errMsg), false);
      }
      else
      {
        var buyerRole = ParseBuyerRoleDetails(xmlInput);
        var response = m_buyerRole.UpdateBuyerRoleByInternalId(id, buyerRole);
        CMS.CMSException.StatusCodes statusCode = CMS.CMSException.StatusCodes.GENERAL_ERROR;
        if (response.ResultType == CMS.AMS.Models.AMSResultType.Success)
        {
          statusCode = CMS.CMSException.StatusCodes.SUCCESS;
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
      m_logger.WriteError("Failed to update buyer role- please see the error log!");
      m_errHandler.ProcessError(ex);
      ProcessException(ex, methodName, ref xmlWriter, ref strWriter);
    }

    CloseResponseXml(ref xmlWriter, ref strWriter, ref xmlResponse);
    Shutdown();

    return xmlResponse;
  }

  [WebMethod]
  public XmlDocument UpdateBuyerRoleByExternalId(string GUID, string externalId, string requestXML)
  {
    string methodName = "UpdateBuyerRoleByExternalId";
    XmlWriter xmlWriter = null;
    StringWriter strWriter = null;
    XmlDocument xmlResponse = new XmlDocument();
    XmlDocument xmlInput = null;
    string errMsg = "";

    try
    {
      Startup();
      m_connectorInc.Init_ResponseXML(methodName, ref strWriter, ref xmlWriter);
      if (!m_authInc.Is_Valid_GUID(GUID))
      {
        m_logger.WriteError(m_phraseLib.Detokenize("term.guidnotvalid", GUID, m_common.Get_AppInfo().AppName));
        m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, CMS.CMSException.StatusCodes.INVALID_GUID, m_phraseLib.Detokenize("term.guidnotvalid", GUID, m_common.Get_AppInfo().AppName), false);
      }
      else if (!m_connectorInc.ConvertStringToXML(requestXML, ref xmlInput))
      {
        m_logger.WriteError(m_phraseLib.Detokenize("term.invalidparameterxmldocument", "requestXML"));
        m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, CMS.CMSException.StatusCodes.INVALID_XML_DOCUMENT, m_phraseLib.Detokenize("term.invalidparameterxmldocument", "requestXML"), false);
      }
      else if (!m_connectorInc.IsValidXmlDocument("BuyerRoleData.xsd", xmlInput, out errMsg))
      {
        m_logger.WriteError(m_phraseLib.Detokenize("term.invalidxmlconformxsd", "requestXML", "BuyerRoleData.xsd: " + errMsg));
        m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, CMS.CMSException.StatusCodes.INVALID_XML_DOCUMENT, m_phraseLib.Detokenize("term.invalidxmlconformxsd", "CustomerDetailXML", "BuyerRoleData.xsd: " + errMsg), false);
      }
      else
      {
        var buyerRole = ParseBuyerRoleDetails(xmlInput);
        var response = m_buyerRole.UpdateBuyerRoleByExternalId(externalId, buyerRole);
        CMS.CMSException.StatusCodes statusCode = CMS.CMSException.StatusCodes.GENERAL_ERROR;
        if (response.ResultType == CMS.AMS.Models.AMSResultType.Success)
        {
          statusCode = CMS.CMSException.StatusCodes.SUCCESS;
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
      m_logger.WriteError("Failed to update buyer role- please see the error log!");
      m_errHandler.ProcessError(ex);
      ProcessException(ex, methodName, ref xmlWriter, ref strWriter);
    }

    CloseResponseXml(ref xmlWriter, ref strWriter, ref xmlResponse);
    Shutdown();

    return xmlResponse;
  }

  [WebMethod]
  public XmlDocument DeleteBuyerRoleById(string GUID, int id)
  {
    string methodName = "DeleteBuyerRoleById";
    XmlWriter xmlWriter = null;
    StringWriter strWriter = null;
    XmlDocument xmlResponse = new XmlDocument();

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
        var response = m_buyerRole.DeleteBuyerRoleByInternalId(id);
        CMS.CMSException.StatusCodes statusCode = CMS.CMSException.StatusCodes.GENERAL_ERROR;
        if (response.ResultType == CMS.AMS.Models.AMSResultType.Success)
        {
          statusCode = CMS.CMSException.StatusCodes.SUCCESS;
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
      m_logger.WriteError("Failed to delete buyer role- please see the error log!");
      m_errHandler.ProcessError(ex);
      ProcessException(ex, methodName, ref xmlWriter, ref strWriter);
    }
    CloseResponseXml(ref xmlWriter, ref strWriter, ref xmlResponse);
    Shutdown();
    return xmlResponse;
  }

  [WebMethod]
  public XmlDocument DeleteBuyerRoleByExternalId(string GUID, string externalId)
  {
    string methodName = "DeleteBuyerRoleByExternalId";
    XmlWriter xmlWriter = null;
    StringWriter strWriter = null;
    XmlDocument xmlResponse = new XmlDocument();

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
        var response = m_buyerRole.DeleteBuyerRoleByExternalId(externalId);
        CMS.CMSException.StatusCodes statusCode = CMS.CMSException.StatusCodes.GENERAL_ERROR;
        if (response.ResultType == CMS.AMS.Models.AMSResultType.Success)
        {
          statusCode = CMS.CMSException.StatusCodes.SUCCESS;
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
      m_logger.WriteError("Failed to delete buyer role- please see the error log!");
      m_errHandler.ProcessError(ex);
      ProcessException(ex, methodName, ref xmlWriter, ref strWriter);
    }
    CloseResponseXml(ref xmlWriter, ref strWriter, ref xmlResponse);
    Shutdown();
    return xmlResponse;
  }

  [WebMethod]
  public XmlDocument LookupBuyerRoleById(string GUID, int id)
  {
    string methodName = "LookupBuyerRoleById";
    XmlWriter xmlWriter = null;
    StringWriter strWriter = null;
    XmlDocument xmlResponse = new XmlDocument();

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
        var response = m_buyerRole.LookupBuyerRoleByInternalId(id);
        CMS.CMSException.StatusCodes statusCode = CMS.CMSException.StatusCodes.GENERAL_ERROR;
        if (response.ResultType == CMS.AMS.Models.AMSResultType.Success)
        {
          statusCode = CMS.CMSException.StatusCodes.SUCCESS;
          FetchBuyerRoleDetails(ref xmlWriter, response.Result);
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
      m_logger.WriteError("Failed to look up buyer role- please see the error log!");
      m_errHandler.ProcessError(ex);
      ProcessException(ex, methodName, ref xmlWriter, ref strWriter);
    }

    CloseResponseXml(ref xmlWriter, ref strWriter, ref xmlResponse);
    Shutdown();
    return xmlResponse;
  }

  [WebMethod]
  public XmlDocument LookupBuyerRoleByExternalId(string GUID, string externalBuyerId)
  {
    string methodName = "LookupBuyerRoleByExternalId";
    XmlWriter xmlWriter = null;
    StringWriter strWriter = null;
    XmlDocument xmlResponse = new XmlDocument();

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
        var response = m_buyerRole.LookupBuyerRoleByExternalId(externalBuyerId);
        CMS.CMSException.StatusCodes statusCode = CMS.CMSException.StatusCodes.GENERAL_ERROR;
        if (response.ResultType == CMS.AMS.Models.AMSResultType.Success)
        {
          statusCode = CMS.CMSException.StatusCodes.SUCCESS;
          FetchBuyerRoleDetails(ref xmlWriter, response.Result);
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
      m_logger.WriteError("Failed to look up buyer role- please see the error log!");
      m_errHandler.ProcessError(ex);
      ProcessException(ex, methodName, ref xmlWriter, ref strWriter);
    }

    CloseResponseXml(ref xmlWriter, ref strWriter, ref xmlResponse);
    Shutdown();
    return xmlResponse;
  }
  #endregion Web Methods

  #region Private Functions

  private void FetchBuyerRoleDetails(ref XmlWriter writer, Buyer buyer)
  {
    writer.WriteStartElement("Buyer");
    writer.WriteElementString("ExternalBuyerID", buyer.ExternalID);
    writer.WriteStartElement("AdminUsers");
    if (buyer.AdminUser != null && buyer.AdminUser.Count > 0)
    {
      foreach (var adminUser in buyer.AdminUser)
        writer.WriteElementString("Username", adminUser.UserName);

    }
    writer.WriteEndElement();
    writer.WriteStartElement("Departments");
    foreach (var department in buyer.Department.Departments)
    {
      writer.WriteElementString("DeptID", department.ExternalID);
    }
    writer.WriteEndElement();
    writer.WriteEndElement();
  }

  private void CloseResponseXml(ref XmlWriter xmlWriter, ref StringWriter strWriter, ref XmlDocument xmlResponse)
  {
    m_connectorInc.Close_ResponseXML(ref xmlWriter);
    xmlWriter.Flush();
    xmlWriter.Close();
    xmlResponse.LoadXml(strWriter.ToString());
  }

  private bool IsValidXmlInput(XmlDocument xmlInput, string methodName, XmlWriter xmlWriter, string requestXml)
  {
    bool isValid = false;

    if (!m_connectorInc.ConvertStringToXML(requestXml, ref xmlInput))
    {
      m_logger.WriteError(m_phraseLib.Detokenize("term.invalidparameterxmldocument", "requestXml"));
      GenerateStatusXml(xmlWriter, methodName, CMSException.StatusCodes.INVALID_XML_DOCUMENT,
      m_phraseLib.Detokenize("term.invalidparameterxmldocument", "requestXml"));
    }
    else
    {
      string errMsg;
      var xsdFileName = "BuyerRoleData.xsd";
      if (!m_connectorInc.IsValidXmlDocument(xsdFileName, xmlInput, out errMsg))
      {
        m_logger.WriteError(m_phraseLib.Detokenize("term.invalidxmlconformxsd", "requestXml", xsdFileName + ": " + errMsg));
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

  private Buyer ParseBuyerRoleDetails(XmlDocument buyerRole)
  {
    XmlNode buyerNode = buyerRole.SelectSingleNode("//Buyer");
    Buyer buyer = new Buyer();
    if (buyerNode != null)
    {
      buyer.ExternalID = ParseNodeValue(buyerNode, "ExternalBuyerID");

      List<string> userNames = ParseRootNode(buyerNode.SelectSingleNode("//AdminUsers"));
      if (userNames != null && userNames.Count != 0)
      {
        buyer.AdminUser = new List<AdminUser>();
        buyer.IsUsernameProvided = true;
        userNames.ForEach(u => buyer.AdminUser.Add(m_adminUser.LookupAdminUserByUsername(u).Result));
      }

      List<string> departments = ParseRootNode(buyerNode.SelectSingleNode("//Departments"));
      if (departments != null && departments.Count != 0)
      {
        buyer.Department = new BuyerDepts();
        buyer.IsDepartmentIDProvided = true;
        List<PHNode> deptsbyExternalID = new List<PHNode>();
        departments.ForEach(d => deptsbyExternalID.Add(m_DepartmentService.GetDepartmentByExternalID(d.ConvertToString()).Result));
        buyer.Department.Departments = deptsbyExternalID;
      }
    }
    return buyer;
  }

  private List<string> ParseRootNode(XmlNode node)
  {
    List<string> nodeList = new List<string>();
    if (node != null)
    {
      for (int i = 0; i < node.ChildNodes.Count; i++)
      {
        nodeList.Add(node.ChildNodes[i].InnerText);
      }
    }
    return nodeList;
  }

  /// <summary>
  /// Parses XmlNode based on name passed in as argument
  /// </summary>
  /// <param name="parentNode">XML node to parse</param>
  /// <param name="childNodeName">Name of the node to parse</param>
  /// <returns>Value of the node parsed from arguments</returns>
  private string ParseNodeValue(XmlNode parentNode, string childNodeName)
  {
    string nodeValue = "";
    XmlNode tempNode;

    try
    {
      tempNode = parentNode.SelectSingleNode(childNodeName);
      if (tempNode != null)
      {
        nodeValue = tempNode.InnerText;
      }
    }
    catch
    {
      nodeValue = "";
    }
    return nodeValue;
  }

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

  private void FetchAdminUserDetails(ref XmlWriter writer, XmlDocument xmlInput, AdminUser adminUser)
  {
    writer.WriteStartElement("AdminUser");
    writer.WriteElementString("ID", Convert.ToString(adminUser.ID));
    writer.WriteElementString("UserName", adminUser.UserName);
    writer.WriteElementString("EmployeeId", Convert.ToString(adminUser.EmployeeId));
    writer.WriteElementString("UserName", adminUser.UserName);
    writer.WriteElementString("FirstName", adminUser.FirstName);
    writer.WriteElementString("LastName", adminUser.LastName);
    writer.WriteElementString("Email", adminUser.Email);
    writer.WriteElementString("AlertEmail", adminUser.AlertEmail);
    writer.WriteElementString("LanguageID", Convert.ToString(adminUser.LanguageID));
    writer.WriteEndElement();
  }

  public void Startup()
  {
    CurrentRequest.Resolver.AppName = "BuyerRoleData";
    m_common = CurrentRequest.Resolver.Resolve<CMS.AMS.Common>();
    if (m_common.LRT_Connection_State() == ConnectionState.Closed) { m_common.Open_LogixRT(); }
    if (m_common.LXS_Connection_State() == ConnectionState.Closed) { m_common.Open_LogixXS(); }
    CurrentRequest.Resolver.RegisterInstance<CommonBase>(m_common);

    m_common.Set_AppInfo();

    m_connectorInc = new ConnectorInc(m_common);
    m_logger = CurrentRequest.Resolver.Resolve<ILogger>();
    m_phraseLib = CurrentRequest.Resolver.Resolve<IPhraseLib>();
    m_errHandler = CurrentRequest.Resolver.Resolve<IErrorHandler>();
    m_authInc = new CMS.AMS.AuthLib(m_common);
    m_adminUser = CurrentRequest.Resolver.Resolve<IAdminUserData>();
    m_buyerRole = CurrentRequest.Resolver.Resolve<IBuyerRoleData>();
    m_DepartmentService = CurrentRequest.Resolver.Resolve<IDepartment>();
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
    m_buyerRole = null;
  }
  #endregion Private Functions
}