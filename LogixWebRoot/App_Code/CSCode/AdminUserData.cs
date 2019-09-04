using System;
using System.Data;
using System.Xml;
using System.Xml.Serialization;
using System.IO;
using System.Web;
using System.Web.Services;
using System.Web.Services.Protocols;
using CMS;
using CMS.Contract;
using CMS.AMS;
using CMS.AMS.Contract;
using CMS.Models;
using CMS.DB;

[WebService(Namespace = "http://ncr.cms.ams.com/")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
// To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
// [System.Web.Script.Services.ScriptService]
public class AdminUserDataPopulation : System.Web.Services.WebService
{
  #region Private Variables
  private CMS.AMS.Common m_common = null;
  private ConnectorInc m_connectorInc = null;
  private ILogger m_logger = null;
  private IErrorHandler m_errHandler = null;
  private CMS.AMS.AuthLib m_authInc = null;
  private IPhraseLib m_phraseLib = null;
  private IAdminUserData m_adminUser = null;
  private IDBAccess m_dbaccess = null;
  private Copient.CommonInc m_commonInc = null;
  #endregion Private Variables

  #region Web Methods
  [WebMethod]
  public XmlDocument CreateAdminUser(string GUID, string requestXML)
  {
    string methodName = "AddAdminUser";
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
      else if (!m_connectorInc.IsValidXmlDocument("AdminUserData.xsd", xmlInput, out errMsg))
      {
        m_logger.WriteError(m_phraseLib.Detokenize("term.invalidxmlconformxsd", "requestXML", "AdminUserData.xsd: " + errMsg));
        m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, CMS.CMSException.StatusCodes.INVALID_XML_DOCUMENT, m_phraseLib.Detokenize("term.invalidxmlconformxsd", "requestXML", "AdminUserData.xsd: " + errMsg), false);
      }
      else
      {
        var adminUser = ParseAdminUserDetails(xmlInput);
        var response = m_adminUser.AddAdminUser(adminUser);
        CMS.CMSException.StatusCodes statusCode = CMS.CMSException.StatusCodes.GENERAL_ERROR;
        if (response.ResultType.CompareTo(CMS.AMS.Models.AMSResultType.Success) == 0)
        {
          statusCode = CMS.CMSException.StatusCodes.SUCCESS;
          xmlWriter.WriteStartElement("AdminUser");
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
      m_logger.WriteError("Failed to create Admin user role- please see the error log!");
      m_errHandler.ProcessError(ex);
      ProcessException(ex, methodName, ref xmlWriter, ref strWriter);
    }

    CloseResponseXml(ref xmlWriter, ref strWriter, ref xmlResponse);
    Shutdown();

    return xmlResponse;
  }
  
  [WebMethod]
  public XmlDocument UpdateAdminUserById(string GUID, int id, string requestXML)
  {
    string methodName = "UpdateAdminUserById";
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
      else if (!m_connectorInc.IsValidXmlDocument("AdminUserData.xsd", xmlInput, out errMsg))
      {
        m_logger.WriteError(m_phraseLib.Detokenize("term.invalidxmlconformxsd", "requestXML", "AdminUserData.xsd: " + errMsg));
        m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, CMS.CMSException.StatusCodes.INVALID_XML_DOCUMENT, m_phraseLib.Detokenize("term.invalidxmlconformxsd", "requestXML", "AdminUserData.xsd: " + errMsg), false);
      }
      else
      {
        var adminUser = ParseAdminUserDetails(xmlInput);
        var response = m_adminUser.UpdateAdminUserById(id, adminUser);
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
      m_logger.WriteError("Failed to update admin user- please see the error log!");
      m_errHandler.ProcessError(ex);
      ProcessException(ex, methodName, ref xmlWriter, ref strWriter);
    }

    CloseResponseXml(ref xmlWriter, ref strWriter, ref xmlResponse);
    Shutdown();

    return xmlResponse;
  }

  [WebMethod]
  public XmlDocument UpdateAdminUserByEmployeeId(string GUID, string employeeId, string requestXML)
  {
    string methodName = "UpdateAdminUserByEmployeeId";
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
      else if (!m_connectorInc.IsValidXmlDocument("AdminUserData.xsd", xmlInput, out errMsg))
      {
        m_logger.WriteError(m_phraseLib.Detokenize("term.invalidxmlconformxsd", "requestXML", "AdminUserData.xsd: " + errMsg));
        m_connectorInc.Generate_Status_XML(ref xmlWriter, methodName, CMS.CMSException.StatusCodes.INVALID_XML_DOCUMENT, m_phraseLib.Detokenize("term.invalidxmlconformxsd", "requestXML", "AdminUserData.xsd: " + errMsg), false);
      }
      else
      {
        var adminUser = ParseAdminUserDetails(xmlInput);
        var response = m_adminUser.UpdateAdminUserByEmployeeId(employeeId, adminUser);
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
      m_logger.WriteError("Failed to update admin user- please see the error log!");
      m_errHandler.ProcessError(ex);
      ProcessException(ex, methodName, ref xmlWriter, ref strWriter);
    }

    CloseResponseXml(ref xmlWriter, ref strWriter, ref xmlResponse);
    Shutdown();

    return xmlResponse;
  }

  [WebMethod]
  public XmlDocument DeleteAdminUserById(string GUID, int id)
  {
    string methodName = "DeleteAdminUserById";
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
        var response = m_adminUser.DeleteAdminUserByID(id);
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
      m_logger.WriteError("Failed to delete admin user- please see the error log!");
      m_errHandler.ProcessError(ex);
      ProcessException(ex, methodName, ref xmlWriter, ref strWriter);
    }

    CloseResponseXml(ref xmlWriter, ref strWriter, ref xmlResponse);
    Shutdown();

    return xmlResponse;
  }

  [WebMethod]
  public XmlDocument DeleteAdminUserByEmployeeId(string GUID, string employeeId)
  {
    string methodName = "DeleteAdminUserByEmployeeId";
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
        var response = m_adminUser.DeleteAdminUserByEmployeeId(employeeId);
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
      m_logger.WriteError("Failed to delete admin user- please see the error log!");
      m_errHandler.ProcessError(ex);
      ProcessException(ex, methodName, ref xmlWriter, ref strWriter);
    }

    CloseResponseXml(ref xmlWriter, ref strWriter, ref xmlResponse);
    Shutdown();

    return xmlResponse;

  }

  [WebMethod]
  public XmlDocument LookupAdminUserById(string GUID, int id)
  {
    string methodName = "LookupAdminUserById";
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
        var response = m_adminUser.LookupAdminUserById(id);
        CMS.CMSException.StatusCodes statusCode = CMS.CMSException.StatusCodes.GENERAL_ERROR;
        if (response.ResultType == CMS.AMS.Models.AMSResultType.Success)
        {
          statusCode = CMS.CMSException.StatusCodes.SUCCESS;
          FetchAdminUserDetails(ref xmlWriter, xmlInput, response.Result);
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
      m_logger.WriteError("Failed to LookUp admin user please see the error log!");
      m_errHandler.ProcessError(ex);
      ProcessException(ex, methodName, ref xmlWriter, ref strWriter);
    }

    CloseResponseXml(ref xmlWriter, ref strWriter, ref xmlResponse);
    Shutdown();

    return xmlResponse;
  }

  [WebMethod]
  public XmlDocument LookupAdminUserByEmployeeId(string GUID, string employeeId)
  {
    string methodName = "LookupAdminUserByEmployeeId";
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
        var response = m_adminUser.LookupAdminUserByEmployeeId(employeeId);
        CMS.CMSException.StatusCodes statusCode = CMS.CMSException.StatusCodes.GENERAL_ERROR;
        if (response.ResultType == CMS.AMS.Models.AMSResultType.Success)
        {
          statusCode = CMS.CMSException.StatusCodes.SUCCESS;
          FetchAdminUserDetails(ref xmlWriter, xmlInput, response.Result);
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
      m_logger.WriteError("Failed to look up admin user please see the error log!");
      m_errHandler.ProcessError(ex);
      ProcessException(ex, methodName, ref xmlWriter, ref strWriter);
    }

    CloseResponseXml(ref xmlWriter, ref strWriter, ref xmlResponse);
    Shutdown();

    return xmlResponse;
  }
  #endregion Web Methods

  #region Private Functions
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
      var xsdFileName = "AdminUser.xsd";
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

  private AdminUser ParseAdminUserDetails(XmlDocument adminUser)
  {
    XmlNode userNode = adminUser.SelectSingleNode("//AdminUser");
    AdminUser user = new AdminUser();
    if (adminUser != null)
    {

      string empid = ParseNodeValue(userNode, "EmployeeID");
      if (!String.IsNullOrWhiteSpace(empid))
        user.EmployeeId = m_commonInc.CleanString(empid);
      else
        user.EmployeeId = null;
      user.FirstName = m_commonInc.CleanString(ParseNodeValue(userNode, "FirstName"));
      user.LastName = m_commonInc.CleanString(ParseNodeValue(userNode, "LastName"));
      user.UserName = m_commonInc.CleanString(ParseNodeValue(userNode, "UserName"));
      user.Email = ParseNodeValue(userNode, "Email");
      user.AlertEmail = ParseNodeValue(userNode, "AlertEmail");
      string sLangId = ParseNodeValue(userNode, "LanguageID");

      if (!String.IsNullOrEmpty(sLangId))
        user.LanguageID = GetLanguageId(sLangId);
      else
        user.LanguageID = 1;

    }
    return user;

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
    writer.WriteElementString("EmployeeId", Convert.ToString(adminUser.EmployeeId));
    writer.WriteElementString("FirstName", adminUser.FirstName);
    writer.WriteElementString("LastName", adminUser.LastName);
    writer.WriteElementString("UserName", adminUser.UserName);
    writer.WriteElementString("Email", adminUser.Email);
    writer.WriteElementString("AlertEmail", adminUser.AlertEmail);
    writer.WriteElementString("LanguageID", Convert.ToString(adminUser.LanguageID));
    writer.WriteEndElement();
  }

  private int GetLanguageId(string javaLocalcode)
  {

    int langaugeId = 0;
    m_logger.WriteInfo("getting language id for given java local code");
    string query = "select LanguageID  from languages where JavaLocaleCode=@JavaLocaleCode and InstalledForUI=1";
    SQLParametersList paramlist = new SQLParametersList();
    paramlist.Add("@JavaLocaleCode", SqlDbType.NVarChar, 10).Value = javaLocalcode;
    DataTable dt = m_dbaccess.ExecuteQuery(DataBases.LogixRT, CommandType.Text, query, paramlist);
    if (dt.Rows.Count > 0)
    {
      langaugeId = Convert.ToInt32(dt.Rows[0]["LanguageID"]);
    }
    return langaugeId;
  }

  public void Startup()
  {
    CurrentRequest.Resolver.AppName = "AdminUserData";
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
    m_dbaccess = CurrentRequest.Resolver.Resolve<IDBAccess>();
    m_commonInc = new Copient.CommonInc();
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
  }
  #endregion Private Functions
}