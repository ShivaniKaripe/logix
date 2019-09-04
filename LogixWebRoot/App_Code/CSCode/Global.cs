using CMS;
using CMS.AMS;
using CMS.AMS.Contract;
using CMS.AMS.Security;
using CMS.Contract;
using CMS.DB;
using Ncr.Nep;
using Ncr.Nep.Web.Health;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Text;
using System.Web;
using System.Web.Http;

public class Global_asax : HttpApplication
{
  private List<ComponentHealthMonitor> componentHealthMonitors;
  private WebRequestResolverBuilder webRequestResolverBuilder;

  public override void Init()
  {
    base.Init();

    webRequestResolverBuilder = new WebRequestResolverBuilder();
    webRequestResolverBuilder.Build();
  }

  public void Application_Start(object sender, EventArgs e)
  {
    //Initialize NEP Application
    NepServiceApplication.Init();

    WebApiConfig.Register(GlobalConfiguration.Configuration);

    // Code that runs on application startup
    CMS.AMS.ResolverBuilder resolver = new ResolverBuilder();
    CurrentRequest.Resolver = resolver.GetResolver();
    CurrentRequest.Resolver.AppName = "Global.asax";
    IDBAccess m_dbaccess = CurrentRequest.Resolver.Resolve<DBAccess>();
    SQLParametersList paramlist = new SQLParametersList();
    String QueryStr = "Update Folders set MassOperationStatus = '~FNIU~'";
    m_dbaccess.ExecuteNonQuery(DataBases.LogixRT, CommandType.Text, QueryStr, paramlist);

    //Adding health monitors
    InitializeHealthMonitors();
    HealthOptions.AddHealthMonitors(componentHealthMonitors.ToArray());
  }

  public void Application_End(object sender, EventArgs e)
  {
    // Code that runs on application shutdown
  }

  public void Application_Error(object sender, EventArgs e)
  {
    // Code that runs when an unhandled error occurs
    Exception ex = Server.GetLastError();
    string appName = "";

    //Check if the exception is "Maximum request length exceeded."
    if ((ex) is System.Web.HttpException & ex.Message == "Maximum request length exceeded.")
    {
      Server.ClearError();
      Response.Clear();
      string queryString = null;
      if (Context.Request.UrlReferrer.Query.Length > 0)
      {
        queryString = "&LargeFile=true";
      }
      else
      {
        queryString = "?LargeFile=true";
      }
      Response.Redirect(Request.UrlReferrer.PathAndQuery + queryString);
    }
    else
    {
      try
      {
        appName = System.IO.Path.GetFileName(Request.PhysicalPath);
        Write_To_Error_Log(ex, appName);
      }
      catch (ArgumentException argEx)
      {
        Write_To_Error_Log(argEx, "Framework");
      }
      catch (Exception exLog)
      {
        // unable to log this error out to either a file or the event logger
      }
      //AMS-2465 Apurva's comments
      Response.Redirect("~/logix/Error.aspx", false);
      Server.ClearError();
    }

  }

  public void Session_Start(object sender, EventArgs e)
  {
    // Ensure SesssionID exists when the Application Pool Recycles
    string sessionId = Session.SessionID;
    CurrentRequest.Resolver.AppName = "Global.asax";
    ICacheData cachedata = CurrentRequest.Resolver.Resolve<ICacheData>();
    //CMS.AMS.Common common = CurrentWebRequest.Resolver.Resolve<CMS.AMS.Common>();
    //if (common.LRT_Connection_State() != ConnectionState.Open)
    //  common.Open_LogixRT();
    int SessionTimeOut = cachedata.GetSystemOption_General_ByOptionId(158).ConvertToInt32();
    if (SessionTimeOut != 0)
    {
      Session.Timeout = SessionTimeOut;
    }
  }

  public void Session_End(object sender, EventArgs e)
  {
    // Code that runs when a session ends. 
    // Note: The Session_End event is raised only when the sessionstate mode
    // is set to InProc in the Web.config file. If session mode is set to StateServer 
    // or SQLServer, the event is not raised.
  }


  protected void Application_BeginRequest(object sender, System.EventArgs e)
  {
    CurrentRequest.Resolver = webRequestResolverBuilder.GetResolver();
    CurrentRequest.CanDispose = true;
  }

  protected void Application_EndRequest(object sender, System.EventArgs e)
  {
    if (CurrentRequest.CanDispose)
      CurrentRequest.DisposeResolver();
    // webRequestResolverBuilder.Container.Dispose();
  }

  private void Write_To_Error_Log(Exception ex, string appName)
  {
    string LogDirName = "";
    StringBuilder OutBuffer = new StringBuilder();
    System.IO.DirectoryInfo d = null;

    try
    {
      d = new System.IO.DirectoryInfo(Server.MapPath("~"));
      LogDirName = d.Parent.FullName + "\\logs";

      if (System.IO.Directory.Exists(LogDirName))
      {
        OutBuffer.Append("-------------------------------------------------------------" + Environment.NewLine);
        OutBuffer.Append("Error in: " + appName + Environment.NewLine);
        OutBuffer.Append("      at: " + DateTime.Today + " " + DateTime.Now.TimeOfDay + Environment.NewLine + Environment.NewLine);
        OutBuffer.Append("The following internal error has occurred:" + Environment.NewLine + " " + Environment.NewLine);
        OutBuffer.AppendLine("Error Description: " + ex.ToString() + Environment.NewLine + Environment.NewLine);

        if (Copient.commonShared.UnauthenticatedPages().Contains(Request.CurrentExecutionFilePath.ToLower()))
        {
          Copient.Logger.Write_Log("unauthenticated-errors.log", OutBuffer.ToString(), true);
        }
        else
        {
          System.IO.File.AppendAllText(LogDirName + "\\ErrorLog." + DateTime.Now.ToString("yyyyMMdd") + ".txt", OutBuffer.ToString());
        }
      }
    }
    catch (Exception)
    {
      try
      {
        Write_To_Event_Log(ex.ToString(), appName, EventLogEntryType.Error);
      }
      catch (Exception)
      {
        // abandon effort to log this error
      }
    }
  }


  //*************************************************************
  //NAME:          WriteToEventLog
  //PURPOSE:       Write to Event Log
  //PARAMETERS:    Entry - Value to Write
  //               AppName - Name of Client Application. Needed 
  //               because before writing to event log, you must 
  //               have a named EventLog source. 
  //               EventType - Entry Type, from EventLogEntryType 
  //               Structure e.g., EventLogEntryType.Warning, 
  //               EventLogEntryType.Error
  //               LogNam1e: Name of Log (System, Application; 
  //               Security is read-only) If you 
  //               specify a non-existent log, the log will be
  //               created
  //RETURNS:       True if successful
  //*************************************************************
  private bool Write_To_Event_Log(string entry, string appName = "NCR", EventLogEntryType eventType = EventLogEntryType.Information, string logName = "Logix")
  {

    EventLog objEventLog = new EventLog();

    try
    {
      //Register the Application as an Event Source
      if (!EventLog.SourceExists(appName))
      {
        EventLog.CreateEventSource(appName, logName);
      }

      //log the entry
      objEventLog.Source = appName;
      objEventLog.WriteEntry(entry, eventType);

      return true;
    }
    catch (Exception)
    {
      return false;
    }
  }

  private void InitializeHealthMonitors()
  {
    ILoginWithNEP loginWithNEP;
    RESTServiceHelper restServiceHelper;
    componentHealthMonitors = new List<ComponentHealthMonitor>();

    //Agents HealthMonitor
    CurrentRequest.Resolver = GetResolver("AgentHealthMonitor");
    componentHealthMonitors.Add(new AgentsHealthMonitor(CurrentRequest.Resolver.Resolve<IDBAccess>()));

    //Database HealthMonitor
    CurrentRequest.Resolver = GetResolver("DatabaseHealthMonitor");
    componentHealthMonitors.Add(new DatabaseHealthMonitor(CurrentRequest.Resolver.Resolve<IDBAccess>()));

    //Diskspace HealthMonitor
    CurrentRequest.Resolver = GetResolver("DiskspaceHealthMonitor");
    componentHealthMonitors.Add(new DiskSpaceHealthMonitor(CurrentRequest.Resolver.Resolve<CMS.AMS.Common>()));

    //FolderPermissions HealthMonitor
    CurrentRequest.Resolver = GetResolver("FolderPermissionsHealthMonitor");
    componentHealthMonitors.Add(new FolderPermissionsHealthMonitor(CurrentRequest.Resolver.Resolve<CMS.AMS.Common>()));

    //Messaging HealthMonitor
    CurrentRequest.Resolver = GetResolver("MessagingHealthMonitor");
    loginWithNEP = new LoginWithNEP(CurrentRequest.Resolver.Resolve<IDBAccess>(), CurrentRequest.Resolver.Resolve<ILogger>(), CurrentRequest.Resolver.Resolve<IErrorHandler>());
    restServiceHelper = new RESTServiceHelper();
    restServiceHelper.m_LoginWithNEP = loginWithNEP;
    componentHealthMonitors.Add(new MessagingHealthMonitor(restServiceHelper, CurrentRequest.Resolver.Resolve<CMS.AMS.Common>()));

    //OCD HealthMonitor
    CurrentRequest.Resolver = GetResolver("OCDHealthMonitor");
    loginWithNEP = new LoginWithNEP(CurrentRequest.Resolver.Resolve<IDBAccess>(), CurrentRequest.Resolver.Resolve<ILogger>(), CurrentRequest.Resolver.Resolve<IErrorHandler>());
    restServiceHelper = new RESTServiceHelper();
    restServiceHelper.m_LoginWithNEP = loginWithNEP;
    componentHealthMonitors.Add(new OCDHealthMonitor(restServiceHelper, CurrentRequest.Resolver.Resolve<CMS.AMS.Common>()));
  }

  private IServiceResolver GetResolver(string name)
  {
    ResolverBuilder resolverBuilder = new ResolverBuilder();
    resolverBuilder.Build();
    IServiceResolver serviceResolver = resolverBuilder.GetResolver();
    serviceResolver.AppName = name;

    return serviceResolver;
  }
}
