using System;
using System.Collections.Generic;
using System.Data;
using System.Web.UI.WebControls;
using CMS.AMS;
using CMS.AMS.Contract;
using CMS.AMS.Models;
using System.Collections;
using System.Linq;
using System.Drawing;
using System.Web;
using CMS;
using System.Web.Services;
using Newtonsoft.Json;
using System.Web.UI.HtmlControls;
using System.IO;
using System.ServiceModel.Web;
using System.Net;
using System.Net.Http;

public partial class UEServerHealthSummary : AuthenticatedUI
{
	#region Private Variables

	private IHealth m_health;
	private IActivityLogService m_ActivityLogService;
	//private NodeHealth serverHealth;
    private HealthSummaryList serverHealthList;

	#endregion Private Variables

	#region Properties

	public int LocationID
	{
		get
		{
			return Convert.ToInt32(hdnLocationID.Value);
		}
		set
		{
			hdnLocationID.Value = value.ToString();
		}
	}

	#endregion Properties

	#region Protected Functions

	protected override void OnInit(EventArgs e)
	{
		AppName = "UEServerHealthSummary.aspx";
		base.OnInit(e);
	}

	protected void Page_Load(object sender, EventArgs e)
	{
		m_health = CurrentRequest.Resolver.Resolve<IHealth>();
		m_ActivityLogService = CurrentRequest.Resolver.Resolve<IActivityLogService>();
		ucServerHealthTabs.LanguageID = LanguageID;
		ucServerHealthTabs.Title = PhraseLib.Lookup("term.health", LanguageID) + " " + PhraseLib.Lookup("term.for", LanguageID).ToLower();
		hdnURL.Value = m_health.HealthServiceURL;

		((UE_MasterPage)this.Master).Tab_Name = "8_10";

		if (!IsPostBack)
		{
			ucServerHealthTabs.EnginesInstalled = m_health.GetInstalledEngines(LanguageID);
			ucServerHealthTabs.SetInfoMessage("", false);
			LocationID = -1;
			LocationID = m_health.GetServerLocationID();

			if (!m_health.UseHealthService)
				Response.Redirect("/logix/UE/store-health-UE.aspx?filterhealth=2");

			if (LocationID <= 0)
			{
				ucServerHealthTabs.SetInfoMessage(PhraseLib.Lookup("serverhealth.novalidserver", LanguageID), true);
				linkConfig.Visible = false;
                loadmoreajaxloader.InnerHtml = loadmoreajaxloader1.InnerHtml = "<center>" + PhraseLib.Lookup("term.data", LanguageID) + " " + PhraseLib.Lookup("term.unavailable", LanguageID).ToLower() + "." + "</center>";
			}
			else
			{
				linkConfig.Text = PhraseLib.Lookup("term.serverconfiguration", LanguageID);
				linkConfig.NavigateUrl = "../store-edit.aspx?LocationID=" + hdnLocationID.Value;

				if (!string.IsNullOrEmpty(Request.QueryString["errorMessage"]))
					ucServerHealthTabs.SetInfoMessage(Request.QueryString["errorMessage"], false);

				PopulateServerSummary();			
			}

			if (!string.IsNullOrEmpty(Request.QueryString["message"]))
				ucServerHealthTabs.SetInfoMessage(Request.QueryString["message"], false);
		}
	}

	#endregion Protected Functions

	#region Private Functions

	private Tuple<HealthSummaryList, DateTime> GetStoredJSONResponse()
	{
		DateTime HealthOn = DateTime.MinValue;
        HealthSummaryList nodeHealth = null;
		try
		{
			string responseString = "";
			string fileName = LoggerExtension.logger.LogPath.TrimEnd(new char[] { '/', '\\' }) + "/" + RESTServiceList.ServerHealthService + ".txt";

			if (File.Exists(fileName))
			{
				responseString = File.ReadAllText(fileName);
				HealthOn = Convert.ToDateTime(responseString.Substring(0, responseString.IndexOf('|')));
				responseString = responseString.Substring(responseString.IndexOf('|') + 1);

				if (responseString != string.Empty)
					nodeHealth = JsonConvert.DeserializeObject<HealthSummaryList>(responseString);
			}
		}
		catch (Exception ex)
		{
			Logger.WriteError("Failed to retrieve stored Server Health:" + ex.ToString());
			HealthOn = DateTime.MinValue;
			nodeHealth = null;
		}
		return new Tuple<HealthSummaryList, DateTime>(nodeHealth, HealthOn);
	}

	private void PopulateServerSummary()
	{
		string messagePhraseTerm = m_health.CheckForIPL(LocationID);
		if (messagePhraseTerm != string.Empty)
		{
			ucServerHealthTabs.SetInfoMessage(Copient.PhraseLib.Lookup(messagePhraseTerm, LanguageID), true);
		}


        //Bind Health
        List<KeyValuePair<string, string>> headers = new List<KeyValuePair<string, string>>();
        IRestServiceHelper RESTServiceHelper = CurrentRequest.Resolver.Resolve<IRestServiceHelper>();
        KeyValuePair<HealthSummaryList, string> result = RESTServiceHelper.CallService<HealthSummaryList>(RESTServiceList.ServerHealthService, (hdnURL.Value != string.Empty ? hdnURL.Value + "/offerdistributionstatus" : ""), LanguageID, HttpMethod.Get, "", true, headers);
        string errorMessage = result.Value;
        serverHealthList = result.Key;

        if (errorMessage == string.Empty)
		{
			BindData(serverHealthList);
			BindWarnings(hdnURL.Value + "/allerrors?&report=true&offset=0&pagesize=" + hdnPageSize1.Value, gvWarnings, hdnPageSize1, hdnPageCount1, loadmoreajaxloader1);
			BindFiles(hdnURL.Value + "/allfiles?offset=0&pagesize=" + hdnPageSize.Value, gvFiles, hdnPageSize, hdnPageCount, loadmoreajaxloader);
		}
		else
		{
			ucServerHealthTabs.SetInfoMessage(errorMessage, true, false);

			var storedServerHealth = GetStoredJSONResponse();
            serverHealthList = storedServerHealth.Item1;
			DateTime HealthOn = storedServerHealth.Item2;

			if (serverHealthList != null && HealthOn != null && HealthOn != DateTime.MinValue)
			{
				ucServerHealthTabs.SetInfoMessage(PhraseLib.Lookup("serverhealth.msg", LanguageID) + " " + HealthOn.ToString("dd/MM/yyyy") + ".", true, true);
				BindData(serverHealthList);
			}
			loadmoreajaxloader.InnerHtml = loadmoreajaxloader1.InnerHtml = "<center>" + PhraseLib.Lookup("term.data", LanguageID) + " " + PhraseLib.Lookup("term.unavailable", LanguageID).ToLower() + "." + "</center>";
		}
	}

	private void BindData(HealthSummaryList serverHealthList)
	{
        if (serverHealthList.healthSet.Count == 1)
            BindSingleNode(serverHealthList);
        else
        {
            long Node1LastFetch = serverHealthList.healthSet[0].LastFetch;
            long Node2LastFetch = serverHealthList.healthSet[1].LastFetch;
            long Node1LastIPL = serverHealthList.healthSet[0].LastIPL;
            long Node2LastIPL = serverHealthList.healthSet[1].LastIPL;

            BindFetchData(Node1LastFetch, Node2LastFetch, serverHealthList);

            BindIPLData(Node1LastIPL, Node2LastIPL, serverHealthList);
        }
    }
    private void BindSingleNode(HealthSummaryList serverHealthList)
    {
        lblLastIPL.Text = (serverHealthList.healthSet[0].LastIPL == 0) ? "-" : serverHealthList.healthSet[0].LastIPL.ConvertToLocalDateTime().ToString();
        lblLastUpdateOffer.Text = (serverHealthList.healthSet[0].LastFetch == 0) ? "-" : serverHealthList.healthSet[0].LastFetch.ConvertToLocalDateTime().ToString();
        lblFetchServerName.Text = Server.HtmlEncode((serverHealthList.healthSet[0].LastFetch == 0) ? "-" : serverHealthList.healthSet[0].NodeName);
        lblFetchServerIP.Text = Server.HtmlEncode((serverHealthList.healthSet[0].LastFetch == 0) ? "-" : serverHealthList.healthSet[0].NodeIP);
        lblIPLServerName.Text = Server.HtmlEncode((serverHealthList.healthSet[0].LastIPL == 0) ? "-" : serverHealthList.healthSet[0].NodeName);
        lblIPLServerIP.Text = Server.HtmlEncode((serverHealthList.healthSet[0].LastIPL == 0) ? "-" : serverHealthList.healthSet[0].NodeIP);
    }
    private void BindFetchData(long Node1LastFetch, long Node2LastFetch, HealthSummaryList serverHealthList)
    {
        if (Node1LastFetch > 0 && Node2LastFetch > 0)
        {
            if ((DateTime.Compare(Node1LastFetch.ConvertToLocalDateTime(), Node2LastFetch.ConvertToLocalDateTime()) < 0))
            {
                lblLastUpdateOffer.Text = Node2LastFetch.ConvertToLocalDateTime().ToString();
                lblFetchServerName.Text = Server.HtmlEncode(serverHealthList.healthSet[1].NodeName);
                lblFetchServerIP.Text = Server.HtmlEncode(serverHealthList.healthSet[1].NodeIP);
            }
            else
            {
                lblLastUpdateOffer.Text = Node1LastFetch.ConvertToLocalDateTime().ToString();
                lblFetchServerName.Text = Server.HtmlEncode(serverHealthList.healthSet[0].NodeName);
                lblFetchServerIP.Text = Server.HtmlEncode(serverHealthList.healthSet[0].NodeIP);
            }
        }
        else if (Node1LastFetch == 0 && Node2LastFetch == 0)
        {
            lblLastUpdateOffer.Text = "-";
            lblFetchServerName.Text = "-";
            lblFetchServerIP.Text = "-";
        }
        else
        {
            lblLastUpdateOffer.Text = (Node1LastFetch > 0) ? Node1LastFetch.ConvertToLocalDateTime().ToString() : Node2LastFetch.ConvertToLocalDateTime().ToString();
            lblFetchServerName.Text = Server.HtmlEncode((Node1LastFetch > 0) ? serverHealthList.healthSet[0].NodeName : serverHealthList.healthSet[1].NodeName);
            lblFetchServerIP.Text = Server.HtmlEncode((Node1LastFetch > 0) ? serverHealthList.healthSet[0].NodeIP : serverHealthList.healthSet[1].NodeIP);
        }
    }
    private void BindIPLData(long Node1LastIPL, long Node2LastIPL, HealthSummaryList serverHealthList)
    {
        if (Node1LastIPL > 0 && Node2LastIPL > 0)
        {
            if ((DateTime.Compare(Node1LastIPL.ConvertToLocalDateTime(), Node2LastIPL.ConvertToLocalDateTime()) < 0))
            {
                lblLastIPL.Text = Node2LastIPL.ConvertToLocalDateTime().ToString();
                lblIPLServerName.Text = Server.HtmlEncode(serverHealthList.healthSet[1].NodeName);
                lblIPLServerIP.Text = Server.HtmlEncode(serverHealthList.healthSet[1].NodeIP);
            }
            else
            {
                lblLastIPL.Text = Node1LastIPL.ConvertToLocalDateTime().ToString();
                lblIPLServerName.Text = Server.HtmlEncode(serverHealthList.healthSet[0].NodeName);
                lblIPLServerIP.Text = Server.HtmlEncode(serverHealthList.healthSet[0].NodeIP);
            }
        }
        else if (Node1LastIPL == 0 && Node2LastIPL == 0)
        {
            lblLastIPL.Text = "-";
            lblIPLServerName.Text = "-";
            lblIPLServerIP.Text = "-";
        }
        else
        {
            lblLastIPL.Text = (Node1LastIPL > 0) ? Node1LastIPL.ConvertToLocalDateTime().ToString() : Node2LastIPL.ConvertToLocalDateTime().ToString();
            lblIPLServerName.Text = Server.HtmlEncode((Node1LastIPL > 0) ? serverHealthList.healthSet[0].NodeName : serverHealthList.healthSet[1].NodeName);
            lblIPLServerIP.Text = Server.HtmlEncode((Node1LastIPL > 0) ? serverHealthList.healthSet[0].NodeIP : serverHealthList.healthSet[1].NodeIP);
        }
    }

	private void BindWarnings(string url, GridView gvFiles, HtmlInputHidden pageSize, HtmlInputHidden pageCount, HtmlGenericControl loader)
	{

		try
		{
			gvWarnings.Columns[0].HeaderText = PhraseLib.Lookup("term.description", LanguageID);
			gvWarnings.Columns[1].HeaderText = PhraseLib.Lookup("term.duration", LanguageID);
      List<KeyValuePair<string, string>> headers = new List<KeyValuePair<string, string>>();
      IRestServiceHelper RESTServiceHelper = CurrentRequest.Resolver.Resolve<IRestServiceHelper>();
			KeyValuePair<NodeHealthSummary, string> result = RESTServiceHelper.CallService<NodeHealthSummary>(RESTServiceList.ServerHealthService, url, LanguageID, HttpMethod.Get, String.Empty, false, headers);
			NodeHealthSummary summary = result.Key;
			string error = result.Value;

			pageCount.Value = "0";

			if (error != string.Empty)
			{
				loader.InnerHtml = "<center>" + error + "</center>";
			}
			else if (summary.RowCount > 0)
			{

				gvWarnings.DataSource = GetErrors(summary, LanguageID);
				gvWarnings.DataBind();

				if (summary.RowCount <= pageSize.Value.ConvertToInt16())
				{
					pageCount.Value = "1";
				}
				else
				{
					decimal value = (summary.RowCount.ConvertToDecimal() / pageSize.Value.ConvertToDecimal());
					pageCount.Value = Math.Ceiling(value).ToString();
				}

				if (pageCount.Value == "1")
				{
					loader.InnerHtml = "<center>" + PhraseLib.Lookup("term.nomorerecords", LanguageID) + "</center>";
				}
			}

			if (pageCount.Value == "0")
			{
				loader.InnerHtml = "<center>" + PhraseLib.Lookup("term.norecords", LanguageID) + "</center>";
			}
		}
		catch (Exception ex)
		{
			loader.InnerHtml = "<center>" + ex.Message + "</center>";
			Logger.WriteError(ex.ToString());
		}
	}

	#endregion Private Functions


	private void BindFiles(string url, GridView gvFiles, HtmlInputHidden pageSize, HtmlInputHidden pageCount, HtmlGenericControl loader)
	{
		try
		{
      List<KeyValuePair<string, string>> headers = new List<KeyValuePair<string, string>>();
      IRestServiceHelper RESTServiceHelper = CurrentRequest.Resolver.Resolve<IRestServiceHelper>();
			KeyValuePair<PendingFiles, string> result = RESTServiceHelper.CallService<PendingFiles>(RESTServiceList.ServerHealthService, url, LanguageID, HttpMethod.Get, String.Empty, false, headers);
			PendingFiles files = result.Key;
			string error = result.Value;

			pageCount.Value = "0";

			if (error != string.Empty)
			{
				loader.InnerHtml = "<center>" + error + "</center>";
			}
			else if (files.Files.Count > 0)
			{
				DataTable dtFiles = new DataTable();
				dtFiles.Columns.AddRange(new DataColumn[] { new DataColumn("RecordText"), new DataColumn("RecordLink"), new DataColumn("FileName"), new DataColumn("Age"), new DataColumn("Created"), new DataColumn("Path") });

				foreach (PendingFile file in files.Files)
				{
					KeyValuePair<string,string> record = GetRecordText(file.FileType, file.ID, LanguageID);
					dtFiles.Rows.Add(record.Value, record.Key, file.FileName, file.CreatedOn.ConvertToAge(), file.CreatedOn.ConvertToLocalDateTime().ToString(), files.FilesPath);
				}

				gvFiles.DataSource = dtFiles;
				gvFiles.DataBind();

				if (files.RowCount <= pageSize.Value.ConvertToInt16())
				{
					pageCount.Value = "1";
				}
				else
				{
					decimal value = (files.RowCount.ConvertToDecimal() / pageSize.Value.ConvertToDecimal());
					pageCount.Value = Math.Ceiling(value).ToString();
				}

			}

			if (pageCount.Value == "0")
			{
				loader.InnerHtml = "<center>" + PhraseLib.Lookup("term.norecords", LanguageID) + "</center>";
			}
			else if (pageCount.Value == "1")
			{
				loader.InnerHtml = "<center>" + PhraseLib.Lookup("term.nomorerecords", LanguageID) + "</center>";
			}

		}
		catch (Exception ex)
		{
			loader.InnerHtml = "<center>" + ex.Message + "</center>";
			Logger.WriteError(ex.ToString());
		}
	}

	private static KeyValuePair<string,string> GetRecordText(string recordType, string ID, int LanguageID)
	{
		string RecordLink = "", RecordText ="";

		switch (recordType.ToLower())
		{
			case PendingFileTypeConstants.Offer:
				RecordText = PhraseLibExtension.PhraseLib.Lookup("term.offer", LanguageID) + " #" + ID;
				RecordLink = "UEoffer-sum.aspx?OfferID=" + ID;
				break;

			case PendingFileTypeConstants.ProductGroup:
				RecordText = PhraseLibExtension.PhraseLib.Lookup("term.productgroup", LanguageID) + " #" + ID;
				RecordLink = "../pgroup-edit.aspx?ProductGroupID=" + ID;
				break;

			case PendingFileTypeConstants.CustomerGroup:
				RecordText = PhraseLibExtension.PhraseLib.Lookup("term.customergroup", LanguageID) + " #" + ID;
				RecordLink = "../cgroup-edit.aspx?CustomerGroupID=" + ID;
				break;

			case PendingFileTypeConstants.GeneralSettings:
				RecordText = PhraseLibExtension.PhraseLib.Lookup("term.generalsettings", LanguageID);
				RecordLink = "../configuration.aspx";
				break;

			default:
				RecordText = PhraseLibExtension.PhraseLib.Lookup("term.unknown", LanguageID);
				RecordLink = "";
				break;
		}

		return new KeyValuePair<string,string>(RecordLink,RecordText);
	}

	[WebMethod]
	[System.ServiceModel.Web.WebInvoke(Method = "POST")]
	public static string GetFiles(string URL, int LanguageID)
	{
		PendingFiles files = null;
    List<KeyValuePair<string, string>> headers = new List<KeyValuePair<string, string>>();
    IRestServiceHelper RESTServiceHelper = CurrentRequest.Resolver.Resolve<IRestServiceHelper>();
		KeyValuePair<PendingFiles, string> result = RESTServiceHelper.CallService<PendingFiles>(RESTServiceList.ServerHealthService, URL, LanguageID, HttpMethod.Get, String.Empty, false, headers);
		files = result.Key;
		string errorMessage = result.Value;

		var pendingFiles = new List<object>();

		if (errorMessage == "")
		{
			if (files != null && files.Files.Count > 0)
			{
				foreach (var file in files.Files)
				{
					KeyValuePair<string,string> record = GetRecordText(file.FileType, file.ID, LanguageID);
					pendingFiles.Add(new { RecordText = record.Value, RecordLink = record.Key, FileName = file.FileName, Age = file.CreatedOn.ConvertToAge(), Created = file.CreatedOn.ConvertToLocalDateTime().ToString(), Path = files.FilesPath });
				}
				return JsonConvert.SerializeObject(pendingFiles);
			}
			else
				return JsonConvert.SerializeObject(PhraseLibExtension.PhraseLib.Lookup("term.nomorerecords", LanguageID));
		}
		else
			return JsonConvert.SerializeObject(errorMessage);
	}


	[WebMethod]
	[System.ServiceModel.Web.WebInvoke(Method = "POST")]
	public static string GetWarnings(string URL, int LanguageID)
	{
		NodeHealthSummary summary = null;
    List<KeyValuePair<string, string>> headers = new List<KeyValuePair<string, string>>();
    IRestServiceHelper RESTServiceHelper = CurrentRequest.Resolver.Resolve<IRestServiceHelper>();
		KeyValuePair<NodeHealthSummary, string> result = RESTServiceHelper.CallService<NodeHealthSummary>(RESTServiceList.ServerHealthService, URL, LanguageID, HttpMethod.Get, String.Empty, false, headers);
		summary = result.Key;
		string errorMessage = result.Value;

		var errors = new DataTable();

		if (errorMessage == "")
		{
			if (summary != null && summary.RowCount > 0)
				errors = GetErrors(summary, LanguageID);

			if (errors.Rows.Count > 0)
			{
				var listErrors = new List<object>();
				foreach (DataRow row in errors.Rows)
					listErrors.Add(new { Description = "<a title='" + row["tooltip"] + "' href='" + row["URL"] + "'>" + row["NodeIP"] + "</a> <span title='" + row["tooltip"] + "'>" + row["Description"] + "</span>", Duration = row["Duration"] });

				return JsonConvert.SerializeObject(listErrors);
			}
			else
				return JsonConvert.SerializeObject(PhraseLibExtension.PhraseLib.Lookup("term.nomorerecords", LanguageID));
		}
		else
			return JsonConvert.SerializeObject(errorMessage);
	}

	private static DataTable GetErrors(NodeHealthSummary summary, int LanguageID)
	{
		//Errors
		DataTable dtErrors = new DataTable();
		dtErrors.Columns.AddRange(new DataColumn[] { new DataColumn("NodeIP"), new DataColumn("Description"), new DataColumn("Duration"), new DataColumn("URL"), new DataColumn("tooltip") });

		foreach (var machine in summary.Machines)
		{
			string errorMessage = "";
			long errorDate = 0;
			int errorsHigh = 0, errorsMedium = 0, errorsLow = 0;
			string term = "";
			string nodeURL = "";
			bool isAlive = true;

			foreach (var component in machine.Components)
			{
				if (component.ComponentName.ToUpper().Contains(BrokerNameConstants.PromotionBroker) || component.ComponentName.ToUpper().Contains(BrokerNameConstants.CustomerBroker))
				{
					term = PhraseLibExtension.PhraseLib.Lookup("term.node", LanguageID);
					nodeURL = "UENodeHealth.aspx?NodeName=" + machine.NodeName;
				}
				else
				{
					term = PhraseLibExtension.PhraseLib.Lookup("term.engine", LanguageID);
					nodeURL = "UEEngineHealth.aspx?NodeName=" + machine.NodeName;
				}

				if (!component.Alive)
				{
					isAlive = false;
					if (errorDate <= 0 || errorDate > component.LastHeard)
						errorDate = component.LastHeard;
				}

				List<CMS.AMS.Models.Attribute> errors = component.Attributes.Where(e => e.Code == RequestStatusConstants.Failure).ToList();
				errorsHigh += errors.Where(e => e.Severity.ToLower() == SeverityConstants.High.ToLower()).Count();
				errorsMedium += errors.Where(e => e.Severity.ToLower() == SeverityConstants.Medium.ToLower()).Count();
				errorsLow += errors.Where(e => e.Severity.ToLower() == SeverityConstants.Low.ToLower()).Count();

				foreach (var item in errors)
				{
					if (errorDate <= 0 || errorDate > item.LastUpdate)
						errorDate = item.LastUpdate;
				}

			}

			if (!isAlive)
				errorMessage = " " + term + " " + PhraseLibExtension.PhraseLib.Lookup("term.is", LanguageID) + " " + PhraseLibExtension.PhraseLib.Lookup("term.disconnected", LanguageID);

			string tooltip = "";
			if ((errorsHigh + errorsMedium + errorsLow) > 0)
			{
				tooltip = PhraseLibExtension.PhraseLib.Lookup("term.high", LanguageID) + ": " + errorsHigh.ToString() + " " + PhraseLibExtension.PhraseLib.Lookup("term.medium", LanguageID) + ": " + errorsMedium.ToString() + " " + PhraseLibExtension.PhraseLib.Lookup("term.low", LanguageID) + ": " + errorsLow.ToString();
				errorMessage = (errorMessage != string.Empty ? errorMessage + " " + PhraseLibExtension.PhraseLib.Lookup("term.and", LanguageID).ToLower() : term) + " " + PhraseLibExtension.PhraseLib.Lookup("term.has", LanguageID) + " " + (errorsHigh + errorsMedium + errorsLow).ToString() + " " + PhraseLibExtension.PhraseLib.Lookup("term.error(s)", LanguageID);
			}

			dtErrors.Rows.Add(machine.NodeIP, errorMessage, errorDate.ConvertToDuration(LanguageID), nodeURL, tooltip);
		}
		return dtErrors;
	}


	[WebMethod]
	[WebInvoke(Method = "POST")]
	public static string ToggleReportAlert(string URL, int LanguageID, bool Enabled)
	{
		string returnMessage = "";
		try
		{
      List<KeyValuePair<string, string>> headers = new List<KeyValuePair<string, string>>();
      IRestServiceHelper RESTServiceHelper = CurrentRequest.Resolver.Resolve<IRestServiceHelper>();
			KeyValuePair<string, string> result = RESTServiceHelper.CallService<string>(RESTServiceList.ServerHealthService, URL, LanguageID, HttpMethod.Post, "{\"enabled\": " + Enabled.ToJSON() + "}", false, headers);
			returnMessage = result.Value;
		}
		catch (Exception ex)
		{
			returnMessage = ex.Message;
		}
		return returnMessage;
	}
}


