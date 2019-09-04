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
using CMS;
using System.Web.Services;
using System.Net;
using System.IO;
using System.Text;
using Newtonsoft.Json;
using System.ServiceModel.Web;
using System.Web.UI.HtmlControls;
using System.Net.Http;

public partial class UEEngineHealth : AuthenticatedUI
{

	#region Private Variables

	private IHealth m_health;
	private IActivityLogService m_ActivityLogService;
	private NodeHealth engineHealth;

	#endregion Private Variables

	protected void Page_Load(object sender, EventArgs e)
	{
		m_health = CurrentRequest.Resolver.Resolve<IHealth>();
		m_ActivityLogService = CurrentRequest.Resolver.Resolve<IActivityLogService>();
		ucServerHealthTabs.LanguageID = LanguageID;
		ucServerHealthTabs.Title = PhraseLib.Lookup("term.health", LanguageID) + " " + PhraseLib.Lookup("term.for", LanguageID).ToLower();
		hdnURL.Value = m_health.HealthServiceURL;
		hdnNodeName.Value = Request.QueryString["NodeName"];
		((UE_MasterPage)this.Master).Tab_Name = "8_10";

		if (!IsPostBack)
		{
			ucServerHealthTabs.EnginesInstalled = m_health.GetInstalledEngines(LanguageID);
			ucServerHealthTabs.SetInfoMessage("", false);
			PopulateEnginesPage();

		}

	}

	protected override void OnInit(EventArgs e)
	{
		AppName = "UEEngineHealth.aspx";
		base.OnInit(e);
	}

	private void PopulateEnginesPage()
	{
		//Bind Health
		string errorMessage = "";
    List<KeyValuePair<string, string>> headers = new List<KeyValuePair<string, string>>();
    IRestServiceHelper RESTServiceHelper = CurrentRequest.Resolver.Resolve<IRestServiceHelper>();
		KeyValuePair<NodeHealth,string> result = RESTServiceHelper.CallService<NodeHealth>(RESTServiceList.ServerHealthService,(hdnURL.Value != string.Empty? hdnURL.Value + "/engines/":"") + Request.QueryString["NodeName"],LanguageID,HttpMethod.Get, String.Empty, false, headers);
		engineHealth=result.Key;
		errorMessage = result.Value;
		
		if (errorMessage==string.Empty && engineHealth != null)
		{
			//Identification
			lblEngineName.Text = Server.HtmlEncode(engineHealth.Machine.NodeName);
			lblIpAddress.Text = Server.HtmlEncode(engineHealth.Machine.NodeIP);
			lblStatus.Text = engineHealth.Machine.Components[0].Alive ? PhraseLib.Lookup("term.connected", LanguageID) : "<font color=\"red\">" + PhraseLib.Lookup("term.disconnected", LanguageID) + "</font>";
			lblType.Text = (engineHealth.Machine.Components[0].EnterpriseEngine) ? PhraseLib.Lookup("serverhealth.enterpriseengine", LanguageID) : PhraseLib.Lookup("serverhealth.storeengine", LanguageID);
			Alert.Checked = engineHealth.Machine.Alert;
			Report.Checked = engineHealth.Machine.Report;
			hdnLocationID.Value = engineHealth.Machine.StoreID;
			hdnEnterpriseEngine.Value = engineHealth.Machine.Components[0].EnterpriseEngine.ToString().ToLower();

			if (engineHealth.Machine.Components[0].EnterpriseEngine)
			{
				linkConfig.Visible = false;
			}
			else
			{
				linkConfig.Text = PhraseLib.Lookup("term.storeconfiguration", LanguageID);
				linkConfig.NavigateUrl = "../store-edit.aspx?LocationID=" + hdnLocationID.Value;
				linkConfig.Visible = true;				
			}

			lblStore.Text = Server.HtmlEncode((engineHealth.Machine.Components[0].EnterpriseEngine) ? PhraseLib.Lookup("term.all", LanguageID) + " " + PhraseLib.Lookup("term.stores", LanguageID) : engineHealth.Machine.StoreName);

			//Last CardHolder Look-Up
			lblCardholderlookup.Text = ServerHealthHelper.RetrieveAttribute(engineHealth.Machine.Components[0].Attributes, (engineHealth.Machine.Components[0].EnterpriseEngine == true ? ServerHealthErrorCodes.EnterpriseEngine_LastCustomerLookup : ServerHealthErrorCodes.PromotionEngine_LastCustomerLookup));

			//Last TransUpload
			lblTransactionupload.Text = ServerHealthHelper.RetrieveAttribute(engineHealth.Machine.Components[0].Attributes, (engineHealth.Machine.Components[0].EnterpriseEngine == true ? ServerHealthErrorCodes.EnterpriseEngine_LastTransactionUpload : ServerHealthErrorCodes.PromotionEngine_LastTransactionUpload));

			//Last Sync
			lblsync.Text = ServerHealthHelper.RetrieveAttribute(engineHealth.Machine.Components[0].Attributes, (engineHealth.Machine.Components[0].EnterpriseEngine == true ? ServerHealthErrorCodes.EnterpriseEngine_LastSync: ServerHealthErrorCodes.PromotionEngine_LastSync));

			//Last Heard
			lblLastHeard.Text = engineHealth.Machine.Components[0].LastHeard.ConvertToLocalDateTime().ToString();

			//Broker to Engine Pending Files
			//BindFiles(hdnURL.Value + "/engines/" + engineHealth.Machine.NodeName + "/files?offset=0&pagesize=" + hdnPageSize.Value, gvEngineFiles, hdnPageSize, hdnPageCount, loadmoreajaxloader);
			BindFiles(engineHealth.PendingFilesURL + "?offset=0&pagesize=" + hdnPageSize.Value, gvEngineFiles, hdnPageSize, hdnPageCount, loadmoreajaxloader);
			hdnPendingFilesURL.Value = engineHealth.PendingFilesURL;

			//Logix to Broker Pending Files
			if (engineHealth.Machine.Components[0].EnterpriseEngine)
			{
				BindFiles(hdnURL.Value + "/enterprise/logixfiles?offset=0&pagesize=" + hdnPageSize1.Value, gvBrokerFiles, hdnPageSize1, hdnPageCount1, loadmoreajaxloader1);
			}
			else
			{
				BindFiles(hdnURL.Value + "/stores/" + hdnLocationID.Value + "/logixfiles?offset=0&pagesize=" + hdnPageSize1.Value, gvBrokerFiles, hdnPageSize1, hdnPageCount1, loadmoreajaxloader1);
			}

			//Errors
			if (!engineHealth.Machine.Components[0].Alive)
				engineHealth.Machine.Components[0].Attributes.Insert(0, new CMS.AMS.Models.Attribute { Severity = engineHealth.Machine.Components[0].Severity, Code = RequestStatusConstants.Failure, ParamID = (engineHealth.Machine.Components[0].EnterpriseEngine ? ServerHealthErrorCodes.EnterpriseEngine_Disconnected : ServerHealthErrorCodes.PromotionEngine_Disconnected), Description = (PhraseLibExtension.PhraseLib.Lookup("term.engine", LanguageID) + " " + PhraseLibExtension.PhraseLib.Lookup("term.disconnected", LanguageID)), LastUpdate = engineHealth.Machine.Components[0].LastHeard });

			DataTable dtErrors = new DataTable();
			dtErrors.Columns.AddRange(new DataColumn[] { new DataColumn("Severity"), new DataColumn("ParamID"), new DataColumn("Description"), new DataColumn("Duration")});

			foreach (var error in engineHealth.Machine.Components[0].Attributes.Where(e => e.Code == RequestStatusConstants.Failure))
			{
				dtErrors.Rows.Add(PhraseLib.Lookup("term." + error.Severity, LanguageID),error.ParamID, ServerHealthHelper.GetErrorDescription(error.ParamID,LanguageID), error.LastUpdate.ConvertToDuration(LanguageID) );
			}

			gvEngineWarnings.Columns[0].HeaderText = PhraseLib.Lookup("term.severity", LanguageID);
			gvEngineWarnings.Columns[1].HeaderText = PhraseLib.Lookup("term.code", LanguageID);
			gvEngineWarnings.Columns[2].HeaderText = PhraseLib.Lookup("term.description", LanguageID);
			gvEngineWarnings.Columns[3].HeaderText = PhraseLib.Lookup("term.duration", LanguageID);

			if (dtErrors.Rows.Count == 0)
				lblError.Text = "<center>" + PhraseLib.Lookup("term.norecords", LanguageID) + "</center>";
			else
			{
				gvEngineWarnings.DataSource = dtErrors;
				gvEngineWarnings.DataBind();
			}

		}
		else
		{
			ucServerHealthTabs.SetInfoMessage(errorMessage, true, true);
			loadmoreajaxloader.Visible = false;
			loadmoreajaxloader1.Visible = false;
		}

	}

	private void BindFiles(string url, GridView gvFiles, HtmlInputHidden pageSize, HtmlInputHidden pageCount, HtmlGenericControl loader)
	{
		try
		{
      List<KeyValuePair<string, string>> headers = new List<KeyValuePair<string, string>>();
      IRestServiceHelper RESTServiceHelper = CurrentRequest.Resolver.Resolve<IRestServiceHelper>();
			KeyValuePair<PendingFiles, string> result = RESTServiceHelper.CallService<PendingFiles>(RESTServiceList.ServerHealthService,url, LanguageID, HttpMethod.Get, string.Empty, false, headers);
			PendingFiles files = result.Key;
			string error = result.Value;
			pageCount.Value = "0";

			if (error != string.Empty)
			{
				loader.InnerHtml = "<center>" + error + "</center>";
			}
			else if (files.Files.Count > 0)
			{
				//Engines

				DataTable dtFiles = new DataTable();
				dtFiles.Columns.AddRange(new DataColumn[] { new DataColumn("OfferID"), new DataColumn("FileName"), new DataColumn("Age"), new DataColumn("Created"), new DataColumn("Path") });

				foreach (PendingFile file in files.Files)
				{
					dtFiles.Rows.Add(file.ID.ToString(), file.FileName, file.CreatedOn.ConvertToAge(), file.CreatedOn.ConvertToLocalDateTime().ToString(), files.FilesPath);
				}

				if (dtFiles.Rows.Count > 0)
				{
					gvFiles.DataSource = dtFiles;
					gvFiles.DataBind();
				}
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

	[WebMethod]
	[System.ServiceModel.Web.WebInvoke(Method = "POST")]
	public static string GetEngineFiles(string URL,int LanguageID)
	{
		string errorMessage = "";

		PendingFiles files = null;
    List<KeyValuePair<string, string>> headers = new List<KeyValuePair<string, string>>();
    IRestServiceHelper RESTServiceHelper = CurrentRequest.Resolver.Resolve<IRestServiceHelper>();
		KeyValuePair<PendingFiles, string> result = RESTServiceHelper.CallService<PendingFiles>(RESTServiceList.ServerHealthService, URL, LanguageID, HttpMethod.Get, String.Empty, false, headers);
		files = result.Key;
		errorMessage = result.Value;

		var pendingFiles = new List<object>();

		if (errorMessage == "")
		{
			if (files != null && files.Files.Count > 0)
			{
				foreach (var file in files.Files)
				{
					pendingFiles.Add(new {OfferID=file.ID, FileName = file.FileName, Age = file.CreatedOn.ConvertToAge(), Created = file.CreatedOn.ConvertToLocalDateTime().ToString(),Path=files.FilesPath });
				}
				return JsonConvert.SerializeObject(pendingFiles);
			}
			else
			{
				return JsonConvert.SerializeObject(PhraseLibExtension.PhraseLib.Lookup("term.nomorerecords",LanguageID));
			}
			
		}
		else
		{
			return JsonConvert.SerializeObject(errorMessage);
		}
		
	}
}


