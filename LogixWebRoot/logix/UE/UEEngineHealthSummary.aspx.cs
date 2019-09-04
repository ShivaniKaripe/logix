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

public partial class UEEngineHealthSummary : AuthenticatedUI
{

	#region Private Variables

	private IHealth m_health;
	private IActivityLogService m_ActivityLogService;
	public NodeHealthSummary engineSummary;

	#endregion Private Variables

	public int PAGEINDEX { get { return hdnPageIndex.Value.ConvertToInt16(); } set { hdnPageIndex.Value = value.ToString(); } }
	public string SORTBY { get { return hdnSortText.Value; } set { hdnSortText.Value = value; } }
	public string SORTDIR { get { return hdnSortDir.Value; } set { hdnSortDir.Value = value; } }
	public string SEARCH { get { return hdnSearch.Value; } set { hdnSearch.Value = value; } }
	public string FILTER { get { return hdnFilter.Value; } set { hdnFilter.Value = value; } }
	public int PAGESIZE { get { return hdnPageSize.Value.ConvertToInt16(); } set { hdnPageSize.Value = value.ToString(); } }


	protected override void OnInit(EventArgs e)
	{
		AppName = "UEEngineHealthSummary.aspx";
		base.OnInit(e);
	}

	private void InitializePage()
	{
		filterEngineHealth.Items.Add(new ListItem(PhraseLib.Lookup("term.allerrors", LanguageID),"1"));
		filterEngineHealth.Items.Add(new ListItem(PhraseLib.Lookup("term.showall", LanguageID),"2"));
		filterEngineHealth.Items.Add(new ListItem(PhraseLib.Lookup("term.all", LanguageID) +" " + PhraseLib.Lookup("term.enterprise", LanguageID) +" " + PhraseLib.Lookup("term.engines", LanguageID),"3"));
		filterEngineHealth.Items.Add(new ListItem( PhraseLib.Lookup("term.all", LanguageID) + " " + PhraseLib.Lookup("term.store", LanguageID) + " " + PhraseLib.Lookup("term.engines", LanguageID),"4"));
		filterEngineHealth.Items.Add(new ListItem( PhraseLib.Lookup("term.disconnected", LanguageID) +" " + PhraseLib.Lookup("term.engines", LanguageID),"5"));
		filterEngineHealth.Items.Add(new ListItem( PhraseLib.Lookup("term.communications", LanguageID) + " " + PhraseLib.Lookup("term.ok", LanguageID),"6"));

		filterEngineHealth.SelectedIndex = 0;
	}

    protected void Page_Load(object sender, EventArgs e)
    {
		m_health = CurrentRequest.Resolver.Resolve<IHealth>();
		m_ActivityLogService = CurrentRequest.Resolver.Resolve<IActivityLogService>();
		ucServerHealthTabs.LanguageID = LanguageID;
		ucServerHealthTabs.Title = PhraseLib.Lookup("term.health", LanguageID) + " " + PhraseLib.Lookup("term.for", LanguageID).ToLower();
		((UE_MasterPage)this.Master).Tab_Name = "8_10";
		gvEngines.PageSize = PAGESIZE = 30;
		btnSearch.Text = PhraseLib.Lookup("term.search", LanguageID);
        string tempsearch = PhraseLib.Lookup("term.stores",LanguageID) + "/" + PhraseLib.Lookup("term.engines",LanguageID);
        txtSearch.Attributes["placeholder"] = tempsearch;
		if (!IsPostBack)
		{
			ucServerHealthTabs.EnginesInstalled = m_health.GetInstalledEngines(LanguageID);
			hdnURL.Value = m_health.HealthServiceURL;
			ucServerHealthTabs.SetInfoMessage("", false);
			InitializePage();
			
			PAGEINDEX = 1;

			if (string.IsNullOrEmpty(Request.QueryString["SORTBY"]))
				SORTBY = "engine";
			else
				SORTBY = Request.QueryString["SORTBY"];

			if (string.IsNullOrEmpty(Request.QueryString["SORTDIR"]))
				SORTDIR = "ASC";
			else
				SORTDIR = Request.QueryString["SORTDIR"];

			switch (SORTBY)
			{
				case "location":
					SetSortOrderImage(div_storeName , SORTDIR);
					break;
				case "report":
					SetSortOrderImage(div_report , SORTDIR);
					break;
				case "alert":
					SetSortOrderImage(div_alert , SORTDIR);
					break;
				default:
					SetSortOrderImage(div_nodeIp , SORTDIR);
					break;
			}
			

			if (!string.IsNullOrEmpty(Request.QueryString["SEARCH"]))
				txtSearch.Text = SEARCH = Request.QueryString["SEARCH"];

			if (string.IsNullOrEmpty(Request.QueryString["FILTER"]))
				FILTER = "0";
			else
			{
				FILTER = Request.QueryString["FILTER"];
				filterEngineHealth.SelectedIndex = FILTER.ConvertToInt16();
			}
			
			PopulateEnginesPage();

		}
    }

	private void PopulateEnginesPage()
	{
		hdnPageCount.Value = "0";
    List<KeyValuePair<string, string>> headers = new List<KeyValuePair<string, string>>();
    IRestServiceHelper RESTServiceHelper = CurrentRequest.Resolver.Resolve<IRestServiceHelper>();
		KeyValuePair<NodeHealthSummary, string> result = RESTServiceHelper.CallService<NodeHealthSummary>(RESTServiceList.ServerHealthService,(hdnURL.Value != string.Empty?  hdnURL.Value + "/engines?offset=0&pagesize=" + PAGESIZE.ToString() + "&sort=" + SORTBY + "&sortdir=" + SORTDIR + "&search=" + Server.UrlEncode(SEARCH) + "&filter=" + GetFilter(FILTER.ConvertToInt16()) + (FILTER.ConvertToInt16() == 0 || FILTER.ConvertToInt16() == 4 ? "&report=true" : "&report=all") :""), LanguageID, HttpMethod.Get, String.Empty, false, headers);
		engineSummary = result.Key;
		string error = result.Value;

		if (error == string.Empty)
		{
			if (engineSummary == null)
			{
				ucServerHealthTabs.SetInfoMessage(error, true);
			}
			else if (engineSummary.Machines.Count > 0)
			{
				//Engines
				gvEngines.DataSource = FormatEngines(engineSummary);
				gvEngines.DataBind();

				if (engineSummary.RowCount <= PAGESIZE)
				{
					hdnPageCount.Value = "1";
				}
				else
				{
					decimal value = (engineSummary.RowCount.ConvertToDecimal() / PAGESIZE.ConvertToDecimal());
					hdnPageCount.Value = Math.Ceiling(value).ToString();
				}
			}
		}
		else
		{
			ucServerHealthTabs.SetInfoMessage(error, true);
		}

		if (hdnPageCount.Value == "0")
		{
			loadmoreajaxloader.InnerHtml = "<center>" + PhraseLib.Lookup("term.norecords", LanguageID) + "</center>";
		}
		else if (hdnPageCount.Value == "1")
		{
			loadmoreajaxloader.InnerHtml = "<center>" + PhraseLib.Lookup("term.nomorerecords", LanguageID) + "</center>";
		}
	}

	private string GetFilter(int Index)
	{
		string filter = "";
		switch (Index)
		{
			case 0:
				filter = "AllErrors";
				break;

			case 1:
				filter = "ShowAll";
				break;

			case 2:
				filter = "AllEnterpriseEngines";
				break;

			case 3:
				filter = "AllStoreEngines";
				break;

			case 4:
				filter = "DisconnectedEngines";
				break;

			case 5:
				filter = "CommunicationsOK";
				break;
		}

		return filter;
	}

	
	protected void EngineSearchChanged_Event(object sender, EventArgs e)
	{
		Response.Redirect("UEEngineHealthSummary.aspx?SORTBY=" + SORTBY + "&SORTDIR=" + SORTDIR + "&PAGEINDEX=1&SEARCH=" + Server.UrlEncode(txtSearch.Text) + "&FILTER=" + filterEngineHealth.SelectedIndex);
	}

	protected void EngineFilterChanged_Event(object sender, EventArgs e)
	{
		Response.Redirect("UEEngineHealthSummary.aspx?SORTBY=" + SORTBY + "&SORTDIR=" + SORTDIR + "&PAGEINDEX=1&SEARCH=" + Server.UrlEncode(txtSearch.Text) + "&FILTER=" + filterEngineHealth.SelectedIndex); 
	}

	protected void SortChanged_Event(object sender, EventArgs e)
	{
		LinkButton lnk = (LinkButton)sender;
		string arg = lnk.CommandArgument.ToString();
		if (SORTBY == arg)
		{
			GetSortDirection();
		}
		else
		{
			SORTBY = arg;
			SORTDIR = "ASC";
		}

		Response.Redirect("UEEngineHealthSummary.aspx?SORTBY=" + SORTBY + "&SORTDIR=" + SORTDIR + "&PAGEINDEX=1&SEARCH=" + Server.UrlEncode(txtSearch.Text) + "&FILTER=" + filterEngineHealth.SelectedIndex);

	}


	protected void gvEngines_OnRowDataBound(object sender, GridViewRowEventArgs e)
	{
		if (e.Row.RowType == DataControlRowType.DataRow)
		{
			int RowNum = gvEngines.DataKeys[e.Row.RowIndex].Value.ConvertToInt16();

			GridView gvErrors = (GridView)e.Row.FindControl("gvErrors");

			e.Row.Cells[6].Style.Add("display", "none");

			if (e.Row.RowIndex % 2 == 0)
				e.Row.Attributes.Add("class", "shaded");

			e.Row.Attributes.Add("id", "row"+ RowNum.ToString());

			DataTable dtErrors = new DataTable();
			dtErrors.Columns.AddRange(new DataColumn[] { new DataColumn("RowNum"),new DataColumn("ParamID"), new DataColumn("Severity"), new DataColumn("Description"), new DataColumn("Duration"), new DataColumn("NodeIP") });

			var machine = engineSummary.Machines.Where(a => a.RowNum == RowNum).First();
			var Errors = machine.Components[0].Attributes.Where(b => b.Code == RequestStatusConstants.Failure);

			HtmlImage image = (HtmlImage)e.Row.FindControl("plusimage");
			HtmlAnchor anchor = (HtmlAnchor)e.Row.FindControl("link");

			if (image != null)
			{
				image.ID = "imgdiv" + RowNum.ToString();
                String toottip = PhraseLib.Lookup("store-health-ue.ViewHideErrorDetails", LanguageID);
				image.Attributes.Add("title", System.Web.HttpUtility.HtmlDecode(toottip));
			}
				
				//localize grid columns
				gvErrors.Columns[0].HeaderText = PhraseLib.Lookup("term.severity", LanguageID);
				gvErrors.Columns[1].HeaderText = PhraseLib.Lookup("term.code", LanguageID);
				gvErrors.Columns[2].HeaderText = PhraseLib.Lookup("term.description", LanguageID);
				gvErrors.Columns[3].HeaderText = PhraseLib.Lookup("term.duration", LanguageID);

			if (Errors.Count() > 0)
			{
				if(machine.Report){
					if (Errors.Where(a => a.Severity == SeverityConstants.High).Count() > 0)
						e.Row.Attributes.Add("class", "shadeddarkred");
					else if(Errors.Where(a => a.Severity == SeverityConstants.Medium).Count() > 0)
						e.Row.Attributes.Add("class", "shadedred");
					else if(Errors.Where(a => a.Severity == SeverityConstants.Low).Count() > 0)
						e.Row.Attributes.Add("class", "shadedlightred");
				}

				foreach (var error in Errors)
				{
					dtErrors.Rows.Add(RowNum.ToString(), error.ParamID, PhraseLibExtension.PhraseLib.Lookup("term." + error.Severity, LanguageID), ServerHealthHelper.GetErrorDescription(error.ParamID, LanguageID), error.LastUpdate.ConvertToDuration(LanguageID).Replace("Days", PhraseLib.Lookup("term.days", LanguageID)).Replace("Minutes", PhraseLib.Lookup("term.minutes", LanguageID)).Replace("Hours", PhraseLib.Lookup("term.hours", LanguageID)), RowNum);
				}
				
				gvErrors.DataSource = dtErrors;
				gvErrors.DataBind();
				if (image != null)
					image.Attributes.Add("src", "../../images/plus2.png");

				if (anchor != null)
					anchor.Attributes.Add("href", "JavaScript:divexpandcollapse('div" + RowNum.ToString()+"')");

			}
			else
			{
				if (image != null)
					image.Attributes.Add("src", "../../images/plus2-disabled.png");

				if (anchor != null)
					anchor.Attributes.Remove("href");

			}

			AddReportandAlert(e, RowNum, machine);
		}
	}

	private void AddReportandAlert(GridViewRowEventArgs e, int RowNum, Machine machine)
	{
		//Report and Alert
		HtmlImage imageReport = (HtmlImage)e.Row.FindControl("Report");
		HtmlImage imageAlert = (HtmlImage)e.Row.FindControl("Alert");

		if (imageReport != null)
		{
			if (machine.Report)
				imageReport.Attributes.Add("src", "../../images/report-on.png");
			else
				imageReport.Attributes.Add("src", "../../images/report-off.png");
            string script = "javascript:ToggleReport(this,'" + machine.NodeName + "','" + hdnURL.Value + "','row" + RowNum.ToString() + "','div" + RowNum.ToString() + "');";
			imageReport.Attributes.Add("onclick", Server.HtmlEncode(script));
			imageReport.Attributes.Add("title", PhraseLib.Lookup("term.clickhere", LanguageID) + " " + PhraseLib.Lookup("term.to", LanguageID) + " " + PhraseLib.Lookup("term.enable", LanguageID) + "/" + PhraseLib.Lookup("term.disable", LanguageID) + " " + PhraseLib.Lookup("term.error", LanguageID).ToLower() + " " + PhraseLib.Lookup("term.reporting", LanguageID).ToLower() + " " + PhraseLib.Lookup("term.for", LanguageID).ToLower() + " " + PhraseLib.Lookup("term.this", LanguageID).ToLower() + " " + PhraseLib.Lookup("term.engine", LanguageID).ToLower());
		}

		if (imageAlert != null)
		{
			if (machine.Alert)
				imageAlert.Attributes.Add("src", "../../images/email.png");
			else
				imageAlert.Attributes.Add("src", "../../images/email-off.png");
            string script = "javascript:ToggleAlert(this,'" + machine.NodeName + "','" + hdnURL.Value + "');";
			imageAlert.Attributes.Add("onclick", Server.HtmlEncode(script));
			imageAlert.Attributes.Add("title", PhraseLib.Lookup("term.clickhere", LanguageID) + " " + PhraseLib.Lookup("term.to", LanguageID) + " " + PhraseLib.Lookup("term.enable", LanguageID) + "/" + PhraseLib.Lookup("term.disable", LanguageID) + " " + PhraseLib.Lookup("term.alertemail", LanguageID).ToLower() + " " + PhraseLib.Lookup("term.for", LanguageID).ToLower() + " " + PhraseLib.Lookup("term.this", LanguageID).ToLower() + " " + PhraseLib.Lookup("term.engine", LanguageID).ToLower());
		}
	}

	
	private void SetSortOrderImage(HtmlGenericControl div, String sortorder)
	{		
		if (div != null)
		{
			if (sortorder == "ASC")
			{				
				
				div.Visible = true;
				div.InnerHtml = "<span class=\"sortarrow\">&#9660;</span>";
			}
			else if (sortorder == "DESC")
			{
				div.Visible = true;
				div.InnerHtml = "<span class=\"sortarrow\">&#9650;</span>";
			}
		}
	}

	private void GetSortDirection()
	{
		if (SORTDIR == "ASC")
		{
			SORTDIR = "DESC";
		}
		else
		{
			SORTDIR = "ASC";
		}
	}



	[WebMethod]
	[WebInvoke(Method = "POST")]
	public static string GetEngines(string URL, int LanguageID)
	{
    List<KeyValuePair<string, string>> headers = new List<KeyValuePair<string, string>>();
    IRestServiceHelper RESTServiceHelper = CurrentRequest.Resolver.Resolve<IRestServiceHelper>();
		KeyValuePair<NodeHealthSummary, string> result = RESTServiceHelper.CallService<NodeHealthSummary>(RESTServiceList.ServerHealthService, URL, LanguageID, HttpMethod.Get, String.Empty, false, headers);
		string errorMessage = result.Value;
		NodeHealthSummary engineSummary = result.Key;
	
		if (errorMessage == string.Empty)
		{
			var engines = new List<object>();
			if (engineSummary.Machines.Count > 0)
			{
				foreach (var machine in engineSummary.Machines)
				{
					if (!machine.Components[0].Alive)
						machine.Components[0].Attributes.Insert(0, new CMS.AMS.Models.Attribute { Severity = machine.Components[0].Severity, Code = RequestStatusConstants.Failure, ParamID = (machine.Components[0].EnterpriseEngine ? ServerHealthErrorCodes.EnterpriseEngine_Disconnected : ServerHealthErrorCodes.PromotionEngine_Disconnected), Description = (PhraseLibExtension.PhraseLib.Lookup("term.engine", LanguageID) + " " + PhraseLibExtension.PhraseLib.Lookup("term.disconnected", LanguageID)), LastUpdate = machine.Components[0].LastHeard });

					string errorHtml = "";
					var errors = machine.Components[0].Attributes.Where(e => e.Code == RequestStatusConstants.Failure);

					string errorString = ServerHealthHelper.FormatErrors(errors, LanguageID);

					errorHtml = ServerHealthHelper.GenerateErrorTable(errors, LanguageID);

					engines.Add(new { nodeIp = machine.NodeIP, storeName = (machine.Components[0].EnterpriseEngine) ? PhraseLibExtension.PhraseLib.Lookup("term.all", LanguageID) + " " + PhraseLibExtension.PhraseLib.Lookup("term.stores", LanguageID) : machine.StoreName, nodeName = machine.NodeName, storeId = machine.StoreID, errorString = errorString, errorHtml = errorHtml, Report=machine.Report, Alert=machine.Alert });
				}
				return JsonConvert.SerializeObject(engines);
			}
			else		{
				return JsonConvert.SerializeObject(PhraseLibExtension.PhraseLib.Lookup("term.nomorerecords", LanguageID));
			}
		}
		else
		{
			return JsonConvert.SerializeObject(errorMessage);
		}
	}


	private DataTable FormatEngines(NodeHealthSummary engineSummary)
	{
		DataTable dtEngines = new DataTable();
		dtEngines.Columns.AddRange(new DataColumn[] { new DataColumn("RowNum"), new DataColumn("NodeIP"), new DataColumn("NodeName"), new DataColumn("StoreID"), new DataColumn("StoreName"), new DataColumn("Errors") });
		
		int RowNum = 0;

		foreach (var machine in engineSummary.Machines)
		{
			if (!machine.Components[0].Alive)
			{
				machine.Components[0].Attributes.Insert(0, new CMS.AMS.Models.Attribute { Severity = machine.Components[0].Severity, Code = RequestStatusConstants.Failure, ParamID = (machine.Components[0].EnterpriseEngine ? ServerHealthErrorCodes.EnterpriseEngine_Disconnected : ServerHealthErrorCodes.PromotionEngine_Disconnected), Description = PhraseLib.Lookup("term.engine", LanguageID) + " " + PhraseLib.Lookup("term.disconnected", LanguageID), LastUpdate = machine.Components[0].LastHeard });
			}

			string errorString = ServerHealthHelper.FormatErrors(machine.Components[0].Attributes.Where(e => e.Code == RequestStatusConstants.Failure),LanguageID);
			
			RowNum++;
			machine.RowNum = RowNum;
			dtEngines.Rows.Add(RowNum, machine.NodeIP, machine.NodeName, machine.StoreID, (machine.Components[0].EnterpriseEngine) ? PhraseLib.Lookup("term.all", LanguageID) + " " + PhraseLib.Lookup("term.stores", LanguageID) : machine.StoreName, errorString);

		}
		return dtEngines;
	}
	
}