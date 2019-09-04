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

public partial class UENodeHealthSummary : AuthenticatedUI
{
	#region Private Variables

	private IHealth m_health;
	private IActivityLogService m_ActivityLogService;
	public NodeHealthSummary nodeSummary;

	#endregion Private Variables

	public int PAGEINDEX { get { return hdnPageIndex.Value.ConvertToInt16(); } set { hdnPageIndex.Value = value.ToString(); } }
	public string SORTBY { get { return hdnSortText.Value; } set { hdnSortText.Value = value; } }
	public string SORTDIR { get { return hdnSortDir.Value; } set { hdnSortDir.Value = value; } }
	public string SEARCH { get { return hdnSearch.Value; } set { hdnSearch.Value = value; } }
	public string FILTER { get { return hdnFilter.Value; } set { hdnFilter.Value = value; } }
	public int PAGESIZE { get { return hdnPageSize.Value.ConvertToInt16(); } set { hdnPageSize.Value = value.ToString(); } }


	protected override void OnInit(EventArgs e)
	{
		AppName = "UENodeHealthSummary.aspx";
		base.OnInit(e);
	}

	private void InitializePage()
	{
		filterNodeHealth.Items.Add(new ListItem(PhraseLib.Lookup("term.allerrors", LanguageID), "1"));
		filterNodeHealth.Items.Add(new ListItem(PhraseLib.Lookup("term.showall", LanguageID), "2"));
		filterNodeHealth.Items.Add(new ListItem(PhraseLib.Lookup("term.disconnected", LanguageID) + " " + PhraseLib.Lookup("term.nodes", LanguageID), "5"));
		filterNodeHealth.Items.Add(new ListItem(PhraseLib.Lookup("term.communications", LanguageID) + " " + PhraseLib.Lookup("term.ok", LanguageID), "6"));

		filterNodeHealth.SelectedIndex = 0;
	}

	protected void Page_Load(object sender, EventArgs e)
	{
		m_health = CurrentRequest.Resolver.Resolve<IHealth>();
		m_ActivityLogService = CurrentRequest.Resolver.Resolve<IActivityLogService>();
		ucServerHealthTabs.LanguageID = LanguageID;
		ucServerHealthTabs.Title = PhraseLib.Lookup("term.health", LanguageID) + " " + PhraseLib.Lookup("term.for", LanguageID).ToLower();
		((UE_MasterPage)this.Master).Tab_Name = "8_10";
		gvNodes.PageSize = PAGESIZE = 30;
		btnSearch.Text = PhraseLib.Lookup("term.search", LanguageID);
        string tempsearch = PhraseLib.Lookup("term.nodes", LanguageID);
        txtSearch.Attributes["placeholder"] = tempsearch;
		if (!IsPostBack)
		{
			ucServerHealthTabs.EnginesInstalled = m_health.GetInstalledEngines(LanguageID);
			hdnURL.Value = m_health.HealthServiceURL;
			ucServerHealthTabs.SetInfoMessage("", false);
			InitializePage();

			PAGEINDEX = 1;
			if (string.IsNullOrEmpty(Request.QueryString["SORTBY"]))
				SORTBY = "node";
			else
				SORTBY = Request.QueryString["SORTBY"];

			if (string.IsNullOrEmpty(Request.QueryString["SORTDIR"]))
				SORTDIR = "ASC";
			else
				SORTDIR = Request.QueryString["SORTDIR"];

			SetSortOrderImage(div_node, SORTDIR);

			switch (SORTBY)
			{
				case "report":
					SetSortOrderImage(div_report, SORTDIR);
					break;
				case "alert":
					SetSortOrderImage(div_alert, SORTDIR);
					break;
				default:
					SetSortOrderImage(div_node, SORTDIR);
					break;
			}

			if (!string.IsNullOrEmpty(Request.QueryString["SEARCH"]))
				txtSearch.Text = SEARCH = Request.QueryString["SEARCH"];

			if (string.IsNullOrEmpty(Request.QueryString["FILTER"]))
				FILTER = "0";
			else
			{
				FILTER = Request.QueryString["FILTER"];
				filterNodeHealth.SelectedIndex = FILTER.ConvertToInt16();
			}

			PopulateNodesPage();

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
				filter = "DisconnectedNodes";
				break;

			case 3:
				filter = "CommunicationsOK";
				break;
		}

		return filter;
	}
	private void PopulateNodesPage()
	{
		hdnPageCount.Value = "0";
    List<KeyValuePair<string, string>> headers = new List<KeyValuePair<string, string>>();
    IRestServiceHelper RESTServiceHelper = CurrentRequest.Resolver.Resolve<IRestServiceHelper>();
		KeyValuePair<NodeHealthSummary, string> result = RESTServiceHelper.CallService<NodeHealthSummary>(RESTServiceList.ServerHealthService,(hdnURL.Value != string.Empty?  hdnURL.Value + "/nodes?offset=0&pagesize=" + PAGESIZE.ToString() + "&sort=" + SORTBY + "&sortdir=" + SORTDIR + "&search=" + Server.UrlEncode(SEARCH) + "&filter=" + GetFilter(FILTER.ConvertToInt16()) + (FILTER.ConvertToInt16() == 0 || FILTER.ConvertToInt16() == 2 ? "&report=true" : "&report=all"):""), LanguageID, HttpMethod.Get, String.Empty, false, headers);
		string error = result.Value;
		nodeSummary = result.Key;

		if (error == string.Empty)
		{
			if (nodeSummary != null && nodeSummary.RowCount > 0 && nodeSummary.Machines != null)
			{
				//Nodes
				DataTable dtNodes = FormatNodes(nodeSummary);
				gvNodes.DataSource = dtNodes;
				gvNodes.DataBind();

				if (nodeSummary.RowCount <= PAGESIZE)
					hdnPageCount.Value = "1";
				else
				{
					decimal value = (nodeSummary.RowCount.ConvertToDecimal() / PAGESIZE.ConvertToDecimal());
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
			loadmoreajaxloader.InnerHtml = "<center>" + PhraseLibExtension.PhraseLib.Lookup("term.norecords", LanguageID) + "</center>";
		}
		else if (hdnPageCount.Value == "1")
		{
			loadmoreajaxloader.InnerHtml = "<center>" + PhraseLibExtension.PhraseLib.Lookup("term.nomorerecords", LanguageID) + "</center>";
		}
	}

	protected void NodeSearchChanged_Event(object sender, EventArgs e)
	{
		Response.Redirect("UENodeHealthSummary.aspx?SORTBY=" + SORTBY + "&SORTDIR=" + SORTDIR + "&PAGEINDEX=1&SEARCH=" + Server.UrlEncode(txtSearch.Text) + "&FILTER=" + filterNodeHealth.SelectedIndex);
	}

	protected void NodeFilterChanged_Event(object sender, EventArgs e)
	{
		Response.Redirect("UENodeHealthSummary.aspx?SORTBY=" + SORTBY + "&SORTDIR=" + SORTDIR + "&PAGEINDEX=1&SEARCH=" + Server.UrlEncode(txtSearch.Text) + "&FILTER=" + filterNodeHealth.SelectedIndex);
	}

	protected void SortChanged_Event(object sender, EventArgs e)
	{
		LinkButton lnk = (LinkButton)sender;
		string arg = lnk.CommandArgument.ToString();
		if (SORTBY == arg)
			GetSortDirection();
		else
		{
			SORTBY = arg;
			SORTDIR = "ASC";
		}

		Response.Redirect("UENodeHealthSummary.aspx?SORTBY=" + SORTBY + "&SORTDIR=" + SORTDIR + "&PAGEINDEX=1&SEARCH=" + Server.UrlEncode(txtSearch.Text) + "&FILTER=" + filterNodeHealth.SelectedIndex);

	}


	protected void gvNodes_OnRowDataBound(object sender, GridViewRowEventArgs e)
	{
		if (e.Row.RowType == DataControlRowType.DataRow)
		{
			int RowNum = gvNodes.DataKeys[e.Row.RowIndex].Value.ConvertToInt16();
			e.Row.Cells[5].Style.Add("display", "none");

			if (e.Row.RowIndex % 2 == 0)
				e.Row.Attributes.Add("class", "shaded");

			e.Row.Attributes.Add("id", "row" + RowNum.ToString());

			HtmlTableRow pblinkdiv = ((HtmlTableRow)e.Row.FindControl("pblinkdiv"));
			HtmlTableRow cblinkdiv = ((HtmlTableRow)e.Row.FindControl("cblinkdiv"));

			pblinkdiv.ID = pblinkdiv.ID + RowNum.ToString();
			cblinkdiv.ID = cblinkdiv.ID + RowNum.ToString();

			IEnumerable<CMS.AMS.Models.Attribute> pbErrors = null;
			IEnumerable<CMS.AMS.Models.Attribute> cbErrors = null;

			var machine = nodeSummary.Machines.Where(a => a.RowNum == RowNum).First();
			if (machine.Components.Where(c => c.ComponentName.ToUpper().Contains(BrokerNameConstants.PromotionBroker)).Count() > 0)
				pbErrors = machine.Components.Where(c => c.ComponentName.ToUpper().Contains(BrokerNameConstants.PromotionBroker)).First().Attributes.Where(b => b.Code == RequestStatusConstants.Failure);
			else
				pblinkdiv.Attributes.Add("show", "hide");

			if (machine.Components.Where(c => c.ComponentName.ToUpper().Contains(BrokerNameConstants.CustomerBroker)).Count() > 0)
				cbErrors = machine.Components.Where(c => c.ComponentName.ToUpper().Contains(BrokerNameConstants.CustomerBroker)).First().Attributes.Where(b => b.Code == RequestStatusConstants.Failure);
			else
				cblinkdiv.Attributes.Add("show", "hide");

			AddReportandAlert(e, RowNum, pbErrors, cbErrors, machine);

			//Promotion Broker
			BindErrors(pbErrors, RowNum, ((GridView)e.Row.FindControl("gvPBErrors")), ((HtmlImage)e.Row.FindControl("pbimgdiv")), ((HtmlAnchor)e.Row.FindControl("pblink")), "pberrordiv" + RowNum.ToString());
			//CUstomer Broker
			BindErrors(cbErrors, RowNum, ((GridView)e.Row.FindControl("gvCBErrors")), ((HtmlImage)e.Row.FindControl("cbimgdiv")), ((HtmlAnchor)e.Row.FindControl("cblink")), "cberrordiv" + RowNum.ToString());

		}

	}

	private void AddReportandAlert(GridViewRowEventArgs e, int RowNum, IEnumerable<CMS.AMS.Models.Attribute> pbErrors, IEnumerable<CMS.AMS.Models.Attribute> cbErrors, Machine machine)
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
            string script = "javascript:ToggleReport(this,'" + machine.NodeName + "','" + hdnURL.Value + "'," + RowNum.ToString() + ");";

            imageReport.Attributes.Add("onclick", Server.HtmlEncode(script));
			imageReport.Attributes.Add("title", PhraseLib.Lookup("term.clickhere", LanguageID) + " " + PhraseLib.Lookup("term.to", LanguageID) + " " + PhraseLib.Lookup("term.enable", LanguageID) + "/" + PhraseLib.Lookup("term.disable", LanguageID) + " " + PhraseLib.Lookup("term.error", LanguageID).ToLower() + " " + PhraseLib.Lookup("term.reporting", LanguageID).ToLower() + " " + PhraseLib.Lookup("term.for", LanguageID).ToLower() + " " + PhraseLib.Lookup("term.this", LanguageID).ToLower() + " " + PhraseLib.Lookup("term.node", LanguageID).ToLower());
		}

		if (imageAlert != null)
		{
			if (machine.Alert)
				imageAlert.Attributes.Add("src", "../../images/email.png");
			else
				imageAlert.Attributes.Add("src", "../../images/email-off.png");
            string script = "javascript:ToggleAlert(this,'" + machine.NodeName + "','" + hdnURL.Value + "');";
			imageAlert.Attributes.Add("onclick", Server.HtmlEncode(script));
			imageAlert.Attributes.Add("title", PhraseLib.Lookup("term.clickhere", LanguageID) + " " + PhraseLib.Lookup("term.to", LanguageID) + " " + PhraseLib.Lookup("term.enable", LanguageID) + "/" + PhraseLib.Lookup("term.disable", LanguageID) + " " + PhraseLib.Lookup("term.alertemail", LanguageID).ToLower() + " " + PhraseLib.Lookup("term.for", LanguageID).ToLower() + " " + PhraseLib.Lookup("term.this", LanguageID).ToLower() + " " + PhraseLib.Lookup("term.node", LanguageID).ToLower());
		}

		if (machine.Report)
		{
			if ((pbErrors != null && pbErrors.Where(a => a.Severity == SeverityConstants.High).Count() > 0) || (cbErrors != null && cbErrors.Where(a => a.Severity == SeverityConstants.High).Count() > 0))
				e.Row.Attributes.Add("class", "shadeddarkred");
			else if ((pbErrors != null && pbErrors.Where(a => a.Severity == SeverityConstants.Medium).Count() > 0) || (cbErrors != null && cbErrors.Where(a => a.Severity == SeverityConstants.Medium).Count() > 0))
				e.Row.Attributes.Add("class", "shadedred");
			else if ((pbErrors != null && pbErrors.Where(a => a.Severity == SeverityConstants.Low).Count() > 0) || (cbErrors != null && cbErrors.Where(a => a.Severity == SeverityConstants.Low).Count() > 0))
				e.Row.Attributes.Add("class", "shadedlightred");
		}
	}

	private void BindErrors(IEnumerable<CMS.AMS.Models.Attribute> Errors, int RowNum, GridView gvErrors, HtmlImage image, HtmlAnchor anchor, string errorDivID)
	{

			//localize grid columns
			gvErrors.Columns[0].HeaderText = PhraseLib.Lookup("term.severity", LanguageID);
			gvErrors.Columns[1].HeaderText = PhraseLib.Lookup("term.code", LanguageID);
			gvErrors.Columns[2].HeaderText = PhraseLib.Lookup("term.description", LanguageID);
			gvErrors.Columns[3].HeaderText = PhraseLib.Lookup("term.duration", LanguageID);

			if (image != null)
			{
				image.ID = image.ID + RowNum.ToString();
                String toottip = PhraseLib.Lookup("store-health-ue.ViewHideErrorDetails", LanguageID);
				image.Attributes.Add("title", System.Web.HttpUtility.HtmlDecode(toottip));
			}

		DataTable dtErrors = new DataTable();
		dtErrors.Columns.AddRange(new DataColumn[] { new DataColumn("RowNum"), new DataColumn("ParamID"), new DataColumn("Severity"), new DataColumn("Description"), new DataColumn("Duration"), new DataColumn("NodeIP") });

		if (Errors != null && Errors.Count() > 0)
		{
			foreach (var error in Errors)
				dtErrors.Rows.Add(RowNum.ToString(), error.ParamID, PhraseLib.Lookup("term." + error.Severity, LanguageID), ServerHealthHelper.GetErrorDescription(error.ParamID, LanguageID), error.LastUpdate.ConvertToDuration(LanguageID), RowNum);
			
			gvErrors.DataSource = dtErrors;
			gvErrors.DataBind();

			if (image != null && anchor != null)
			{
				image.Attributes.Add("src", "../../images/plus2.png");
				anchor.Attributes.Add("href", "JavaScript:divexpandcollapse('" + errorDivID + "','" + image.ID + "','','')");
			}

		}
		else
		{
			if (image != null)
				image.Attributes.Add("src", "../../images/plus2-disabled.png");

			if (anchor != null)
				anchor.Attributes.Remove("href");

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
			SORTDIR = "DESC";
		else
			SORTDIR = "ASC";
	}

	[WebMethod]
	[WebInvoke(Method = "POST")]
	public static string GetNodes(string URL, int LanguageID)
	{
    List<KeyValuePair<string, string>> headers = new List<KeyValuePair<string, string>>();
    IRestServiceHelper RESTServiceHelper = CurrentRequest.Resolver.Resolve<IRestServiceHelper>();
		KeyValuePair<NodeHealthSummary, string> result = RESTServiceHelper.CallService<NodeHealthSummary>(RESTServiceList.ServerHealthService, URL, LanguageID, HttpMethod.Get, String.Empty, false, headers);
		NodeHealthSummary nodeSummary = result.Key;
		string errorMessage = result.Value;

		if (errorMessage == string.Empty)
		{
			var nodes = new List<object>();
			if (nodeSummary.Machines.Count > 0)
			{
				foreach (var machine in nodeSummary.Machines)
				{

					List<CMS.AMS.Models.Attribute> allComponentsErrors = new List<CMS.AMS.Models.Attribute>();

					string PBerrorHtml = "", CBerrorHtml = "";
					bool HasPB = false, HasCB = false;

					foreach (var component in machine.Components)
					{
						if (!component.Alive)
							component.Attributes.Insert(0, new CMS.AMS.Models.Attribute { Severity = component.Severity, ParamID = (component.ComponentName.ToUpper().Contains(BrokerNameConstants.PromotionBroker) ? (component.IsPromoFetchNode ? ServerHealthErrorCodes.PromoFecthNode_Disconnected : ServerHealthErrorCodes.PromotionBroker_Disconnected) : ServerHealthErrorCodes.CustomerBroker_Disconnected), Code = RequestStatusConstants.Failure, Description = ((component.ComponentName.ToUpper().Contains(BrokerNameConstants.PromotionBroker) ? PhraseLibExtension.PhraseLib.Lookup("term.promotionbroker", LanguageID) : PhraseLibExtension.PhraseLib.Lookup("term.customerbroker", LanguageID)) + " " + PhraseLibExtension.PhraseLib.Lookup("term.disconnected", LanguageID)), LastUpdate = component.LastHeard });

						var errors = component.Attributes.Where(e => e.Code == RequestStatusConstants.Failure);
						allComponentsErrors.AddRange(errors);

						if (component.ComponentName.ToUpper().Contains(BrokerNameConstants.PromotionBroker))
						{
							PBerrorHtml = ServerHealthHelper.GenerateErrorTable(errors, LanguageID);
							HasPB = true;
						}
						else if (component.ComponentName.ToUpper().Contains(BrokerNameConstants.CustomerBroker))
						{
							CBerrorHtml = ServerHealthHelper.GenerateErrorTable(errors, LanguageID);
							HasCB = true;
						}						
					}

					string errorString = ServerHealthHelper.FormatErrors(allComponentsErrors, LanguageID);
					nodes.Add(new { nodeIp = machine.NodeIP, storeName = machine.StoreName, nodeName = machine.NodeName, storeId = machine.StoreID, errorString = errorString, PBerrorHtml = PBerrorHtml, CBerrorHtml = CBerrorHtml, hasCB = HasCB, hasPB = HasPB, Report=machine.Report,Alert=machine.Alert });
				}
				return JsonConvert.SerializeObject(nodes);
			}
			else
				return JsonConvert.SerializeObject(PhraseLibExtension.PhraseLib.Lookup("term.nomorerecords", LanguageID));
		}
		else
			return JsonConvert.SerializeObject(errorMessage);
	}


	private DataTable FormatNodes(NodeHealthSummary nodeSummary)
	{
		DataTable dtNodes = new DataTable();
		dtNodes.Columns.AddRange(new DataColumn[] { new DataColumn("RowNum"), new DataColumn("NodeIP"), new DataColumn("NodeName"), new DataColumn("StoreID"), new DataColumn("StoreName"), new DataColumn("Errors")});

		int RowNum = 0;

		foreach (var machine in nodeSummary.Machines)
		{
			string errorString = "";
			List<CMS.AMS.Models.Attribute> allComponentsErrors = new List<CMS.AMS.Models.Attribute>();

			foreach (var component in machine.Components)
			{
				if (!component.Alive)
					component.Attributes.Insert(0, new CMS.AMS.Models.Attribute { Severity = component.Severity, ParamID = (component.ComponentName.ToUpper().Contains(BrokerNameConstants.PromotionBroker) ? (component.IsPromoFetchNode ? ServerHealthErrorCodes.PromoFecthNode_Disconnected : ServerHealthErrorCodes.PromotionBroker_Disconnected) : ServerHealthErrorCodes.CustomerBroker_Disconnected), Code = RequestStatusConstants.Failure, Description = ((component.ComponentName.ToUpper().Contains(BrokerNameConstants.PromotionBroker) ? PhraseLibExtension.PhraseLib.Lookup("term.promotionbroker", LanguageID) : PhraseLibExtension.PhraseLib.Lookup("term.customerbroker", LanguageID)) + " " + PhraseLibExtension.PhraseLib.Lookup("term.disconnected", LanguageID)), LastUpdate = component.LastHeard });

				allComponentsErrors.AddRange(component.Attributes.Where(e => e.Code == RequestStatusConstants.Failure));
			}

			errorString = ServerHealthHelper.FormatErrors(allComponentsErrors, LanguageID);
			RowNum++;
			machine.RowNum = RowNum;
			dtNodes.Rows.Add(RowNum, machine.NodeIP, machine.NodeName, machine.StoreID, machine.StoreName, errorString);

		}
		return dtNodes;
	}
}