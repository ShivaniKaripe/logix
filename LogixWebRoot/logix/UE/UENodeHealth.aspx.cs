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

public partial class UENodeHealth : AuthenticatedUI
{

	#region Private Variables

	private IHealth m_health;
	private IActivityLogService m_ActivityLogService;
	private NodeHealth nodeHealth;

	#endregion Private Variables

	protected void Page_Load(object sender, EventArgs e)
	{
		m_health = CurrentRequest.Resolver.Resolve<IHealth>();
		m_ActivityLogService = CurrentRequest.Resolver.Resolve<IActivityLogService>();
		ucServerHealthTabs.LanguageID = LanguageID;
		ucServerHealthTabs.Title = PhraseLib.Lookup("term.health", LanguageID) + " " + PhraseLib.Lookup("term.for", LanguageID).ToLower();
		((UE_MasterPage)this.Master).Tab_Name = "8_10";
		hdnURL.Value = m_health.HealthServiceURL;

		if (!IsPostBack)
		{
			ucServerHealthTabs.EnginesInstalled = m_health.GetInstalledEngines(LanguageID);
			ucServerHealthTabs.SetInfoMessage("", false);
			PopulateNodesPage();
		}

	}

	protected override void OnInit(EventArgs e)
	{
		AppName = "UEnodeHealth.aspx";
		base.OnInit(e);
	}

	private void PopulateNodesPage()
	{
		//Bind Health
    List<KeyValuePair<string, string>> headers = new List<KeyValuePair<string, string>>();
    IRestServiceHelper RESTServiceHelper = CurrentRequest.Resolver.Resolve<IRestServiceHelper>();
		KeyValuePair<NodeHealth, string> result = RESTServiceHelper.CallService<NodeHealth>(RESTServiceList.ServerHealthService, (m_health.HealthServiceURL!= string.Empty? m_health.HealthServiceURL + "/nodes/" + Request.QueryString["NodeName"]:""), LanguageID, HttpMethod.Get, String.Empty, false, headers);
		string errorMessage = result.Value;
		nodeHealth = result.Key;

		//Errors
		DataTable dtErrors = new DataTable();
		dtErrors.Columns.AddRange(new DataColumn[] { new DataColumn("Severity"), new DataColumn("ParamID"), new DataColumn("Description"), new DataColumn("Duration")});
		gvNodeWarnings.Columns[0].HeaderText = PhraseLib.Lookup("term.severity", LanguageID);
		gvNodeWarnings.Columns[1].HeaderText = PhraseLib.Lookup("term.code", LanguageID);
		gvNodeWarnings.Columns[2].HeaderText = PhraseLib.Lookup("term.description", LanguageID);
		gvNodeWarnings.Columns[3].HeaderText = PhraseLib.Lookup("term.duration", LanguageID);

		if (errorMessage == string.Empty && nodeHealth != null)
		{
			Alert.Checked = nodeHealth.Machine.Alert;
			Report.Checked = nodeHealth.Machine.Report;

			bool hasPromotionBroker = false;
			bool hasCustomerBroker = false;

			foreach (var component in nodeHealth.Machine.Components)
			{
				if (!component.Alive)
					component.Attributes.Insert(0, new CMS.AMS.Models.Attribute { Severity = component.Severity, Code = RequestStatusConstants.Failure, ParamID = (component.ComponentName.ToUpper().Contains(BrokerNameConstants.PromotionBroker) ? (component.IsPromoFetchNode ? ServerHealthErrorCodes.PromoFecthNode_Disconnected: ServerHealthErrorCodes.PromotionBroker_Disconnected) : ServerHealthErrorCodes.CustomerBroker_Disconnected), Description = (PhraseLibExtension.PhraseLib.Lookup("term.disconnected", LanguageID)), LastUpdate = component.LastHeard });

				foreach (var error in component.Attributes.Where(e => e.Code == RequestStatusConstants.Failure))
					dtErrors.Rows.Add(PhraseLib.Lookup("term." + error.Severity, LanguageID), error.ParamID, ServerHealthHelper.GetErrorDescription(error.ParamID, LanguageID), error.LastUpdate.ConvertToDuration(LanguageID));

				if (component.ComponentName.ToUpper().Contains(BrokerNameConstants.PromotionBroker))
				{
					lblPBStatus.Text = component.Alive ? PhraseLib.Lookup("term.connected", LanguageID) : "<font color=\"red\">" + PhraseLib.Lookup("term.disconnected", LanguageID) + "</font>";
					lblPBLastHeard.Text = component.LastHeard.ConvertToLocalDateTime().ToString();

                    foreach(var attr in component.Attributes)
                    {
                        if (attr.Description.ToUpper().Contains("FETCH"))
                            lblLastUpdateOffer.Text = attr.LastUpdate.ConvertToLocalDateTime().ToString();
                        if (attr.Description.ToUpper().Contains("IPL"))
                            lblLastIPL.Text = attr.LastUpdate.ConvertToLocalDateTime().ToString();
                    }
                    if (lblLastUpdateOffer.Text == "") lblLastUpdateOffer.Text = "-";
                    if (lblLastIPL.Text == "") lblLastIPL.Text = "-";
                    hasPromotionBroker = true;
				}
				else if (component.ComponentName.ToUpper().Contains(BrokerNameConstants.CustomerBroker))
				{
					lblCBStatus.Text = component.Alive ? PhraseLib.Lookup("term.connected", LanguageID) : "<font color=\"red\">" + PhraseLib.Lookup("term.disconnected", LanguageID) + "</font>";
					lblCBLastHeard.Text = component.LastHeard.ConvertToLocalDateTime().ToString();
					lblLastLookUp.Text = ServerHealthHelper.RetrieveAttribute(component.Attributes,ServerHealthErrorCodes.CustomerBroker_LastCustomerLookup);
					lblLastTransDownload.Text = ServerHealthHelper.RetrieveAttribute(component.Attributes,ServerHealthErrorCodes.CustomerBroker_LastTransactionDownload);
					lblLastTransUpload.Text = ServerHealthHelper.RetrieveAttribute(component.Attributes, ServerHealthErrorCodes.CustomerBroker_LastTransactionUpload);
					hasCustomerBroker = true;
				}

			}
			divCB.Visible = hasCustomerBroker;
			divPB.Visible = hasPromotionBroker;
            communication.Visible = hasPromotionBroker;

            if (dtErrors.Rows.Count > 0)
			{
				gvNodeWarnings.DataSource = dtErrors;
				gvNodeWarnings.DataBind();
			}
			else
				lblError.Text = "<center>" + PhraseLib.Lookup("term.norecords", LanguageID) + "</center>";

			//Identification
			lblNodeName.Text = Server.HtmlEncode(nodeHealth.Machine.NodeName);
			lblIpAddress.Text = Server.HtmlEncode(nodeHealth.Machine.NodeIP);

		}
		else
			ucServerHealthTabs.SetInfoMessage(errorMessage, true, true);
	}

}


