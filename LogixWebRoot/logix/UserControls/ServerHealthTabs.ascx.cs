using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using CMS.AMS;
using CMS.AMS.Models;
using System.Data;

public partial class ServerHealthTabs : System.Web.UI.UserControl
{
	public int LanguageID { get; set; }
	public string Title { get; set; }
	public DataTable EnginesInstalled
	{
		set
		{
			if (value.Rows.Count > 1)
			{
				foreach (DataRow item in value.Rows)
					ddlEngines.Items.Add(new ListItem(item["EngineName"].ToString(), item["EngineID"].ToString()));
				ddlEngines.SelectedValue = ((int)Engines.UE).ToString();
			}
			else if (value.Rows.Count == 1)
			{
				ddlEngines.Visible = false;
				infobar.Style.Remove("height");
				Title += " " + value.Rows[0]["EngineName"].ToString();
			}
		}
	}

	protected void Page_Load(object sender, EventArgs e)
	{
		lblTitle.Text = Title;
		SetTabCssClass(Request.Url.ToString());
	}

	protected void ServerHealthSummary_Click(object sender, EventArgs e)
	{
		Response.Redirect("/logix/UE/UEServerHealthSummary.aspx");
	}

	protected void EngineHealth_Click(object sender, EventArgs e)
	{
		Response.Redirect("/logix/UE/UEEngineHealthSummary.aspx");
	}

	protected void NodeHealth_Click(object sender, EventArgs e)
	{
		Response.Redirect("/logix/UE/UENodeHealthSummary.aspx");
	}

	public void SetInfoMessage(string Text, bool IsError, bool isAppend = false)
	{
		if (Text == string.Empty)
		{
			infobar.Text = "&nbsp;";
			infobar.CssClass = "warnings infobar";
		}
		else
		{
			if (isAppend)
				infobar.Text = infobar.Text + " " + Server.HtmlEncode(Text);
			else
				infobar.Text = Server.HtmlEncode(Text);

			if (IsError)
				infobar.CssClass = "warnings red-background infobar";
			else
				infobar.CssClass = "warnings green-background infobar";
		}
	}

	private void SetTabCssClass(string URL)
	{
		if (URL.Contains("NodeHealth"))
		{
			NodeHealth.CssClass += " on";
			ServerHealthSummary.CssClass = ServerHealthSummary.CssClass.Replace("on", "");
			EngineHealth.CssClass = EngineHealth.CssClass.Replace("on", "");
		}
		else if (URL.Contains("EngineHealth"))
		{
			EngineHealth.CssClass += " on";
			ServerHealthSummary.CssClass = ServerHealthSummary.CssClass.Replace("on", "");
			NodeHealth.CssClass = NodeHealth.CssClass.Replace("on", "");
		}
		else
		{
			ServerHealthSummary.CssClass += " on";
			EngineHealth.CssClass = EngineHealth.CssClass.Replace("on", "");
			NodeHealth.CssClass = NodeHealth.CssClass.Replace("on", "");
		}

	}

	protected void ddlEngines_SelectedIndexChanged(object sender, EventArgs e)
	{
		if (ddlEngines.SelectedValue == ((int)Engines.CPE).ToString())
			Response.Redirect("/logix/store-health-cpe.aspx?filterhealth=2");
		else if (ddlEngines.SelectedValue == ((int)Engines.CM).ToString())
			Response.Redirect("/logix/store-health-cm.aspx?filterhealth=2");
	}
}