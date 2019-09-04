using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Services;
using CMS;
using CMS.AMS;
using CMS.AMS.Contract;
using CMS.AMS.Models;
using System.Web.Script.Serialization;
using System.Net;
using System.IO;
using System.Text;
using Newtonsoft.Json;
using CMS.DB;

/// <summary>
/// Summary description for AjaxProcessingFunctions
/// </summary>
[WebService(Namespace = "http://tempuri.org/")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
// To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
[System.Web.Script.Services.ScriptService]
public class AjaxProcessingFunctions : System.Web.Services.WebService
{
    #region Private Variables

    private IDBAccess m_dbAccess;
    private ICollisionDetectionService m_collisionDetectionService;
    private static bool IsShaded = true;

    #endregion Private Variables

    #region Models

    [Serializable]
    [JsonObject("OfferData")]
    public class OfferData
    {
        [JsonProperty("offerId")]
        public int offerId { get; set; }

        [JsonProperty("storeNames")]
        public string[] storeNames { get; set; }
    }

    #endregion Models

    #region Constructors

    public AjaxProcessingFunctions()
    {
        CurrentRequest.Resolver.AppName = "AjaxProcessingFunctions.asmx";
        m_collisionDetectionService = CurrentRequest.Resolver.Resolve<ICollisionDetectionService>();
        m_dbAccess = CurrentRequest.Resolver.Resolve<IDBAccess>();
    }

    #endregion Constructors

    #region Web Methods

    /// <summary>
    /// 
    /// </summary>
    /// <param name="pageindex"></param>
    /// <param name="sortKey"></param>
    /// <param name="sortOrder"></param>
    /// <param name="searchingText"></param>
    /// <param name="userid"></param>
    /// <returns></returns>
    [WebMethod]
    public AMSResult<List<CMS.AMS.Models.OCD.Offer>> GetCollisionReportOfferList(Int32 pageindex, String sortKey, string sortOrder, String searchingText, Int32 userid)
    {
        return m_collisionDetectionService.GetCollisionReports(pageindex, sortKey, sortOrder, searchingText, userid);
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="OfferID"></param>
    /// <param name="RemoveType"></param>
    /// <returns></returns>
    [WebMethod]
    public AMSResult<Boolean> RemoveCollidingProducts(Int32 OfferID, Int32 RemoveType)
    {
        return m_collisionDetectionService.RemoveCollidingProducts(0, OfferID, RemoveType);
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="OfferID"></param>
    /// <param name="pageIndex"></param>
    /// <param name="sortKey"></param>
    /// <param name="sortOrder"></param>
    /// <returns></returns>
    [WebMethod]
    public AMSResult<CMS.AMS.Models.OCD.ProductList> GetCollidingProducts(Int32 OfferID, Int32 pageIndex, String sortKey, String sortOrder)
    {
        return m_collisionDetectionService.GetCollidingProducts(OfferID, pageIndex, sortKey, sortOrder);
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="OfferID"></param>
    /// <param name="QueueStatus"></param>
    /// <returns></returns>
    [WebMethod]
    public AMSResult<Boolean> UpdateCollideOfferStatus(Int64 OfferID, Int32 QueueStatus)
    {
        return m_collisionDetectionService.UpdateCollideOfferStatus(OfferID, (CMS.AMS.Models.OCD.QueueStatus)QueueStatus);
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="enableRestrictedAccessToUEOB"></param>
    /// <param name="conditionalQuery"></param>
    /// <param name="IsStoreUser"></param>
    /// <param name="ValidLocIDs"></param>
    /// <param name="ValidSU"></param>
    /// <param name="PageNum"></param>
    /// <param name="MaxRecords"></param>
    /// <returns></returns>
    [WebMethod]
    public string GetNotificationList(bool enableRestrictedAccessToUEOB = false, string conditionalQuery = "", bool IsStoreUser = false, string ValidLocIDs = "", string ValidSU = "", int LanguageId = 1, int PageNum = 1, int MaxRecords = 8)
    {
        string queryStr = "";
        string joinStr = "";
        string whereStr = "";
        SQLParametersList sqlParams = new SQLParametersList();
        DataTable ds = new DataTable();
        string returnString = "";

        //Coverity CID - 93931, 93926, 93923
        if (IsStoreUser)
        {
            //sqlParams = new SQLParametersList();
            sqlParams.Add("@ValidLocIDs", SqlDbType.VarChar).Value = ValidLocIDs;

            //sqlParams = new SQLParametersList();
            sqlParams.Add("@ValidSU", SqlDbType.VarChar).Value = ValidSU;

            joinStr = "Inner Join OfferLocUpdate olu with (NoLock) on Table1.OfferID=olu.OfferID ";
            whereStr = " where (LocationID in (@ValidLocIDs) or (CreatedByAdminID in (@ValidSU) and Isnull(LocationID,0)=0)) ";
        }

        queryStr = "SELECT * FROM (select ROW_NUMBER() OVER ( ORDER BY Date, Event ) AS NUMBER, Table1.OfferID, Date, Event, CreatedByAdminID from ( " +
                   "select OfferID, ProdStartDate as Date, 'starts' as Event, CreatedByAdminID " +
                   "from Offers as O with (NoLock) " +
                   "where Deleted = 0 and isnull(isTemplate,0) = 0 and ProdStartDate = CONVERT(datetime, CONVERT(varchar, GETDATE(), 103), 103) " +
                   "union " +
                   "select OfferID, ProdEndDate as Date, 'ends' as Event, CreatedByAdminID " +
                   "from Offers as O with (NoLock) " +
                   "where Deleted = 0 and isnull(isTemplate, 0) = 0 and ProdEndDate = CONVERT(datetime, CONVERT(varchar, GETDATE(), 103), 103) " +
                   "union " +
                   "select IncentiveID as OfferID, StartDate as Date, 'starts' as Event, CreatedByAdminID " +
                   "from CPE_Incentives as I with (NoLock) " +
                   "where Deleted=0 and isnull(isTemplate,0)=0 and StartDate=CONVERT(datetime, CONVERT(varchar, GETDATE(), 103), 103) ";
        if (enableRestrictedAccessToUEOB && !String.IsNullOrEmpty(conditionalQuery))
        {
            //sqlParams = new SQLParametersList();
            sqlParams.Add("@conditionalQuery", SqlDbType.VarChar).Value = conditionalQuery;
            queryStr += "@conditionalQuery ";
        }

        queryStr += "union " +
                    "select IncentiveID as OfferID, EndDate as Date, 'ends' as Event, CreatedByAdminID " +
                    "from CPE_Incentives as I with (NoLock) " +
                    "where Deleted=0 and isnull(isTemplate,0)=0 and EndDate=CONVERT(datetime, CONVERT(varchar, GETDATE(), 103), 103) ";
        if (enableRestrictedAccessToUEOB && !String.IsNullOrEmpty(conditionalQuery))
        {
            //sqlParams = new SQLParametersList();
            sqlParams.Add("@conditionalQuery2", SqlDbType.VarChar).Value = conditionalQuery;
            queryStr += "@conditionalQuery2 ";
        }

        //sqlParams = new SQLParametersList();
        sqlParams.Add("@joinStr", SqlDbType.VarChar).Value = joinStr;

        //sqlParams = new SQLParametersList();
        sqlParams.Add("@whereStr", SqlDbType.VarChar).Value = whereStr;

        //sqlParams = new SQLParametersList();
        sqlParams.Add("@PageNum", SqlDbType.Int).Value = PageNum;

        //sqlParams = new SQLParametersList();
        sqlParams.Add("@MaxRecords", SqlDbType.Int).Value = MaxRecords;

        queryStr += ") as Table1 @joinStr @whereStr) AS A WHERE NUMBER BETWEEN ((@PageNum - 1) * @MaxRecords + 1) AND (@PageNum * @MaxRecords) ORDER BY Date, Event Desc";
        if (joinStr == "")
        {
            queryStr = queryStr.Replace("@joinStr", "");
        }
        if (whereStr == "")
        {
            queryStr = queryStr.Replace("@whereStr", "");
        }

        ds = m_dbAccess.ExecuteQuery(DataBases.LogixRT, CommandType.Text, queryStr, sqlParams);
        foreach (DataRow row in ds.Rows)
        {
            if (row["Event"].ToString().Equals("starts", StringComparison.InvariantCultureIgnoreCase))
            {
                returnString += ConstructOutputString(row["OfferID"].ConvertToInt32(), IsShaded, true, LanguageId);
            }
            else
            {
                returnString += ConstructOutputString(row["OfferID"].ConvertToInt32(), IsShaded, false, LanguageId);
            }
            IsShaded = !IsShaded;
        }
        return returnString;
    }

    /// <summary>
    /// Gets the data from the given URL using WebRequest(using Http verb: POST)
    /// </summary>
    /// <param name="URL">URL to get the data(Parameters can be passed using Query string)</param>
    /// <returns>Response for the HttpGet</returns>
    [WebMethod]
    public string HttpPost(string URL, int offerId, string storeNames)
    {
        try
        {
          IRestServiceHelper restservice = CurrentRequest.Resolver.Resolve<IRestServiceHelper>();
          OfferData ofd = new OfferData() { offerId = offerId, storeNames = storeNames.Split(',') };
          var PostData = Newtonsoft.Json.JsonConvert.SerializeObject(ofd);
          KeyValuePair<String, String> response = restservice.CallService<String>(RESTServiceList.PromotionBrokerService, URL, 1, System.Net.Http.HttpMethod.Post, PostData, false);
          return response.Key;
        }
        catch (Exception ex)
        {
            return ex.Message;
        }
    }

    string ConstructOutputString(int OfferId, bool IsShaded, bool isStartEvent, int LanguageId)
    {
        string returnString = "";
        string shadedString = "";
        if (IsShaded)
        {
            shadedString = "class=\"shaded\"";
        }
        else
        {
            shadedString = "";
        }

        returnString = "<p " + shadedString + " style=\"padding: 2px; margin-bottom: 1px; margin-left: 1px; margin-top: 0px; border: 0px;\">";

        if (isStartEvent)
        {
            returnString += "<a href=\"offer-redirect.aspx?OfferID=" + OfferId + "\">" +
              Copient.PhraseLib.Lookup("term.offer", LanguageId) + " " + OfferId + "</a> " +
            Copient.PhraseLib.Lookup("term.starts", LanguageId).ToLower();
        }
        else
        {
            returnString += "<a href=\"offer-redirect.aspx?OfferID=" + OfferId + "\">" +
              Copient.PhraseLib.Lookup("term.offer", LanguageId) + " " + OfferId + "</a> " +
            Copient.PhraseLib.Lookup("term.ends", LanguageId).ToLower();
        }

        returnString += "</p>";

    return returnString;
  }

    [WebMethod]
    public string GetLockedSystemOptions(string offerID)
    {
        SQLParametersList sqlParams = new SQLParametersList();
        string sqlQuery = "SELECT (Select EndDate from CPE_Incentives where IncentiveID = @offerID) as EndDate,(Select OptionValue from UE_SystemOptions where OptionID = 80) as OptionValue";
        sqlParams.Add("@offerID", SqlDbType.Int).Value = Convert.ToInt32(offerID);
        DataTable DT = m_dbAccess.ExecuteQuery(DataBases.LogixRT, CommandType.Text, sqlQuery, sqlParams);
        if (DT != null && DT.Rows.Count > 0)
        {
            if(Convert.ToDateTime(DT.Rows[0]["EndDate"]) <= DateTime.Today && Convert.ToInt32(DT.Rows[0]["OptionValue"]) == 1)
            {
                return "true";
            }
        }
        return "false";
    }

    #endregion WebMethods
}
