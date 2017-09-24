using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Security.Claims;
using System.Threading.Tasks;
using Office365PlannerTask.Utils;
using System.Net.Http;
using System.Net.Http.Headers;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System.Text;
using System.Collections.Generic;
using System;
using System.Net;
using System.Configuration;
using Microsoft.Graph;

namespace Office365PlannerTask.Models
{
    public class MyPlansRepository
    {
        public static string GraphResourceUrl = "https://graph.microsoft.com/V1.0";
        public static string TenantId = ConfigurationManager.AppSettings["ida:TenantId"];
        public static bool UseSDK = Convert.ToBoolean(ConfigurationManager.AppSettings["UseSDK"]);

        public static async Task<GraphServiceClient> GetGraphServiceAsync()
        {
            var accessToken = await GraphAuthHelper.GetGraphAccessTokenAsync();
            var graphserviceClient = new GraphServiceClient(GraphResourceUrl,
                                          new DelegateAuthenticationProvider(
                                                        (requestMessage) =>
                                                        {
                                                            requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", accessToken);
                                                            return Task.FromResult(0);
                                                        }));

            return graphserviceClient;
        }

        public async Task<List<MyPlan>> GetPlans()
        {
            if (UseSDK)
            {
                return await GetPlansSDK();
            }
            else
            {
                return await GetPlansREST();
            }

        }

        public async Task<List<MyPlan>> GetPlansREST()
        {
            var plansResult = new List<MyPlan>();
            var accessToken = await GraphAuthHelper.GetGraphAccessTokenAsync();
            var restURL = string.Format("{0}me/planner/plans/", SettingsHelper.GraphResourceUrl);
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    var accept = "application/json";

                    client.DefaultRequestHeaders.Add("Accept", accept);
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                    using (var response = await client.GetAsync(restURL))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            var jsonresult = JObject.Parse(await response.Content.ReadAsStringAsync());

                            foreach (var item in jsonresult["value"])
                            {
                                plansResult.Add(new MyPlan
                                {
                                    id = item["id"].ToString(),
                                    title = item["title"].ToString(),
                                    owner = !string.IsNullOrEmpty(item["owner"].ToString()) ? item["owner"].ToString() : "",
                                    createdBy = !string.IsNullOrEmpty(item["createdBy"].ToString()) ? item["createdBy"].ToString() : "",
                                    Etag = !string.IsNullOrEmpty(item["@odata.etag"].ToString()) ? item["@odata.etag"].ToString() : ""
                                });
                            }
                        }
                    }
                }
            }
            catch (Exception el)
            {
                el.ToString();
            }

            return plansResult;
        }

        public async Task<List<MyPlan>> GetPlansSDK()
        {
            var plansResult = new List<MyPlan>();

            try
            {
                var graphServiceClient = await GetGraphServiceAsync();
                var plans = await graphServiceClient.Me.Planner.Plans.Request().GetAsync();


                foreach (var item in plans)
                {
                    plansResult.Add(new MyPlan
                    {
                        id = item.Id,
                        title = item.Title,
                        owner = item.Owner,
                        createdBy = item.CreatedBy.User.DisplayName,
                        Etag = item.GetEtag().ToString()
                    });
                }



            }
            catch (Exception el)
            {
                el.ToString();
            }

            return plansResult;
        }

        public async Task<MyPlan> GetPlan(string id)
        {
            if (UseSDK)
            {
                return await GetPlanSDK(id);
            }
            else
            {
                return await GetPlanREST(id);
            }
        }


        public async Task<MyPlan> GetPlanREST(string id)
        {
            MyPlan plan = new MyPlan();
            var accessToken = await GraphAuthHelper.GetGraphAccessTokenAsync();
            var restURL = string.Format("{0}planner/plans/{1}", SettingsHelper.GraphResourceUrl, id);
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    var accept = "application/json";

                    client.DefaultRequestHeaders.Add("Accept", accept);
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                    using (var response = await client.GetAsync(restURL))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            var item = JObject.Parse(await response.Content.ReadAsStringAsync());

                            if (item != null)
                            {
                                plan.title = !string.IsNullOrEmpty(item["title"].ToString()) ? item["title"].ToString() : string.Empty;
                                plan.Etag = !string.IsNullOrEmpty(item["@odata.etag"].ToString()) ? item["@odata.etag"].ToString() : "";
                            }
                        }
                    }
                }
            }
            catch (Exception el)
            {
                el.ToString();
            }

            return plan;
        }

        public async Task<MyPlan> GetPlanSDK(string id)
        {
            MyPlan plan = new MyPlan();
            var graphServiceClient = await GetGraphServiceAsync();
            var reqPlan = await graphServiceClient.Planner.Plans[id].Request().GetAsync();

            try
            {

                if (reqPlan != null)
                {
                    plan.title = reqPlan.Title;
                    plan.Etag = reqPlan.GetEtag();
                }

            }
            catch (Exception el)
            {
                el.ToString();
            }

            return plan;
        }

        public async Task CreatePlan(MyPlan myPlan)
        {
            if (UseSDK)
            {
                 await CreatePlanSDK(myPlan);
            }
            else
            {
                 await CreatePlanREST(myPlan);
            }
        }

        public async Task CreatePlanSDK(MyPlan myPlan)
        {
            try
            {
                string groupId = await CreateGroupSDK(myPlan.title);
                await CreatePlanREST(myPlan, groupId);
            }
            catch (Exception el)
            {
                el.ToString();
            }
        }

        public async Task CreatePlanREST(MyPlan myPlan)
        {
            try
            {
                string groupId = await CreateGroupREST(myPlan.title);
                await CreatePlanREST(myPlan, groupId);
            }
            catch (Exception el)
            {
                el.ToString();
            }
        }

        private async Task<string> CreateGroupSDK(string groupTitle)
        {
            var graphServiceClient = await GetGraphServiceAsync();
            string groupId = string.Empty;
            var newGroup = new Group
            {
                DisplayName = groupTitle,
                MailNickname = groupTitle.Replace(" ", ""),
                SecurityEnabled = false,
                MailEnabled = true,
                GroupTypes = new List<string>() { "Unified" }
            };

            try
            {
                var createdGroup = await graphServiceClient.Groups.Request().AddAsync(newGroup);

                var me = await graphServiceClient.Me.Request().GetAsync();

                await graphServiceClient.Groups[createdGroup.Id].Members.References.Request().AddAsync(me);

                groupId = createdGroup.Id;

            }
            catch (Exception)
            {

                throw;
            }

            
            return groupId;
        }

        private async Task<string> CreateGroupREST(string groupTitle)
        {
            var accessToken = await GraphAuthHelper.GetGraphAccessTokenAsync();
            string groupId = string.Empty;
            dynamic groupJSON = new JObject();
            groupJSON.displayName = groupTitle;
            groupJSON.mailNickname = groupTitle.Replace(" ", "");
            groupJSON.securityEnabled = false;
            groupJSON.mailEnabled = true;
            groupJSON.groupTypes = new JArray("Unified");
            HttpRequestMessage message = new HttpRequestMessage(HttpMethod.Post, string.Format("{0}groups", SettingsHelper.GraphResourceUrl));
            message.Content = new StringContent(groupJSON.ToString(), System.Text.Encoding.UTF8, "application/json");
            message.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            message.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            using (HttpClient client = new HttpClient())
            {
                var responseMessage = await client.SendAsync(message);

                if (responseMessage.StatusCode != System.Net.HttpStatusCode.Created)
                    throw new Exception(responseMessage.StatusCode.ToString());

                var payload = await responseMessage.Content.ReadAsStringAsync();

                groupId = JObject.Parse(payload)["id"].ToString();

                await AddMemberForGroup(groupId);
            }

            return groupId;
        }

        private async Task AddMemberForGroup(string groupid)
        {
            var accessToken = await GraphAuthHelper.GetGraphAccessTokenAsync();
            var userObjectId = ClaimsPrincipal.Current.FindFirst(SettingsHelper.ClaimTypeObjectIdentifier).Value;
            var restURL = string.Format("{0}groups('{1}')/members/$ref", SettingsHelper.GraphResourceUrl, groupid);
            string strAddMememberToGroup = "{\"@odata.id\":\"" + SettingsHelper.GraphResourceUrl + "users('" + userObjectId + "')\"}";

            try
            {
                using (HttpClient client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                    var requestMessage = new HttpRequestMessage(HttpMethod.Post, restURL);
                    requestMessage.Content = new StringContent(strAddMememberToGroup, System.Text.Encoding.UTF8, "application/json");
                    using (var response = await client.SendAsync(requestMessage))
                    {
                        if (response.IsSuccessStatusCode)
                            return;
                        else
                            throw new Exception("add memeber to group error: " + response.StatusCode);
                    }
                }
            }
            catch (Exception el)
            {
                el.ToString();
            }
        }

        private async Task CreatePlanREST(MyPlan myPlan, string groupId)
        {
            var accessToken = await GraphAuthHelper.GetGraphAccessTokenAsync();
            var restURL = string.Format("{0}planner/plans/", SettingsHelper.GraphResourceUrl);
            dynamic postPlanJSON = new JObject();
            postPlanJSON.title = myPlan.title;
            postPlanJSON.owner = groupId;

            try
            {
                using (HttpClient client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                    var requestMessage = new HttpRequestMessage(HttpMethod.Post, restURL);
                    requestMessage.Content = new StringContent(postPlanJSON.ToString(), System.Text.Encoding.UTF8, "application/json");
                    using (var response = await client.SendAsync(requestMessage))
                    {
                        if (response.IsSuccessStatusCode)
                            return;
                        else
                            throw new Exception("add plan error: " + response.StatusCode);
                    }
                }
            }
            catch (Exception el)
            {
                el.ToString();
            }
        }

        public async Task UpdatePlan(MyPlan myPlan)
        {
            if (UseSDK)
            {
                await UpdatePlanSDK(myPlan);
            }
            else
            {
                await UpdatePlanREST(myPlan);
            }
        }

        public async Task UpdatePlanSDK(MyPlan myPlan)
        {
            
            try
            {
                var graphServiceClient = await GetGraphServiceAsync();
                var reqPlan = await graphServiceClient.Planner.Plans[myPlan.id].Request().GetAsync();

                string etag = reqPlan.GetEtag();
                var plan = new PlannerPlan
                {
                    Title = myPlan.title
                };
                //reqPlan.Owner = myPlan.owner;

                await graphServiceClient.Planner.Plans[reqPlan.Id].Request().Header("If-Match", etag).Header("Prefer", "return=representation").UpdateAsync(plan);

            }
            catch (Exception el)
            {
                el.ToString();
            }
        }

        public async Task UpdatePlanREST(MyPlan myPlan)
        {
            var accessToken = await GraphAuthHelper.GetGraphAccessTokenAsync();
            var restURL = string.Format("{0}planner/plans/{1}", SettingsHelper.GraphResourceUrl, myPlan.id);
            dynamic postPlanJSON = new JObject();
            postPlanJSON.title = myPlan.title;
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    byte[] btBodys = Encoding.UTF8.GetBytes(postPlanJSON.ToString());
                    HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create(restURL);
                    request.Method = "PATCH";
                    request.Accept = "application/json";
                    request.ContentType = "application/json";
                    request.Headers.Add("Authorization", "Bearer " + accessToken);
                    request.Headers.Add("If-Match", myPlan.Etag);
                    request.GetRequestStream().Write(btBodys, 0, btBodys.Length);
                    using (HttpWebResponse response = await request.GetResponseAsync() as HttpWebResponse)
                    {
                        if (response.StatusCode == HttpStatusCode.NoContent)
                        {
                            //update successfully
                        }
                    }
                }
            }
            catch (Exception el)
            {
                el.ToString();
            }
        }
    }
}