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
using Microsoft.IdentityModel.Protocols;
using System.Configuration;
using Microsoft.Graph;

namespace Office365PlannerTask.Models
{
    public class MyTasksRepository
    {
        //public async Task<string> GetGraphAccessTokenAsync()
        //{
        //    var signInUserId = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;
        //    var userObjectId = ClaimsPrincipal.Current.FindFirst(SettingsHelper.ClaimTypeObjectIdentifier).Value;

        //    var clientCredential = new ClientCredential(SettingsHelper.ClientId, SettingsHelper.ClientSecret);
        //    var userIdentifier = new UserIdentifier(userObjectId, UserIdentifierType.UniqueId);

        //    // create auth context
        //    AuthenticationContext authContext = new AuthenticationContext(SettingsHelper.AzureAdAuthority, new ADALTokenCache(signInUserId));
        //    var result = await authContext.AcquireTokenSilentAsync(SettingsHelper.AzureAdGraphResourceURL, clientCredential, userIdentifier);

        //    return result.AccessToken;
        //}

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

        public async Task<List<MyTask>> GetTasks(string planid)
        {
            if (UseSDK)
            {
                return await GetTasksSDK(planid);
            }
            else
            {
                return await GetTasksREST(planid);
            }
        }

        private async Task<List<MyTask>> GetTasksSDK(string planid)
        {
            var tasksResult = new List<MyTask>();

            try
            {
                var graphServiceClient = await GetGraphServiceAsync();

                var tasks = await graphServiceClient.Me.Planner.Plans[planid].Tasks.Request().GetAsync();

                foreach (var item in tasks)
                {
                    tasksResult.Add(
                        new MyTask
                        {
                            id = item.Id,
                            title = item.Title,
                            percentComplete = item.PercentComplete ?? 0,
                            Etag = item.GetEtag(),
                            planId = item.PlanId
                        }
                        );

                }
            }
            catch (Exception)
            {
            }

            return tasksResult;

        }

        public async Task<List<MyTask>> GetTasksREST(string planid)
        {
            var tasksResult = new List<MyTask>();
            var accessToken = await GraphAuthHelper.GetGraphAccessTokenAsync();
            var restURL = string.Format("{0}planner/plans/{1}/tasks", SettingsHelper.GraphResourceUrl, planid);
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
                                tasksResult.Add(new MyTask
                                {
                                    id = item["id"].ToString(),
                                    title = item["title"].ToString(),
                                    percentComplete = !string.IsNullOrEmpty(item["percentComplete"].ToString()) ? Convert.ToInt32(item["percentComplete"].ToString()) : 0,
                                    planId = planid,
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

            return tasksResult;
        }

        public async Task<MyTask> GetTask(string id)
        {
            if (UseSDK)
            {
                return await GetTaskSDK(id);
            }
            else
            {
                return await GetTaskREST(id);
            }
        }

        private async Task<MyTask> GetTaskSDK(string id)
        {
            MyTask task = new MyTask();

            var graphServiceClient = await GetGraphServiceAsync();

            var plannerTask = await graphServiceClient.Planner.Tasks[id].Request().GetAsync();

            if (plannerTask != null)
            {
                task.title = plannerTask.Title;
                task.id = plannerTask.Id;
                task.planId = plannerTask.PlanId;
                task.percentComplete = plannerTask.PercentComplete ?? 0;
                task.Etag = plannerTask.GetEtag();
            }

            return task;
        }

        public async Task<MyTask> GetTaskREST(string id)
        {
            MyTask task = new MyTask();
            var accessToken = await GraphAuthHelper.GetGraphAccessTokenAsync();
            var restURL = string.Format("{0}planner/tasks/{1}", SettingsHelper.GraphResourceUrl, id);
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
                                task.id = item["id"].ToString();
                                task.title = item["title"].ToString();
                                task.percentComplete = !string.IsNullOrEmpty(item["percentComplete"].ToString()) ? Convert.ToInt32(item["percentComplete"].ToString()) : 0;
                                task.planId = item["planId"].ToString();
                                task.Etag = !string.IsNullOrEmpty(item["@odata.etag"].ToString()) ? item["@odata.etag"].ToString() : "";
                            }
                        }
                    }
                }
            }
            catch (Exception el)
            {
                el.ToString();
            }

            return task;
        }

        public async Task CreateTask(MyTask myTask)
        {
            if (UseSDK)
            {
                await CreateTaskSDK(myTask);
            }
            else
            {
                await CreateTaskREST(myTask);
            }
        }

        private async Task CreateTaskSDK(MyTask myTask)
        {
            if (myTask == null)
            {
                return;
            }

            if (String.IsNullOrWhiteSpace(myTask.planId))
            {
                return;
            }

            if (String.IsNullOrWhiteSpace(myTask.title))
            {
                return;
            }

            try
            {
                var graphServiceClient = await GetGraphServiceAsync();

                var plannerTask = new PlannerTask
                {
                    Title = myTask.title,
                    PlanId = myTask.planId,
                    PercentComplete = myTask.percentComplete
                };

                await graphServiceClient.Planner.Tasks.Request().AddAsync(plannerTask);
            }
            catch (Exception)
            {
            }

        }

        public async Task CreateTaskREST(MyTask myTask)
        {
            var accessToken = await GraphAuthHelper.GetGraphAccessTokenAsync();
            var restURL = string.Format("{0}/planner/tasks", SettingsHelper.GraphResourceUrl);
            dynamic postTaskJSON = new JObject();
            postTaskJSON.title = myTask.title;
            postTaskJSON.percentComplete = myTask.percentComplete;
            postTaskJSON.planId = myTask.planId;
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);


                    var requestMessage = new HttpRequestMessage(HttpMethod.Post, restURL);
                    requestMessage.Content = new StringContent(postTaskJSON.ToString(), System.Text.Encoding.UTF8, "application/json");
                    using (var response = await client.SendAsync(requestMessage))
                    {
                        if (response.IsSuccessStatusCode)
                            return;
                        else
                            throw new Exception("add task error: " + response.StatusCode);
                    }
                }
            }
            catch (Exception el)
            {
                el.ToString();
            }
        }

        public async Task UpdateTask(MyTask myTask)
        {
            if (UseSDK)
            {
                await UpdateTaskSDK(myTask);
            }
            else
            {
                await UpdateTaskREST(myTask);
            }
        }

        private async Task UpdateTaskSDK(MyTask myTask)
        {
            try
            {
                var graphServiceClient = await GetGraphServiceAsync();

                var taskToUpdate = await graphServiceClient.Planner.Tasks[myTask.id].Request().GetAsync();

                string etag = taskToUpdate.GetEtag();
                taskToUpdate.Title = myTask.title;
                taskToUpdate.PlanId = myTask.planId;

                //await graphServiceClient.Planner.Plans[reqPlan.Id].Request().Header("If-Match", etag).Header("Prefer", "return=representation").UpdateAsync(plan);

                await graphServiceClient.Planner.Tasks[myTask.id].Request().Header("If-Match", etag).UpdateAsync(taskToUpdate);

            }
            catch (Exception)
            {

                //throw;
            }
        }

        public async Task UpdateTaskREST(MyTask myTask)
        {
            var accessToken = await GraphAuthHelper.GetGraphAccessTokenAsync();
            var restURL = string.Format("{0}planner/tasks/{1}", SettingsHelper.GraphResourceUrl, myTask.id);
            dynamic postTaskJSON = new JObject();
            postTaskJSON.title = myTask.title;
            postTaskJSON.percentComplete = myTask.percentComplete;
            try
            {
                byte[] btBodys = Encoding.UTF8.GetBytes(postTaskJSON.ToString());
                HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create(restURL);
                request.Method = "PATCH";
                request.Accept = "application/json";
                request.ContentType = "application/json";
                request.Headers.Add("Authorization", "Bearer " + accessToken);
                request.Headers.Add("If-Match", myTask.Etag);
                request.GetRequestStream().Write(btBodys, 0, btBodys.Length);
                using (HttpWebResponse response = await request.GetResponseAsync() as HttpWebResponse)
                {
                    if (response.StatusCode == HttpStatusCode.NoContent)
                    {
                        //update successfully
                    }
                }
            }
            catch (Exception el)
            {
                el.ToString();
            }
        }

        public async Task DeleteTask(string id, string eTag)
        {
            if (UseSDK)
            {
                await DeleteTaskSDK(id, eTag);
            }
            else
            {
                await DeleteTaskREST(id, eTag);

            }
        }

        private async Task DeleteTaskSDK(string id, string eTag)
        {
            var graphServiceClient = await GetGraphServiceAsync();

            await graphServiceClient.Planner.Tasks[id].Request().Header("If-Match", eTag).DeleteAsync();

        }

        public async Task DeleteTaskREST(string id, string eTag)
        {
            var accessToken = await GraphAuthHelper.GetGraphAccessTokenAsync();
            var restURL = string.Format("{0}planner/tasks/{1}", SettingsHelper.GraphResourceUrl, id);
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    var accept = "application/json";

                    client.DefaultRequestHeaders.Add("Accept", accept);
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    client.DefaultRequestHeaders.Add("If-Match", eTag);

                    using (var response = await client.DeleteAsync(restURL))
                    {
                        if (response.IsSuccessStatusCode)
                            return;
                        else
                            throw new Exception("delete task error: " + response.StatusCode);
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