using A2B_App.Shared;
using A2B_App.Shared.User;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace A2B_App.Client.Services
{
    public class AdminService
    {

        private readonly ClientSettings settings;

        public AdminService(ClientSettings _settings)
        {
            settings = _settings;
        }

        public async Task<(HttpResponseMessage, string)> UserManagement(AppUser appUser, HttpClient Http, string title)
        {
            string url = string.Empty;
            HttpRequestMessage httpRequest = null;
            switch (title)
            {
                case "Add User":
                    url = $"api/Admin/admin/user/create";
                    httpRequest = new HttpRequestMessage(new HttpMethod("POST"), url);
                    break;
                case "Update User":
                    url = $"api/Admin/admin/user/update";
                    httpRequest = new HttpRequestMessage(new HttpMethod("PUT"), url);
                    break;
                case "Delete User":
                    url = $"api/Admin/admin/user/delete";
                    httpRequest = new HttpRequestMessage(new HttpMethod("DELETE"), url);
                    break;
                case "Add Role":
                    url = $"api/Admin/admin/user/updaterole";
                    httpRequest = new HttpRequestMessage(new HttpMethod("PUT"), url);
                    break;
                case "Delete Role":
                    url = $"api/Admin/admin/user/deleterole";
                    httpRequest = new HttpRequestMessage(new HttpMethod("PUT"), url);
                    break;
            }

            using (var request = httpRequest)
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(appUser));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                //Debug.WriteLine(response.Content.ReadAsStringAsync().Result);
                //Debug.WriteLine(response.StatusCode.ToString());
                return (response, title);
            }

        }

        public async Task<HttpResponseMessage> GetUser(PageTableFilter tableFilter, HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Admin/admin/user"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(tableFilter));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                Debug.WriteLine(response.Content.ReadAsStringAsync().Result);
                Debug.WriteLine(response.StatusCode.ToString());
                return response;
            }

        }

        public async Task<HttpResponseMessage> GetExternalLink(HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("GET"), $"api/Config/externalRef"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");
                var response = await Http.SendAsync(request);
                return response;
            }

        }

        public async Task<HttpResponseMessage> GetData(HttpClient Http, string url)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("GET"), $"{url}"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");
                var response = await Http.SendAsync(request);
                return response;
            }

        }

        public async Task<(HttpResponseMessage, string)> RoleManagement(AppRole appRole, HttpClient Http, string title)
        {
            string url = string.Empty;
            HttpRequestMessage httpRequest = null;
            switch (title)
            {
                case "Get Roles":
                    url = $"api/Admin/admin/roles";
                    httpRequest = new HttpRequestMessage(new HttpMethod("GET"), url);
                    break;
                case "Add Role":
                    url = $"api/Admin/admin/roles/create";
                    httpRequest = new HttpRequestMessage(new HttpMethod("POST"), url);
                    break;
                case "Update Role":
                    url = $"api/Admin/admin/roles/update";
                    httpRequest = new HttpRequestMessage(new HttpMethod("PUT"), url);
                    break;
                case "Delete Role":
                    url = $"api/Admin/admin/roles/delete";
                    httpRequest = new HttpRequestMessage(new HttpMethod("DELETE"), url);
                    break;
            }

            using (var request = httpRequest)
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                if(title != "Get Roles") //Get cannot have json body, so we exclude GET request
                {
                    request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(appRole));
                    request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");
                }

                var response = await Http.SendAsync(request);
                //Debug.WriteLine(response.Content.ReadAsStringAsync().Result);
                //Debug.WriteLine(response.StatusCode.ToString());
                return (response, title);
            }

        }


    }
}
