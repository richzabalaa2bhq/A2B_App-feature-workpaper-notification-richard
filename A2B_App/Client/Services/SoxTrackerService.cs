using A2B_App.Shared;
using A2B_App.Shared.Sox;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Threading.Tasks;


namespace A2B_App.Client.Services
{
    public class SoxTrackerService
    {

        private readonly ClientSettings settings;
        public SoxTrackerService(ClientSettings clientSettings)
        {
            settings = clientSettings;
        }

        public async Task<HttpResponseMessage> GetTrackerData(HttpClient Http, string url)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("GET"), $"{url}"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");
                var response = await Http.SendAsync(request);
                return response;
            }

        }

        public async Task<List<string>> GetSoxTrackerFy(HttpClient Http)
        {
            List<string> listFy = new List<string>();
            //listFy = await Http.GetFromJsonAsync<List<string>>($"api/SoxTracker/fy");

            //Updated 5.13.2021 | Levin jay Tagapan
            var response = await GetTrackerData(Http, $"api/SoxTracker/fy");
            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                string result = response.Content.ReadAsStringAsync().Result.ToString();
                listFy = JsonConvert.DeserializeObject<List<string>>(result);
            }

            return listFy;
        }

        public async Task<HttpResponseMessage> GetSoxTrackerClient(RcmQuestionnaireFilter filter, HttpClient Http)
        {
            //List<string> listFy;
            //listFy = await Http.GetFromJsonAsync<List<string>>($"api/Rcm/q1fy");
            //return listFy;

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/SoxTracker/clientbyyear"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(filter));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                Debug.WriteLine($"Response Result GetSoxTrackerClient: {response.Content.ReadAsStringAsync().Result}");
                Debug.WriteLine($"Response Status Code: {response.StatusCode}");
                return response;
            }

        }

        public async Task<HttpResponseMessage> GenerateSoxTrackerControl(string clientName, string Fy, HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/SoxTracker/generate/{clientName}/{Fy}"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                //request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(rcm));
                //request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                Debug.WriteLine($"Response Result GenerateSoxTrackerControl: {response.Content.ReadAsStringAsync().Result}");
                Debug.WriteLine($"Response Status Code: {response.StatusCode}");
                return response;
            }

        }

        public async Task<HttpResponseMessage> GetSoxTracker(RcmQuestionnaireFilter filter, HttpClient Http)
        {
            //List<string> listFy;
            //listFy = await Http.GetFromJsonAsync<List<string>>($"api/Rcm/q1fy");
            //return listFy;

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/SoxTracker/control"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(filter));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                Debug.WriteLine($"Response Result GetSoxTracker: {response.Content.ReadAsStringAsync().Result}");
                Debug.WriteLine($"Response Status Code: {response.StatusCode}");
                return response;
            }

        }

        public async Task<HttpResponseMessage> GetSoxTracker2(KeyReportFilter filter, HttpClient Http)
        {
            //List<string> listFy;
            //listFy = await Http.GetFromJsonAsync<List<string>>($"api/Rcm/q1fy");
            //return listFy;

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/SoxTracker/tracker"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(filter));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                Debug.WriteLine($"Response Result: {response.Content.ReadAsStringAsync().Result}");
                Debug.WriteLine($"Response Status Code: {response.StatusCode}");
                return response;
            }

        }


        public async Task<HttpResponseMessage> UploadToSharefile(SharefileItem item, HttpClient Http)
        {
            //List<string> listFy;
            //listFy = await Http.GetFromJsonAsync<List<string>>($"api/Rcm/q1fy");
            //return listFy;

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Sharefile/upload"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(item));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                Debug.WriteLine($"Response Result: {response.Content.ReadAsStringAsync().Result}");
                Debug.WriteLine($"Response Status Code: {response.StatusCode}");
                return response;
            }

        }

        public async Task<SoxTrackerQuestionnaire> GetSoxTrackerQuestionnaire(HttpClient Http)
        {
            SoxTrackerQuestionnaire soxTrackerQuestionnaire = new SoxTrackerQuestionnaire();
            //soxTrackerQuestionnaire = await Http.GetFromJsonAsync<SoxTrackerQuestionnaire>($"api/SoxTracker/questionnaire");

            //Updated 5.13.2021 | Levin jay Tagapan
            var response = await GetTrackerData(Http, $"api/SoxTracker/questionnaire");
            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                string result = response.Content.ReadAsStringAsync().Result.ToString();
                soxTrackerQuestionnaire = JsonConvert.DeserializeObject<SoxTrackerQuestionnaire>(result);
            }

            return soxTrackerQuestionnaire;
        }

        public async Task<HttpResponseMessage> CreatePodioSoxTracker(SoxTracker soxTracker, HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/SoxTracker/podiocreate"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(soxTracker));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                Debug.WriteLine($"Response Result: {response.Content.ReadAsStringAsync().Result}");
                Debug.WriteLine($"Response Status Code: {response.StatusCode}");
                return response;
            }

        }

        public async Task<HttpResponseMessage> SaveSoxTrackerDB(SoxTracker soxTracker, HttpClient Http)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/SoxTracker/save"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(soxTracker));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                Debug.WriteLine($"Response Result: {response.Content.ReadAsStringAsync().Result}");
                Debug.WriteLine($"Response Status Code: {response.StatusCode}");
                return response;
            }

        }

        public async Task<HttpResponseMessage> UpdatePodioSoxTrackerControl(SoxTracker soxTracker, HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/SoxTracker/podioupdate"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(soxTracker));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                Debug.WriteLine($"Response Result: {response.Content.ReadAsStringAsync().Result}");
                Debug.WriteLine($"Response Status Code: {response.StatusCode}");
                return response;
            }

        }
        public async Task<HttpResponseMessage> fetch_sox_tracker(HttpClient Http, String FY_SoxTracker, String ClientName_SoxTracker)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("GET"), $"api/SoxTracker/view_sox_tracker/{FY_SoxTracker}/{ClientName_SoxTracker}"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                ///request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject());
                //request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                Debug.WriteLine($"Response Result: {response.Content.ReadAsStringAsync().Result}");
                Debug.WriteLine($"Response Status Code: {response.StatusCode}");
                return response;
            }

        }

        
    }
}
