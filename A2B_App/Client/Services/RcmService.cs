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
    public class RcmService
    {

        private readonly ClientSettings settings;
        public RcmService(ClientSettings clientSettings)
        {
            settings = clientSettings;
        }

        public async Task<HttpResponseMessage> GetRCMData(HttpClient Http, string url)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("GET"), $"{url}"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");
                var response = await Http.SendAsync(request);
                return response;
            }

        }

        public async Task<Rcm> GetRcmByPodioItemId(string PodioItemId, HttpClient Http)
        {
            //HttpClient Http = new HttpClient();
            //ClientSettings settings = new ClientSettings();
            //string server = settings.GetApiServer();
            Rcm RCM = new Rcm();
            //RCM = await Http.GetFromJsonAsync<Rcm>($"api/Rcm/rcm/get/{PodioItemId}");

            //Updated 5.13.2021 | Levin jay Tagapan
            var response = await GetRCMData(Http, $"api/Rcm/rcm/get/{PodioItemId}");
            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                string result = response.Content.ReadAsStringAsync().Result.ToString();
                RCM = JsonConvert.DeserializeObject<Rcm>(result);
            }

            return RCM;
        }

        public async Task<SampleSelection> GetSampleSelectionByRcmPodioItemId(string PodioItemId, HttpClient Http)
        {
            //HttpClient Http = new HttpClient();
            //ClientSettings settings = new ClientSettings();
            //string server = settings.GetApiServer();
            SampleSelection sampleSelection = new SampleSelection();
            //sampleSelection = await Http.GetFromJsonAsync<SampleSelection>($"api/SampleSelection/data/rcm/sampleselection/{PodioItemId}");

            //Updated 5.13.2021 | Levin jay Tagapan
            var response = await GetRCMData(Http, $"api/SampleSelection/data/rcm/sampleselection/{PodioItemId}");
            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                string result = response.Content.ReadAsStringAsync().Result.ToString();
                sampleSelection = JsonConvert.DeserializeObject<SampleSelection>(result);
            }

            return sampleSelection;
        }

        public async Task<HttpResponseMessage> GetRcmPodioFilterBy(RcmItemFilter filter, HttpClient Http)
        {

            //ClientSettings settings = new ClientSettings();
            //string server = settings.GetApiServer();
           
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Rcm/rcm/get"))
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

        public async Task<List<string>> GetClient(HttpClient Http)
        {
            //HttpClient Http = new HttpClient();
            //ClientSettings settings = new ClientSettings();
            //string server = settings.GetApiServer();
            List<string> listClient = new List<string>();
            //listClient = await Http.GetFromJsonAsync<List<string>>($"api/Rcm/rcm/clients");

            //Updated 5.13.2021 | Levin jay Tagapan
            var response = await GetRCMData(Http, $"api/Rcm/rcm/clients");
            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                string result = response.Content.ReadAsStringAsync().Result.ToString();
                listClient = JsonConvert.DeserializeObject<List<string>>(result);
            }

            return listClient;
        }

        public async Task<List<string>> GetControlName(HttpClient Http)
        {
            //HttpClient Http = new HttpClient();
            //ClientSettings settings = new ClientSettings();
            //string server = settings.GetApiServer();
            List<string> listControlName = new List<string>();
            //listControlName = await Http.GetFromJsonAsync<List<string>>($"api/Rcm/rcm/controlname");

            //Updated 5.13.2021 | Levin jay Tagapan
            var response = await GetRCMData(Http, $"api/Rcm/rcm/controlname");
            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                string result = response.Content.ReadAsStringAsync().Result.ToString();
                listControlName = JsonConvert.DeserializeObject<List<string>>(result);
            }

            return listControlName;
        }

        public async Task<List<string>> GetRcmQuestionnaireFy(HttpClient Http)
        {
            List<string> listFy = new List<string>();
            //listFy = await Http.GetFromJsonAsync<List<string>>($"api/Rcm/q1fy");

            //Updated 5.13.2021 | Levin jay Tagapan
            var response = await GetRCMData(Http, $"api/Rcm/q1fy");
            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                string result = response.Content.ReadAsStringAsync().Result.ToString();
                listFy = JsonConvert.DeserializeObject<List<string>>(result);
            }

            return listFy;
        }

        public async Task<List<string>> GetRcmFy(HttpClient Http)
        {
            List<string> listFy = new List<string>();
            //listFy = await Http.GetFromJsonAsync<List<string>>($"api/Rcm/fy");

            //Updated 5.13.2021 | Levin jay Tagapan
            var response = await GetRCMData(Http, $"api/Rcm/fy");
            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                string result = response.Content.ReadAsStringAsync().Result.ToString();
                listFy = JsonConvert.DeserializeObject<List<string>>(result);
            }

            return listFy;
        }

        public async Task<HttpResponseMessage> GetRcmClient(RcmQuestionnaireFilter filter, HttpClient Http)
        {
            //List<string> listFy;
            //listFy = await Http.GetFromJsonAsync<List<string>>($"api/Rcm/q1fy");
            //return listFy;

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Rcm/clientbyyear"))
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

        public async Task<HttpResponseMessage> GetRcmQuestionnaireClient(RcmQuestionnaireFilter filter, HttpClient Http)
        {

    
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Rcm/q2client"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(filter));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                //Debug.WriteLine($"Response Result: {response.Content.ReadAsStringAsync().Result}");
                //Debug.WriteLine($"Response Status Code: {response.StatusCode}");
                return response;
            }

        }

        public async Task<HttpResponseMessage> GetRcmQuestionnaireProcess(RcmQuestionnaireFilter filter, HttpClient Http)
        {
            //List<string> listFy;
            //listFy = await Http.GetFromJsonAsync<List<string>>($"api/Rcm/q1fy");
            //return listFy;

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Rcm/q3process"))
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

        public async Task<HttpResponseMessage> GetRcmQuestionnaireSubProcess(RcmQuestionnaireFilter filter, HttpClient Http)
        {
            //List<string> listFy;
            //listFy = await Http.GetFromJsonAsync<List<string>>($"api/Rcm/q1fy");
            //return listFy;

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Rcm/q4subprocess"))
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

        public async Task<HttpResponseMessage> GetRcmQuestionnaireControlId(RcmQuestionnaireFilter filter, HttpClient Http)
        {
            //List<string> listFy;
            //listFy = await Http.GetFromJsonAsync<List<string>>($"api/Rcm/q1fy");
            //return listFy;

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Rcm/q9controlid"))
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

        public async Task<HttpResponseMessage> GetRcmQuestionnaireQ13toQ19(RcmQuestionnaireFilter filter, HttpClient Http)
        {
            //List<string> listFy;
            //listFy = await Http.GetFromJsonAsync<List<string>>($"api/Rcm/q1fy");
            //return listFy;

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Rcm/q13toq19"))
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


        public async Task<HttpResponseMessage> GetRcmControl(RcmQuestionnaireFilter filter, HttpClient Http)
        {
            //List<string> listFy;
            //listFy = await Http.GetFromJsonAsync<List<string>>($"api/Rcm/q1fy");
            //return listFy;

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Rcm/getcontrol"))
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

        public async Task<HttpResponseMessage> SaveRcmControl(Rcm rcm, HttpClient Http)
        {
            //List<string> listFy;
            //listFy = await Http.GetFromJsonAsync<List<string>>($"api/Rcm/q1fy");
            //return listFy;

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Rcm/save"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(rcm));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                Debug.WriteLine($"Response Result: {response.Content.ReadAsStringAsync().Result}");
                Debug.WriteLine($"Response Status Code: {response.StatusCode}");
                return response;
            }
        }
        public async Task<HttpResponseMessage> CheckDuplicatesRcmControl(Rcm rcm, HttpClient Http)
        {
            //List<string> listFy;
            //listFy = await Http.GetFromJsonAsync<List<string>>($"api/Rcm/q1fy");
            //return listFy;

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Rcm/check-duplicates"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(rcm));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                Debug.WriteLine($"Response Result: {response.Content.ReadAsStringAsync().Result}");
                Debug.WriteLine($"Response Status Code: {response.StatusCode}");
                return response;
            }
        }

        public async Task<HttpResponseMessage> SavePodioRcmControl(Rcm rcm, HttpClient Http)
        {
            //List<string> listFy;
            //listFy = await Http.GetFromJsonAsync<List<string>>($"api/Rcm/q1fy");
            //return listFy;

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Rcm/podiosave"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(rcm));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                Debug.WriteLine($"Response Result: {response.Content.ReadAsStringAsync().Result}");
                Debug.WriteLine($"Response Status Code: {response.StatusCode}");
                return response;
            }

        }

        public async Task<HttpResponseMessage> DeleteRcmControl(Rcm rcm, HttpClient Http)
        {
            //List<string> listFy;
            //listFy = await Http.GetFromJsonAsync<List<string>>($"api/Rcm/q1fy");
            //return listFy;

            using (var request = new HttpRequestMessage(new HttpMethod("DELETE"), $"api/Rcm/delete/{rcm.PodioItemId}"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                //request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(rcm));
                //request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                Debug.WriteLine($"Response Result: {response.Content.ReadAsStringAsync().Result}");
                Debug.WriteLine($"Response Status Code: {response.StatusCode}");
                return response;
            }

        }

        public async Task<HttpResponseMessage> UpdateRcmControl(Rcm rcm, HttpClient Http)
        {
            //List<string> listFy;
            //listFy = await Http.GetFromJsonAsync<List<string>>($"api/Rcm/q1fy");
            //return listFy;

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Rcm/update"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(rcm));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                Debug.WriteLine($"Response Result: {response.Content.ReadAsStringAsync().Result}");
                Debug.WriteLine($"Response Status Code: {response.StatusCode}");
                return response;
            }

        }

        public async Task<HttpResponseMessage> UpdatePodioRcmControl(Rcm rcm, HttpClient Http)
        {
            //List<string> listFy;
            //listFy = await Http.GetFromJsonAsync<List<string>>($"api/Rcm/q1fy");
            //return listFy;

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Rcm/podioupdate"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(rcm));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                Debug.WriteLine($"Response Result: {response.Content.ReadAsStringAsync().Result}");
                Debug.WriteLine($"Response Status Code: {response.StatusCode}");
                return response;
            }

        }

        public async Task<HttpResponseMessage> GenerateRcmControl(string clientName, string Fy, HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Rcm/generate/{clientName}/{Fy}"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                //request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(rcm));
                //request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

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

        public async Task<HttpResponseMessage> GetRcm(KeyReportFilter filter, HttpClient Http)
        {
            //List<string> listFy;
            //listFy = await Http.GetFromJsonAsync<List<string>>($"api/Rcm/q1fy");
            //return listFy;

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Rcm/rcm"))
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
        public async Task<HttpResponseMessage> fetch_rcm(HttpClient Http, String FY_RCM, String ClientName_RCM)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("GET"), $"api/Rcm/view_rcm/{FY_RCM}/{ClientName_RCM}"))
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
