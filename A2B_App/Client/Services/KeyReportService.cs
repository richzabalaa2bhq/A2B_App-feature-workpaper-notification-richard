using A2B_App.Shared;
using A2B_App.Shared.Sox;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Threading.Tasks;

namespace A2B_App.Client.Services
{
    public class KeyReportService
    {
        private readonly ClientSettings settings;
        public KeyReportService(ClientSettings clientSettings)
        {
            settings = clientSettings;
        }

        public async Task<HttpResponseMessage> GetKeyReportControlActivity(HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("GET"), $"api/KeyReport/getControlActivity"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");
                var response = await Http.SendAsync(request);
                return response;
            }
        }

        public async Task<HttpResponseMessage> GetKeyReportKeyControl(HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("GET"), $"api/KeyReport/getKeyControl"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");
                var response = await Http.SendAsync(request);
                return response;
            }
        }


        public async Task<HttpResponseMessage> GetKeyReportName(HttpClient Http)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("GET"), $"api/KeyReport/getReportName"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");
                var response = await Http.SendAsync(request);
                return response;
            }
        }

        public async Task<HttpResponseMessage> GetKeyReportSystemSource(HttpClient Http)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("GET"), $"api/KeyReport/getSystemSource"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");
                var response = await Http.SendAsync(request);
                return response;
            }
        }

        public async Task<HttpResponseMessage> GetKeyReportNonKeyReport(HttpClient Http)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("GET"), $"api/KeyReport/getKeyOrNonKeyReport"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");
                var response = await Http.SendAsync(request);
                return response;
            }
        }

        public async Task<HttpResponseMessage> GetKeyReportReportCustomized(HttpClient Http)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("GET"), $"api/KeyReport/getReportCustomized"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");
                var response = await Http.SendAsync(request);
                return response;
            }
        }

        public async Task<HttpResponseMessage> GetKeyReportIUCType(HttpClient Http)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("GET"), $"api/KeyReport/getIUCType"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");
                var response = await Http.SendAsync(request);
                return response;
            }
        }

        public async Task<HttpResponseMessage> GetKeyReportControlsRelyingIUC(HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("GET"), $"api/KeyReport/getControlsRelyingIUC"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");
                var response = await Http.SendAsync(request);
                return response;
            }
        }

        public async Task<HttpResponseMessage> GetKeyReportPreparer(HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("GET"), $"api/KeyReport/getPreparer"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");
                var response = await Http.SendAsync(request);
                return response;
            }
        }

        public async Task<HttpResponseMessage> GetKeyReportUniqueKeyReport(HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("GET"), $"api/KeyReport/getUniqueKeyReport"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");
                var response = await Http.SendAsync(request);
                return response;
            }
        }

        public async Task<HttpResponseMessage> GetKeyReportNotes(HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("GET"), $"api/KeyReport/getNotes"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");
                var response = await Http.SendAsync(request);
                return response;
            }
        }

        public async Task<HttpResponseMessage> GetKeyReportNumber(HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("GET"), $"api/KeyReport/getNumber"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");
                var response = await Http.SendAsync(request);
                return response;
            }
        }

        public async Task<HttpResponseMessage> GetKeyReportTester(HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("GET"), $"api/KeyReport/getTester"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");
                var response = await Http.SendAsync(request);
                return response;
            }
        }

        public async Task<HttpResponseMessage> GetKeyReportReviewer(HttpClient Http)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("GET"), $"api/KeyReport/getReviewer"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");
                var response = await Http.SendAsync(request);
                return response;
            }
        }

        public async Task<HttpResponseMessage> GetKeyReportItem(HttpClient Http, string apiName)
        {
            string url = string.Empty;
            switch (apiName)
            {
                case "GetRcmControlId":
                    url = $"api/KeyReport/getRcmControlId";
                    break;
                case "GetRcmProcess":
                    url = $"api/KeyReport/getRcmProcess";
                    break;
                case "GetRcmControlOwner":
                    url = $"api/KeyReport/getRcmControlOwner";
                    break;
                default:
                    url = string.Empty;
                    break;
            }

            using (var request = new HttpRequestMessage(new HttpMethod("GET"), url))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");
                var response = await Http.SendAsync(request);
                return response;
            }
        }

        public async Task<HttpResponseMessage> SavePodioKeyReportItem(HttpClient Http, string apiName, KeyReportIds listUserInput)
        {
            string url = string.Empty;
            switch (apiName)
            {
                case "SavePodioKeyReportOrigFormat":
                    url = $"api/KeyReport/savepodio/origformat";
                    break;
                case "SavePodioKeyReportAllIUC":
                    url = $"api/KeyReport/savepodio/alliuc";
                    break;
                case "SavePodioKeyReportTestStatus":
                    url = $"api/KeyReport/savepodio/teststatus";
                    break;
                case "SavePodioKeyReportException":
                    url = $"api/KeyReport/savepodio/exception";
                    break;
                case "SavePodioKeyReportLeadsheet":
                    url = $"api/KeyReport/savepodio/leadsheet";
                    break;
                case "SaveDBKeyReport":
                    url = $"api/KeyReport/save/keyreport";
                    break;
                case "SaveDBKeyReportIfNew":
                    url = $"api/KeyReport/save/keyreport_new";
                    break;
                default:
                    url = string.Empty;
                    break;
            }

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), url))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");
                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(listUserInput));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                //System.Diagnostics.Debug.WriteLine(response.Content.ReadAsStringAsync().Result);
                //System.Diagnostics.Debug.WriteLine(response.StatusCode.ToString());
                return response;
            }


        }


        public async Task<HttpResponseMessage> SavePodioKeyReportItem2(HttpClient Http, string apiName, 
                                                                       List<KeyReportUserInput> listUserInput)
        {
            string url = string.Empty;
           
            switch (apiName)
            {
                case "SavePodioKeyReportOrigFormat":
                    url = $"api/KeyReport/savepodio/origformat";
                    
                                       
                    break;
                case "SavePodioKeyReportAllIUC":
                    url = $"api/KeyReport/savepodio/alliuc";
                    
                    break;
                case "SavePodioKeyReportTestStatus":
                    url = $"api/KeyReport/savepodio/teststatus";
                    break;
                case "SavePodioKeyReportException":
                    url = $"api/KeyReport/savepodio/exception";
                    break;
                case "SavePodioKeyReportLeadsheet":
                    url = $"api/KeyReport/savepodio/leadsheet";
                    break;
                case "SaveDBKeyReport":
                    url = $"api/KeyReport/save/keyreport";
                    break;
                default:
                    url = string.Empty;
                    break;
            }

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), url))
            {
                
                
               
                request.Headers.TryAddWithoutValidation("accept", "text/plain");
                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(listUserInput));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                //System.Diagnostics.Debug.WriteLine(response.Content.ReadAsStringAsync().Result);
                //System.Diagnostics.Debug.WriteLine(response.StatusCode.ToString());
                return response;
            }


        }

        public async Task<HttpResponseMessage> GenerateLeadsheetOutputFile(KeyReportScreenshot filter, HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/KeyReport/leadsheet/generate"))
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

        public async Task<HttpResponseMessage> GenerateLeadsheetOutputFilePWC(KeyReportScreenshot filter, HttpClient Http)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/KeyReport/leadsheetPWC/generatePWC"))
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


        public async Task<HttpResponseMessage> SaveScreenshot(List<KeyReportScreenshotUpload> listScreenshot, HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/KeyReport/save/screenshot"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(listScreenshot));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                //Debug.WriteLine($"Response Result: {response.Content.ReadAsStringAsync().Result}");
                //Debug.WriteLine($"Response Status Code: {response.StatusCode}");
                return response;
            }

        }


        public async Task<HttpResponseMessage> SaveReportFile(List<KeyReportFile> listReportFile, HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/KeyReport/save/reportfile"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(listReportFile));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                //Debug.WriteLine($"Response Result: {response.Content.ReadAsStringAsync().Result}");
                //Debug.WriteLine($"Response Status Code: {response.StatusCode}");
                return response;
            }

        }



        public async Task<HttpResponseMessage> GenerateKeyReportViewAccess(KeyReportScreenshot filter, HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/KeyReport/leadsheet/fetch"))
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

        public async Task<HttpResponseMessage> GenerateTrackerOutputFile(KeyReportFilter filter, HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/KeyReport/tracker/generate"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");
                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(filter));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");
                var response = await Http.SendAsync(request);
                return response;
            }

        }

        public async Task<HttpResponseMessage> GetFy(HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("GET"), $"api/KeyReport/leadsheet/getfy"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                //request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(filter));
                //request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                //Debug.WriteLine($"Response Result: {response.Content.ReadAsStringAsync().Result}");
                //Debug.WriteLine($"Response Status Code: {response.StatusCode}");
                return response;
            }
        }

        public async Task<HttpResponseMessage> GetClient(KeyReportFilter filter, HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/KeyReport/leadsheet/getclient"))
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

        public async Task<HttpResponseMessage> GetReportName(KeyReportFilter filter, HttpClient Http)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/KeyReport/leadsheet/getreportname"))
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

        public async Task<HttpResponseMessage> GetReportName2(KeyReportFilter filter, HttpClient Http)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/KeyReport/getreportname"))
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

        public async Task<HttpResponseMessage> GetControlId(KeyReportFilter filter, HttpClient Http)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/KeyReport/getcontrolid"))
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

        public async Task<HttpResponseMessage> GetKeyReport(KeyReportFilter filter, HttpClient Http, int num)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/KeyReport/keyreport{num}"))
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

        public async Task<HttpResponseMessage> GetItemId(KeyReportFilter filter, HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/KeyReport/getitemid"))
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

        public async Task<HttpResponseMessage> GetKeyReport2(HttpClient Http, int itemId)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("GET"), $"api/KeyReport/getkeyreportitem/{itemId}"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                //request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(filter));
                //request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                //Debug.WriteLine($"Response Result: {response.Content.ReadAsStringAsync().Result}");
                //Debug.WriteLine($"Response Status Code: {response.StatusCode}");
                return response;
            }

        }

        public async Task<HttpResponseMessage> GetSharefileScreenshot(KeyReportFilter filter, HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/ShareFile/screenshot"))
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

        public async Task<HttpResponseMessage> GetScreenshot(KeyReportFilter filter, HttpClient Http)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/KeyReport/getscreenshot"))
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

        public async Task<HttpResponseMessage> GetReportFile(KeyReportFilter filter, HttpClient Http)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/KeyReport/getreportfile"))
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

        public async Task<HttpResponseMessage> GetMethods(HttpClient Http, String type)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/KeyReport/methods"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");
                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(type));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");
                var response = await Http.SendAsync(request);
                Debug.WriteLine($"Response Result: {response.Content.ReadAsStringAsync().Result}");
                Debug.WriteLine($"Response Status Code: {response.StatusCode}");
                return response;
            }
        }
        public async Task<HttpResponseMessage> GetQuestions(HttpClient Http, String apiName, KeyReportQuestionsFilter filter)
        {
            string url = string.Empty;
            switch (apiName)
            {
                case "Parameter":
                    url = $"api/KeyReport/fetch/parameters";
                    break;
                case "Report":
                    url = $"api/KeyReport/fetch/reports";
                    break;
                case "Completeness":
                    url = $"api/KeyReport/fetch/completeness";
                    break;
                case "Accuracy":
                    url = $"api/KeyReport/fetch/accuracy";
                    break;
                default:
                    url = string.Empty;
                    break;
            }


            using (var request = new HttpRequestMessage(new HttpMethod("POST"), url))
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
        public async Task<HttpResponseMessage> GenerateTrackerOriginalFormat(KeyReportFilter filter, HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/KeyReport/tracker/generate-records-orig-format"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");
                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(filter));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");
                var response = await Http.SendAsync(request);
                return response;
            }

        }
        public async Task<HttpResponseMessage> GenerateTrackerAllIuc(KeyReportFilter filter, HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/KeyReport/tracker/generate-records-all-iuc"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");
                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(filter));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");
                var response = await Http.SendAsync(request);
                return response;
            }

        }
        public async Task<HttpResponseMessage> GenerateTestStatusTracker(KeyReportFilter filter, HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/KeyReport/tracker/generate-records-test-status-tracker"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");
                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(filter));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");
                var response = await Http.SendAsync(request);
                return response;
            }

        }
        public async Task<HttpResponseMessage> GenerateListExceptions(KeyReportFilter filter, HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/KeyReport/tracker/generate-records-exceptions"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");
                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(filter));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");
                var response = await Http.SendAsync(request);
                return response;
            }

        }
    }
}
