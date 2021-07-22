using A2B_App.Shared;
using A2B_App.Shared.Podio;
using A2B_App.Shared.Sox;
using System;
using System.Diagnostics;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace A2B_App.Client.Services
{
    public class PodioSyncService
    {

        private readonly ClientSettings settings;
        public PodioSyncService(ClientSettings _settings)
        {
            settings = _settings;
        }

        public async Task<HttpResponseMessage> SyncPodioRcmAsync(SyncDateRange syncDateRange, HttpClient Http)
        {
            //await Http.GetFromJsonAsync<WeatherForecast[]>("WeatherForecast");

            //ClientSettings settings = new ClientSettings();
            string server = settings.GetApiServer();

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/PodioSync/podio/sync/rcm"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                var test = Newtonsoft.Json.JsonConvert.SerializeObject(syncDateRange);
                Debug.WriteLine($"{test}");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(syncDateRange));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                //Debug.WriteLine(response.Content.ReadAsStringAsync().Result);
                //Debug.WriteLine(response.StatusCode.ToString());
                return response;
            }


        }

        public async Task<HttpResponseMessage> SyncPodioQuestionnaireFieldAsync(QuestionnaireFieldParam questionnaireFieldParam, HttpClient Http, string title)
        {
            string URI = string.Empty;
            switch (title)
            {
                case "Questionnaire Field":
                    URI = $"api/PodioSync/podio/sync/questionnaire/fields";
                    break;
                case "IUC System Gen Field":
                    URI = $"api/PodioSync/iuc/systemgen/fields";
                    break;
                case "IUC System Non Gen Field":
                    URI = $"api/PodioSync/iuc/nonsystemgen/fields";
                    break;

            }

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), URI))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                var test = Newtonsoft.Json.JsonConvert.SerializeObject(questionnaireFieldParam);
                Debug.WriteLine($"{test}");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(questionnaireFieldParam));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                //Debug.WriteLine(response.Content.ReadAsStringAsync().Result);
                //Debug.WriteLine(response.StatusCode.ToString());
                return response;
            }
        }

        public async Task<HttpResponseMessage> SyncPodioSampleSelectionClientAsync(SyncDateRange syncDateRange, HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/PodioSync/podio/sync/sampleselection/client"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                var test = Newtonsoft.Json.JsonConvert.SerializeObject(syncDateRange);
                Debug.WriteLine($"{test}");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(syncDateRange));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                Debug.WriteLine(response.Content.ReadAsStringAsync().Result);
                Debug.WriteLine(response.StatusCode.ToString());
                return response;
            }

        }

        public async Task<HttpResponseMessage> SyncPodioSampleSelectionMatrixAsync(SyncDateRange syncDateRange, HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/PodioSync/podio/sync/sampleselection/matrix"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                var test = Newtonsoft.Json.JsonConvert.SerializeObject(syncDateRange);
                Debug.WriteLine($"{test}");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(syncDateRange));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                //Debug.WriteLine(response.Content.ReadAsStringAsync().Result);
                //Debug.WriteLine(response.StatusCode.ToString());
                return response;
            }

        }

        public async Task<(HttpResponseMessage, string)> SyncByDateAsync(SyncDateRange syncDateRange, HttpClient Http, string title)
        {
            string url = string.Empty;
            switch (title)
            {
                case "RCM":
                    url = $"api/PodioSync/podio/sync/rcm";
                    break;
                case "Sample Selection Client":
                    url = $"api/PodioSync/podio/sync/sampleselection/client";
                    break;
                case "Sample Selection Matrix":
                    url = $"api/PodioSync/podio/sync/sampleselection/matrix";
                    break;
                case "Client Reference":
                    url = $"api/PodioSync/podio/sync/time/clientreference";
                    break;
                case "Project Reference":
                    url = $"api/PodioSync/podio/sync/time/projectreference";
                    break;
                case "Task Reference":
                    url = $"api/PodioSync/podio/sync/time/taskreference";
                    break;
                case "Time Code":
                    url = $"api/PodioSync/podio/sync/time/timecode";
                    break;
                case "Sox Tracker":
                    url = $"api/PodioSync/soxtracker";
                    break;
                case "Sox Tracker Questionnaire":
                    url = $"api/PodioSync/soxtracker/questionnaire";
                    break;
                case "RCM Questionnaire":
                    url = $"api/PodioSync/rcm/questionnaire";
                    break;
                case "RCM Process":
                    url = $"api/PodioSync/podio/sync/rcmprocess";
                    break;
                case "RCM Subprocess":
                    url = $"api/PodioSync/podio/sync/rcmsubprocess";
                    break;
                case "RCM Control Id":
                    url = $"api/PodioSync/podio/sync/rcmcontrolid";
                    break;

                case "Key Report Leadsheet Workpaper":
                    url = $"api/PodioSync/keyreport/questionnaire/leadsheet/fields";
                    break;
                case "Key Report Consolidated Original Format Workpaper":
                    url = $"api/PodioSync/keyreport/questionnaire/consoleformat/fields";
                    break;
                case "Key Report IUC KR Questionnaire Workpaper":
                    url = $"api/PodioSync/keyreport/questionnaire/alliuc/fields";
                    break;
                case "Key Report IUC Report List Workpaper":
                    url = $"api/PodioSync/keyreport/questionnaire/statustracker/fields";
                    break;
                case "Key Report Exception Workpaper":
                    url = $"api/PodioSync/keyreport/questionnaire/exception/fields";
                    break;

                case "Key Report Consolidated Original Format Items":
                    url = $"api/PodioSync/keyreport/consoleformat";
                    break;
                case "Key Report IUC KR Questionnaire Items":
                    url = $"api/PodioSync/keyreport/alliuc";
                    break;
                case "Key Report IUC Report List Items":
                    url = $"api/PodioSync/keyreport/statustracker";
                    break;
                case "Key Report Exception Items":
                    url = $"api/PodioSync/keyreport/exception";
                    break;

                case "Key Report Control Activity":
                    url = $"api/PodioSync/keyreport/KeyReportControlActivity";
                    break;
                case "Key Report Key Control":
                    url = $"api/PodioSync/keyreport/KeyReportKeyControl";
                    break;
                case "Key Report Name":
                    url = $"api/PodioSync/keyreport/KeyReportName";
                    break;
                case "Key Report System Source":
                    url = $"api/PodioSync/keyreport/KeyReportSystemSource";
                    break;
                case "Key Report - Non Key Report":
                    url = $"api/PodioSync/keyreport/KeyReportNonKeyReport";
                    break;
                case "Key Report Customized":
                    url = $"api/PodioSync/keyreport/KeyReportReportCustomized";
                    break;
                case "Key Report IUC Type":
                    url = $"api/PodioSync/keyreport/KeyReportIUCType";
                    break;
                case "Key Report Controls Relying IUC":
                    url = $"api/PodioSync/keyreport/KeyReportControlsRelyingIUC";
                    break;
                case "Key Report Preparer":
                    url = $"api/PodioSync/keyreport/KeyReportPreparer";
                    break;
                case "Key Report Unique Key Report":
                    url = $"api/PodioSync/keyreport/KeyReportUniqueKeyReport";
                    break;
                case "Key Report Notes":
                    url = $"api/PodioSync/keyreport/KeyReportNotes";
                    break;
                case "Key Report Number":
                    url = $"api/PodioSync/keyreport/KeyReportNumber";
                    break;
                case "Key Report Tester":
                    url = $"api/PodioSync/keyreport/KeyReportTester";
                    break;
                case "Key Report Reviewer":
                    url = $"api/PodioSync/keyreport/KeyReportReviewer";
                    break;
                case "Parameter Library":
                    url = $"api/PodioSync/keyreport/Parameter";
                    break;
                case "Report Library":
                    url = $"api/PodioSync/keyreport/Report";
                    break;
                case "Completeness Library":
                    url = $"api/PodioSync/keyreport/Completeness";
                    break;
                case "Accuracy Library":
                    url = $"api/PodioSync/keyreport/Accuracy";
                    break;
                case "Method Library":
                    url = $"api/PodioSync/keyreport/CAMethod";
                    break;

            }

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), url))
            {             
                request.Headers.TryAddWithoutValidation("accept", "text/plain");
                var test = Newtonsoft.Json.JsonConvert.SerializeObject(syncDateRange);
                //Debug.WriteLine($"{test}");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(syncDateRange));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                return (response, title);
            }


        }
    

    
    
    
    }
}
