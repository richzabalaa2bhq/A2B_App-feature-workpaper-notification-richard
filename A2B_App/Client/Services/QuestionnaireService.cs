using A2B_App.Shared;
using A2B_App.Shared.Sox;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace A2B_App.Client.Services
{
    public class QuestionnaireService
    {

        private readonly ClientSettings settings;
        public QuestionnaireService(ClientSettings _settings)
        {
            settings = _settings;
        }

        public async Task<HttpResponseMessage> CreateQuestionnaireAsync(Questionnaire Questionnaire, HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Questionnaire/questionnaire/create"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(Questionnaire));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                System.Diagnostics.Debug.WriteLine(response.Content.ReadAsStringAsync().Result);
                System.Diagnostics.Debug.WriteLine(response.StatusCode.ToString());
                return response;
            }
            
        }

        public async Task<HttpResponseMessage> CreateExcelAsync(Questionnaire Questionnaire, HttpClient Http)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Questionnaire/excel/create"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(Questionnaire));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                System.Diagnostics.Debug.WriteLine(response.Content.ReadAsStringAsync().Result);
                System.Diagnostics.Debug.WriteLine(response.StatusCode.ToString());
                return response;
            }
            
        }

        public async Task<HttpResponseMessage> CreateExcelAsync2(QuestionnaireExcelData Questionnaire, HttpClient Http)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Questionnaire/excel/elc/create"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(Questionnaire));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                //Console.WriteLine(response.Content.ReadAsStringAsync().Result);
                //Console.WriteLine(response.StatusCode.ToString());
                return response;
            }
        }

        public async Task<HttpResponseMessage> CreateExcelAsync3(QuestionnaireRoundSet roundSet, HttpClient Http)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Questionnaire/excel/create"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(roundSet));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                //Console.WriteLine(response.Content.ReadAsStringAsync().Result);
                //Console.WriteLine(response.StatusCode.ToString());
                return response;
            }
        }

        public async Task<HttpResponseMessage> GetQuestionnaireAppId(QuestionnaireFieldParam QuestionnaireFieldParam, HttpClient Http)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Questionnaire/appId"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(QuestionnaireFieldParam));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                //Console.WriteLine(response.Content.ReadAsStringAsync().Result);
                //Console.WriteLine(response.StatusCode.ToString());
                return response;
            }  
        }

        public async Task<HttpResponseMessage> GetQuestionnaireField(QuestionnaireFieldParam QuestionnaireFieldParam, HttpClient Http)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Questionnaire/field"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(QuestionnaireFieldParam));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                //Console.WriteLine(response.Content.ReadAsStringAsync().Result);
                //Console.WriteLine(response.StatusCode.ToString());
                return response;
            }
        }

        public async Task<HttpResponseMessage> CreateQuestionnaireNotesAsync(List<NotesItem> ListNotesItem, HttpClient Http)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Questionnaire/podio/create/questionnaire/notes"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(ListNotesItem));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                //Console.WriteLine(response.Content.ReadAsStringAsync().Result);
                //Console.WriteLine(response.StatusCode.ToString());
                return response;
            }
        }

        public async Task<HttpResponseMessage> CreateQuestionnaireRoundsAsync(List<RoundItem> ListRoundItem, HttpClient Http)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Questionnaire/podio/create/questionnaire/testround"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(ListRoundItem));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                //Console.WriteLine(response.Content.ReadAsStringAsync().Result);
                //Console.WriteLine(response.StatusCode.ToString());
                return response;
            }
        }

        public async Task<HttpResponseMessage> CreateQuestionnaireTesterAsync(QuestionnaireTesterSet QuestionnaireTesterSet, HttpClient Http)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Questionnaire/createDraftWorkpaper"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(QuestionnaireTesterSet));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json"); 

                var response = await Http.SendAsync(request);
                //Console.WriteLine(response.Content.ReadAsStringAsync().Result);
                //Console.WriteLine(response.StatusCode.ToString());
                return response;
            }
        }

        public async Task<HttpResponseMessage> SaveQuestionnaireTesterAsync(List<QuestionnaireTesterSet> ListQuestionnaireTesterSet, HttpClient Http)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Questionnaire/saveDraftWorkpaper"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(ListQuestionnaireTesterSet));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                //Console.WriteLine(response.Content.ReadAsStringAsync().Result);
                //Console.WriteLine(response.StatusCode.ToString());
                return response;
            }
        }

        public async Task<HttpResponseMessage> GetTesterSetAsync(Rcm rcm, HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("GET"), $"api/Questionnaire/getAllTesterWorkpaper?rcmItemId={rcm.PodioItemId}"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                var response = await Http.SendAsync(request);
                //Console.WriteLine(response.Content.ReadAsStringAsync().Result);
                //Console.WriteLine(response.StatusCode.ToString());
                return response;
            }

        }


        public async Task<HttpResponseMessage> GetQuestionnaireTesterAsync(QuestionnaireDraftId draftId, HttpClient Http)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Questionnaire/getDraftWorkpaper"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(draftId));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                //Console.WriteLine(response.Content.ReadAsStringAsync().Result);
                //Console.WriteLine(response.StatusCode.ToString());
                return response;
            }
        }

        public async Task<HttpResponseMessage> GetQuestionnaireTesterFinalAsync(QuestionnaireDraftId draftId, HttpClient Http)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Questionnaire/getFinalWorkpaper"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(draftId));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                //Console.WriteLine(response.Content.ReadAsStringAsync().Result);
                //Console.WriteLine(response.StatusCode.ToString());
                return response;
            }
        }

        public async Task<HttpResponseMessage> CreateQuestionnaireIUCSystemGenAsync(List<IUCSystemGenAnswer> ListIUCSystemGen, HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Questionnaire/podio/create/questionnaire/systemgen"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(ListIUCSystemGen));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                //Console.WriteLine(response.Content.ReadAsStringAsync().Result);
                //Console.WriteLine(response.StatusCode.ToString());
                return response;
            }
           
        }

        public async Task<HttpResponseMessage> CreateQuestionnaireIUCNonSystemGenAsync(List<IUCNonSystemGenAnswer> ListIUCNonSystemGen, HttpClient Http)
        {


            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Questionnaire/podio/create/questionnaire/nonsystemgen"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(ListIUCNonSystemGen));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                //Console.WriteLine(response.Content.ReadAsStringAsync().Result);
                //Console.WriteLine(response.StatusCode.ToString());
                return response;
            }
          
        }

        public async Task<HttpResponseMessage> CreateQuestionnaireAsync2(List<QuestionnaireUserInput> ListUserInput, HttpClient Http)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Questionnaire/podio/create/questionnaire"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(ListUserInput));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                //Console.WriteLine(response.Content.ReadAsStringAsync().Result);
                //Console.WriteLine(response.StatusCode.ToString());
                return response;
            }
        }

        public async Task<HttpResponseMessage> CreateQuestionnaireAsync3(QuestionnaireRoundSet roundSet, HttpClient Http)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Questionnaire/podio/create/questionnaire2"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(roundSet));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                //Console.WriteLine(response.Content.ReadAsStringAsync().Result);
                //Console.WriteLine(response.StatusCode.ToString());
                return response;
            }
        }

        public async Task<HttpResponseMessage> GetIUCSystem(QuestionnaireFieldParam QuestionnaireFieldParam, HttpClient Http)
        {
            
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Questionnaire/iucsystem"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(QuestionnaireFieldParam));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                //Console.WriteLine(response.Content.ReadAsStringAsync().Result);
                //Console.WriteLine(response.StatusCode.ToString());
                return response;
            } 
        }

        public async Task<HttpResponseMessage> GetIUCNonSystem(QuestionnaireFieldParam QuestionnaireFieldParam, HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Questionnaire/iucnonsystem"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(QuestionnaireFieldParam));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                //Console.WriteLine(response.Content.ReadAsStringAsync().Result);
                //Console.WriteLine(response.StatusCode.ToString());
                return response;
            }
            
        }

        public async Task<HttpResponseMessage> GetClientWithAppId(HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Questionnaire/client/list"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                //request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                //Console.WriteLine(response.Content.ReadAsStringAsync().Result);
                //Console.WriteLine(response.StatusCode.ToString());
                return response;
            }
            
        }

        public async Task<HttpResponseMessage> GetClientListParam(string ClientName, HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Questionnaire/listparam?ClientName={ClientName}"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                //request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(QuestionnaireFieldParam));
                //request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                //Console.WriteLine(response.Content.ReadAsStringAsync().Result);
                //Console.WriteLine(response.StatusCode.ToString());
                return response;
            }
            
        }

        public async Task<HttpResponseMessage> LoadSoxClientParam(string fileName, HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Questionnaire/load/clientapp?fileName={fileName}"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                //request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(QuestionnaireFieldParam));
                //request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                //Console.WriteLine(response.Content.ReadAsStringAsync().Result);
                //Console.WriteLine(response.StatusCode.ToString());
                return response;
            }

        }

        public async Task<HttpResponseMessage> SaveQuestionnaireAsync(QuestionnaireRoundSet roundSet, HttpClient Http)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Questionnaire/save"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(roundSet));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                //Console.WriteLine(response.Content.ReadAsStringAsync().Result);
                //Console.WriteLine(response.StatusCode.ToString());
                return response;
            }
        }

        public async Task<HttpResponseMessage> GetAllRoundSetAsync(Rcm rcm, HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("GET"), $"api/Questionnaire/allroundset/{rcm.PodioItemId}"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                var response = await Http.SendAsync(request);
                //Console.WriteLine(response.Content.ReadAsStringAsync().Result);
                //Console.WriteLine(response.StatusCode.ToString());
                return response;
            }

        }

        public async Task<HttpResponseMessage> GetSpecificRoundSetAsync(string Id, HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("GET"), $"api/Questionnaire/roundset/{Id}"))
            {
                Http.Timeout = TimeSpan.FromMinutes(30);
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                var response = await Http.SendAsync(request);
                //Console.WriteLine(response.Content.ReadAsStringAsync().Result);
                //Console.WriteLine(response.StatusCode.ToString());
                return response;
            }

        } 
		
        public async Task<HttpResponseMessage> CreateSOXTrackerData(QuestionnaireRoundSet roundSet, HttpClient Http)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Questionnaire/sox/tracker/create"))
            {
                //Http.Timeout = TimeSpan.FromMinutes(10);
                request.Headers.TryAddWithoutValidation("accept", "text/plain");
				
                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(roundSet));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                //Console.WriteLine(response.Content.ReadAsStringAsync().Result);
                //Console.WriteLine(response.StatusCode.ToString());
                return response;
            }
        }

        //032321
        public async Task<HttpResponseMessage> UploadToSharefile(SharefileItem item, HttpClient Http)
        {


            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Sharefile/upload"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(item));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                //Debug.WriteLine($"Response Result: {response.Content.ReadAsStringAsync().Result}");
                //Debug.WriteLine($"Response Status Code: {response.StatusCode}");
                return response;
            }

        }
        //

        public async Task<HttpResponseMessage> AddQuestionaireAdditionalInputs(QuestionaireAddedInputs QuestionaireAddedInputss, HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Questionnaire/additionalInputs"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(QuestionaireAddedInputss));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                System.Diagnostics.Debug.WriteLine(response.Content.ReadAsStringAsync().Result);
                System.Diagnostics.Debug.WriteLine(response.StatusCode.ToString());
                return response;
            }

        }

        /*
        public async Task<HttpResponseMessage> GetAdditionalInputs(QuestionaireAddedInputs QuestionaireAddedInputss, HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("GET"), $"api/Questionnaire/additionalInputs/{QuestionaireAddedInputss.roundset}"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                var response = await Http.SendAsync(request);
                //Console.WriteLine(response.Content.ReadAsStringAsync().Result);
                //Console.WriteLine(response.StatusCode.ToString());
                return response;
            }

        }
        */


    }
}
