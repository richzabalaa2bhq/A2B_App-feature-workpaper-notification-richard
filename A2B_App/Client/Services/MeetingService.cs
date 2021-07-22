using A2B_App.Shared;
using A2B_App.Shared.Meeting;
using A2B_App.Shared.Skype;
using A2B_App.Shared.Sox;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Threading.Tasks;

namespace A2B_App.Client.Services
{
    public class MeetingService
    {

        private readonly ClientSettings settings;
        public MeetingService(ClientSettings clientSettings)
        {
            settings = clientSettings;
        }

        public async Task<HttpResponseMessage> GetAllMeeting(HttpClient Http)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("GET"), $"api/Meeting/dailyMeeting"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                //request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(filter));
                //request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                //Debug.WriteLine($"Response Result GetAllMeeting: {response.Content.ReadAsStringAsync().Result}");
                //Debug.WriteLine($"Response Status Code: {response.StatusCode}");
                return response;
            }
        }
        public async Task<HttpResponseMessage> FetchAttendees(int meetingID, HttpClient Http)
        {
           
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Meeting/fetch-attendees"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(meetingID));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                Debug.WriteLine($"Response Result: {response.Content.ReadAsStringAsync().Result}");
                Debug.WriteLine($"Response Status Code: {response.StatusCode}");
                return response;
            }
        }

        public async Task<HttpResponseMessage> GetAllMeetingBizDev(HttpClient Http)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("GET"), $"api/Meeting/dailyMeetingBizDev"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                //request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(filter));
                //request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                //Debug.WriteLine($"Response Result GetAllMeetingBizDev: {response.Content.ReadAsStringAsync().Result}");
                //Debug.WriteLine($"Response Status Code: {response.StatusCode}");
                return response;
            }
        }

        public async Task<HttpResponseMessage> GetAllRecordings(HttpClient Http)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("GET"), $"api/Meeting/recordings"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");
                //request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(filter));
                //request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");
                var response = await Http.SendAsync(request);
                Debug.WriteLine($"Response Result GetSoxTrackerClient: {response.Content.ReadAsStringAsync().Result}");
                Debug.WriteLine($"Response Status Code: {response.StatusCode}");
                return response;
            }
        }
        public async Task<HttpResponseMessage> GetWeeklyRecordings(HttpClient Http)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("GET"), $"api/Meeting/weeklyrecordings"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                //request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(filter));
                //request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                Debug.WriteLine($"Response Result GetSoxTrackerClient: {response.Content.ReadAsStringAsync().Result}");
                Debug.WriteLine($"Response Status Code: {response.StatusCode}");
                return response;

            }
        }
        public async Task<HttpResponseMessage> GetMonthlyRecording(HttpClient Http)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("GET"), $"api/Meeting/monthlyrecordings"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                //request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(filter));
                //request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                Debug.WriteLine($"Response Result GetSoxTrackerClient: {response.Content.ReadAsStringAsync().Result}");
                Debug.WriteLine($"Response Status Code: {response.StatusCode}");
                return response;
            }
        }
        public async Task<HttpResponseMessage> GetDownload(HttpClient Http, recording_gtm item)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Meeting/downloadlink"))
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
        public async Task<HttpResponseMessage> GetWeekly(HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("GET"), $"api/Meeting/weeklyMeeting"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                //request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(filter));
                //request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                Debug.WriteLine($"Response Result GetSoxTrackerClient: {response.Content.ReadAsStringAsync().Result}");
                Debug.WriteLine($"Response Status Code: {response.StatusCode}");
                return response;
            }

        }

        public async Task<HttpResponseMessage> GetRecordingsByDateRange(HttpClient Http, DateRangeCustom range)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Meeting/daterangerecordings"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(range));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                Debug.WriteLine($"Response Result: {response.Content.ReadAsStringAsync().Result}");
                Debug.WriteLine($"Response Status Code: {response.StatusCode}");
                return response;

            }
        }


        public async Task<HttpResponseMessage> GetMonthly(HttpClient Http)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("GET"), $"api/Meeting/monthlyMeeting"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                //request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(filter));
                //request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                Debug.WriteLine($"Response Result GetSoxTrackerClient: {response.Content.ReadAsStringAsync().Result}");
                Debug.WriteLine($"Response Status Code: {response.StatusCode}");
                return response;
            }
        }
        public async Task<HttpResponseMessage> GetByDateRange(HttpClient Http, DateRangeCustom range)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Meeting/dateRangeMeeting"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(range));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                Debug.WriteLine($"Response Result: {response.Content.ReadAsStringAsync().Result}");
                Debug.WriteLine($"Response Status Code: {response.StatusCode}");
                return response;
                
            }
        }



        public async Task<HttpResponseMessage> GetAllGC(HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("GET"), $"api/Meeting/allMeetingNotif"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                //request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(filter));
                //request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                Debug.WriteLine($"Response Result GetSoxTrackerClient: {response.Content.ReadAsStringAsync().Result}");
                Debug.WriteLine($"Response Status Code: {response.StatusCode}");
                return response;
            }

        }

        public async Task<HttpResponseMessage> SaveSkypeGC(SkypeAddress skypeAddress, HttpClient Http)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Meeting/saveMeetingNotif"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(skypeAddress));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                //Debug.WriteLine($"Response Result: {response.Content.ReadAsStringAsync().Result}");
                //Debug.WriteLine($"Response Status Code: {response.StatusCode}");
                return response;
            }

        }

        public async Task<HttpResponseMessage> DeleteSkypeGC(SkypeAddress skypeAddress, HttpClient Http)
        {
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Meeting/removeMeetingNotif"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(skypeAddress));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                //Debug.WriteLine($"Response Result: {response.Content.ReadAsStringAsync().Result}");
                //Debug.WriteLine($"Response Status Code: {response.StatusCode}");
                return response;
            }

        }

        

    }
}
