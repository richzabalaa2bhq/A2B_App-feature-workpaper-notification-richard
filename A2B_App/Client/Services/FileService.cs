using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using A2B_App.Shared.Sox;
using BlazorInputFile;
using Microsoft.AspNetCore.Components;

namespace A2B_App.Client.Services
{
    public class FileService
    {
        public async Task<HttpResponseMessage> UploadFileAsync(MemoryStream ms, string fileName, HttpClient Http)
        {

            var content = new MultipartFormDataContent();
            content.Headers.ContentDisposition = new ContentDispositionHeaderValue("form-data");
            content.Add(new ByteArrayContent(ms.GetBuffer()), "file", fileName);

            var response = await Http.PostAsync($"api/fileupload/image", content);
            System.Diagnostics.Debug.WriteLine(response.Content.ReadAsStringAsync().Result);
            System.Diagnostics.Debug.WriteLine(response.StatusCode.ToString());

            return response;
        }

        public async Task<HttpResponseMessage> UploadReportFileAsync(MemoryStream ms, string fileName, HttpClient Http)
        {

            var content = new MultipartFormDataContent();
            content.Headers.ContentDisposition = new ContentDispositionHeaderValue("form-data");
            content.Add(new ByteArrayContent(ms.GetBuffer()), "file", fileName);

            var response = await Http.PostAsync($"api/fileupload/keyreport", content);
            System.Diagnostics.Debug.WriteLine(response.Content.ReadAsStringAsync().Result);
            System.Diagnostics.Debug.WriteLine(response.StatusCode.ToString());

            return response;
        }

        public async Task<HttpResponseMessage> UploadFileForImport(MemoryStream ms, string fileName,int process, HttpClient Http)
        {
            string url = string.Empty;
            switch (process)
            {
                case 1:
                    url = "api/fileupload/uploadFileRcm";
                    break;
                case 2:
                    url = "api/fileupload/uploadFileSod";
                    break;
                case 3:
                    url = "api/fileupload/uploadFileKeyReport";
                    break;
                default:
                    break;
            }

            var content = new MultipartFormDataContent();
            content.Headers.ContentDisposition = new ContentDispositionHeaderValue("form-data");
            content.Add(new ByteArrayContent(ms.GetBuffer()), "file", fileName);
            //content.Add(new StringContent($"Process"), fileImport.Process.ToString());

            var response = await Http.PostAsync($"{url}", content);
            Debug.WriteLine($"Response Result: {response.Content.ReadAsStringAsync().Result}");
            Debug.WriteLine($"Response Status Code: {response.StatusCode}");

            return response;

            //using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/fileupload/uploadFile"))
            //{
            //    //request.Headers.TryAddWithoutValidation("accept", "*/*");
            //    request.Headers.Remove("Content-Type");
            //    request.Headers.TryAddWithoutValidation("Content-Type", "multipart/form-data");
            //    //ByteArrayContent byteArr = new ByteArrayContent(ms.GetBuffer()), "file", fileName);
            //    request.Content = new ByteArrayContent(ms.GetBuffer());
            //    request.Content = new StringContent($"Process={fileImport.Process}");


            //    var response = await Http.SendAsync(request);
            //    Debug.WriteLine($"Response Result: {response.Content.ReadAsStringAsync().Result}");
            //    Debug.WriteLine($"Response Status Code: {response.StatusCode}");
            //    return response;
            //}

        }

        public async Task<HttpResponseMessage> UploadFileForImportSod(MemoryStream ms, string fileName, FileImport fileImport, HttpClient Http)
        {
            string url = "api/fileupload/uploadFileSod";
            var content = new MultipartFormDataContent();

            content.Add(new ByteArrayContent(ms.GetBuffer()), "File", fileName);
            content.Add(new StringContent(fileImport.Process.ToString()), "Process");
            content.Add(new StringContent(fileImport.ClientName), "ClientName");

            var response = await Http.PostAsync($"{url}", content);
            Debug.WriteLine($"Response Result: {response.Content.ReadAsStringAsync().Result}");
            Debug.WriteLine($"Response Status Code: {response.StatusCode}");

            return response;
        }

        public async Task<HttpResponseMessage> UploadFileForImportSodSoxRox(
            MemoryStream ms1,
            MemoryStream ms2,
            MemoryStream ms3,
            MemoryStream ms4,
            string fileNameRole,
            string fileNamePerm,
            string fileNameConflict,
            string fileNameDesc,
            SodSoxRoxImport fileImport,
            HttpClient Http)
        {
            string url = "api/fileupload/uploadFileSodSoxRox";
            var content = new MultipartFormDataContent();
            
            
            content.Add(new ByteArrayContent(ms1.GetBuffer()), "FileRoleUser", fileNameRole);
            content.Add(new ByteArrayContent(ms2.GetBuffer()), "FileRolePerm", fileNamePerm);
            content.Add(new ByteArrayContent(ms3.GetBuffer()), "FileConflictPerm", fileNameConflict);
            content.Add(new ByteArrayContent(ms4.GetBuffer()), "FileDescToPerm", fileNameDesc);
            content.Add(new StringContent(fileImport.ClientName), "ClientName");
            content.Add(new StringContent(fileImport.RequestedBy), "RequestedBy");

            var response = await Http.PostAsync($"{url}", content);
            Debug.WriteLine($"Response Result: {response.Content.ReadAsStringAsync().Result}");
            Debug.WriteLine($"Response Status Code: {response.StatusCode}");

            return response;
        }
        
        public async Task<HttpResponseMessage> StartImportRcm(ImportFields importFields, HttpClient Http)
        {

            using var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/fileupload/importRcm");
            request.Headers.TryAddWithoutValidation("accept", "text/plain");

            request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(importFields));
            request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

            var response = await Http.SendAsync(request);
            Debug.WriteLine($"Response Result: {response.Content.ReadAsStringAsync().Result}");
            Debug.WriteLine($"Response Status Code: {response.StatusCode}");
            return response;

        }

    }
}
