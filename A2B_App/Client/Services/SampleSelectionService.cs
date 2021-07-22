using A2B_App.Shared;
using A2B_App.Shared.Podio;
using A2B_App.Shared.Sox;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text;
using System.Threading.Tasks;

namespace A2B_App.Client.Services
{
    public class SampleSelectionService
    {

        private readonly ClientSettings settings;

        public SampleSelectionService(ClientSettings _settings)
        {
            settings = _settings;
        }

        //Create excel file and return filename
        public async Task<HttpResponseMessage> CreateExcelAsync(SampleSelection sampleSelection, HttpClient Http)
        {

            //ClientSettings settings = new ClientSettings();
            
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/SampleSelection/excel/create"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(sampleSelection));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                System.Diagnostics.Debug.WriteLine(response.Content.ReadAsStringAsync().Result);
                System.Diagnostics.Debug.WriteLine(response.StatusCode.ToString());
                return response;
            }
            
        }

        public async Task<HttpResponseMessage> SampleSelectionTestAsync(SampleSelection sampleSelection, HttpClient Http)
        {

            //ClientSettings settings = new ClientSettings();
           
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/SampleSelection/excel/test"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(sampleSelection));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                System.Diagnostics.Debug.WriteLine(response.Content.ReadAsStringAsync().Result);
                System.Diagnostics.Debug.WriteLine(response.StatusCode.ToString());
                return response;
            }
            
        }

        //Create podio test rounds and return list of id's
        public async Task<HttpResponseMessage> CreatePodioTestRoundAsync(SampleSelection sampleSelection, HttpClient Http)
        {

            //ClientSettings settings = new ClientSettings();
            WriteLog writelog = new WriteLog();
            writelog.Display("Aleo Testing2");
            writelog.Display(sampleSelection);
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/SampleSelection/podio/create/testround"))
            {
                request.Headers.TryAddWithoutValidation("accept", "*/*");

                request.Content = new StringContent(JsonConvert.SerializeObject(sampleSelection));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                System.Diagnostics.Debug.WriteLine(response.Content.ReadAsStringAsync().Result);
                System.Diagnostics.Debug.WriteLine(response.StatusCode.ToString());
                return response;
            }
            
        }

        //Create podio sample selection and return integer id
        public async Task<HttpResponseMessage> CreatePodioSampleSelectionAsync(SampleSelection sampleSelection, HttpClient Http)
        {

            //ClientSettings settings = new ClientSettings();
            //string server = settings.GetApiServer();
            
                using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/SampleSelection/podio/create/sampleselection"))
                {
                    request.Headers.TryAddWithoutValidation("accept", "*/*");

                    request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(sampleSelection));
                    request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                    var response = await Http.SendAsync(request);
                    return response;
                }
            
        }

        public async Task<HttpResponseMessage> GetRcmAsync(int itemId, HttpClient Http)
        {

            //ClientSettings settings = new ClientSettings();
            //string server = settings.GetApiServer();
            
                using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/Rcm/rcm/get/{itemId}"))
                {
                    request.Headers.TryAddWithoutValidation("accept", "*/*");
                    Rcm rcm = new Rcm();
                    request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(rcm));
                    request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                    var response = await Http.SendAsync(request);
                System.Diagnostics.Debug.WriteLine(response.Content.ReadAsStringAsync().Result);
                System.Diagnostics.Debug.WriteLine(response.StatusCode.ToString());
                    return response;
                }
            
        }

        public async Task<HttpResponseMessage> GetPopulation(HttpClient Http, string filename, int round)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("GET"), $"api/sampleSelection/download/population2/{filename}/{round}"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");
                var response = await Http.SendAsync(request);
                return response;
            }

        }

        public async Task<HttpResponseMessage> GetSampleSelectionData(HttpClient Http, string url)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("GET"), $"{url}"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");
                var response = await Http.SendAsync(request);
                return response;
            }

        }

        //Get tranasctional and materiality population data base on excel filename
        public async Task<List<Population>> SetPopulationTMAsync(string filename, int round, HttpClient Http)
        {

            //HttpClient Http = new HttpClient();
            ////ClientSettings settings = new ClientSettings();
            //string server = settings.GetApiServer();

            List<Population> listPopulation = null;
            //listPopulation = await Http.GetFromJsonAsync<List<Population>>($"api/sampleSelection/download/population2/{filename}/{round}");

            var response = await GetPopulation(Http, filename, round);
            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                string result = response.Content.ReadAsStringAsync().Result.ToString();
                listPopulation = JsonConvert.DeserializeObject<List<Population>>(result);
            }
            return listPopulation;
        }

        //Get tranasctional and materiality population data base on excel filename
        public async Task<List<List<string>>> SetPopulationTM2Async(string filename, int round, HttpClient Http)
        {
            List<List<string>> listPopulation = null;
            try
            {
                //listPopulation = await Http.GetFromJsonAsync<List<List<string>>>($"api/sampleSelection/download/population2/{filename}/{round}");
                var response = await GetPopulation(Http, filename, round);
                if (response.StatusCode == System.Net.HttpStatusCode.OK)
                {
                    string result = response.Content.ReadAsStringAsync().Result.ToString();
                    listPopulation = JsonConvert.DeserializeObject<List<List<string>>>(result);
                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                //throw;
            }

            return listPopulation;
        }

        public List<string> SetDropDownPopulationAsync(List<Population> listPop)
     {

            List<string> listDropdownPop = new List<string>
            {
                listPop[0].PurchaseOrder,
                listPop[0].SupplierSched,
                listPop[0].PoRev,
                listPop[0].PoLine,
                listPop[0].Requisition,
                listPop[0].RequisitionLine,
                listPop[0].EnteredBy,
                listPop[0].Status,
                listPop[0].Buyer,
                listPop[0].Contact,
                listPop[0].OrderDate,
                listPop[0].Supplier,
                listPop[0].ShipTo,
                listPop[0].SortName,
                listPop[0].Telephone,
                listPop[0].ItemNumber,
                listPop[0].ProdLine,
                listPop[0].ProdDescription,
                listPop[0].Site,
                listPop[0].Location,
                listPop[0].ItemRevision,
                listPop[0].SupplierItem,
                listPop[0].QuantityOrdered,
                listPop[0].UnitOfMeasure,
                listPop[0].UMConversion,
                listPop[0].QtyOrderedXPOCost,
                listPop[0].QuantityReceived,
                listPop[0].QtyOpen,
                listPop[0].QtyReturned,
                listPop[0].DueDate,
                listPop[0].OverDue,
                listPop[0].PerformanceDate,
                listPop[0].Currency,
                listPop[0].StandardCost,
                listPop[0].PurchasedCost,
                listPop[0].PurCostBC,
                listPop[0].OpenPoCost,
                listPop[0].PpvPerUnit,
                listPop[0].Type,
                listPop[0].StdMtlCostNow,
                listPop[0].WorkOrderId,
                listPop[0].Operation,
                listPop[0].PurchAcct,
                listPop[0].GlAccountDesc,
                listPop[0].CostCenter,
                listPop[0].GlDescription,
                listPop[0].Project,
                listPop[0].Description,
                listPop[0].Taxable,
                listPop[0].Comments
            };
            return listDropdownPop;

        }

        public async Task<List<string>> SetDropDownPopulationAsync2(List<List<string>> listPop)
        {
            await Task.Delay(1);
            List<string> listDropdownPop = new List<string>();
            foreach (var item in listPop[0])
            {
                listDropdownPop.Add(item);
            }

            return listDropdownPop;

        }

       public List<StringIndex> GetUniqueAsync(List<List<string>> listPop)
       {

            List<StringIndex> result = new List<StringIndex>();
            List<StringIndex> tempList = new List<StringIndex>();

            for (int i = 1; i < listPop.Count; i++)
            {
                StringIndex strindex = new StringIndex();
                strindex.Text = listPop[i][0];
                strindex.Index = i;
                tempList.Add(strindex);
                //Console.WriteLine($"Found Item: {listPop[i][0]}");
            }

            if (tempList != null)
            {
                //result = tempList.Select(x => x).Distinct().Count();
                foreach (var item in tempList)
                {
                    if (!result.Any(x => item.Text == x.Text))
                    {
                        result.Add(item);
                        //Console.WriteLine($"Added Unique Item: {item.Text} : Index({item.Index})");
                    }
                }
            }

            return result;

        }

    public List<string> GetListPopulationIndexOfAsync(List<List<string>> listPop, int index)
     {

            List<string> result = new List<string>();

            foreach (var item in listPop[index])
            {
                //Console.WriteLine($"Added {item}");
                result.Add(item);
            }

            return result;

        }

        //skip first row
        public List<List<string>> GetListPopulationContentUniqueAsync(List<List<string>> listPop)
        {
            
            List<List<string>> parent = new List<List<string>>();
            List<string> tempList = new List<string>();

            int row = 0;
            foreach (var item in listPop)
            {
                //start adding at row 1 since row 0 is the header in population file
                if (row > 0)
                {
                    List<string> child = new List<string>();
                    StringBuilder sb = new StringBuilder();
                    int count = 0;
                    bool isEmpty = true;
                    foreach (var innerItem in item)
                    {
                        count++;
                        if (count <= 1)
                        {
                            sb.Append(innerItem);
                        }
                        child.Add(innerItem);

                        if (innerItem != string.Empty && innerItem != null)
                        {
                            isEmpty = false;
                        }
                    }

                    //concatenate column A and column B and set it as a unique Id
                    //add unique id in 2nd to last child element
                    child.Add(sb.ToString());

                    //add row index at the last child element
                    child.Add(row.ToString());

                    //Console.WriteLine($"check if exists {child[child.Count - 2]}");
                    if (!tempList.Contains(child[child.Count - 2]) && !isEmpty)
                    {
                        parent.Add(child);
                        tempList.Add(child[child.Count - 2]);
                        //Console.WriteLine($"Added {string.Join(",",child)}");
                    }


                }
                row++;
            }

            //Display content in console
            row = 0;
            int column;
            foreach (var item in parent)
            {
                //Console.WriteLine($"Row : {row}");
                column = 0;
                foreach (var innerItem in item)
                {
                    //Console.WriteLine($"Column: {column} | Value : {innerItem}");
                    column++;
                }
                row++;
            }


            return parent;

        }

        public List<List<string>> GetListPopulationContentUnique2Async(List<List<string>> listPop)
        {
            
            List<List<string>> parent = new List<List<string>>();
            List<string> tempList = new List<string>();

            int row = 0;
            foreach (var item in listPop)
            {
                //start adding at index 1 since index 0 is the header in population file

                List<string> child = new List<string>();
                StringBuilder sb = new StringBuilder();
                int count = 0;
                bool isEmpty = true;
                foreach (var innerItem in item)
                {
                    count++;
                    if (count <= 1)
                    {
                        sb.Append(innerItem);
                    }
                    child.Add(innerItem);

                    if (innerItem != string.Empty && innerItem != null)
                    {
                        isEmpty = false;
                    }
                }

                //add unique id in 2nd to last child element
                child.Add(sb.ToString());

                //add row index at the last child element
                child.Add(row.ToString());

                //filter unique values on column 0, we add them if it doesn't exists
                //Console.WriteLine($"check if exists {child[child.Count - 2]}");
                if (!tempList.Contains(child[child.Count - 2]) && !isEmpty)
                {
                    parent.Add(child);
                    tempList.Add(child[child.Count - 2]);
                    //Console.WriteLine($"Added {string.Join(",", child)}");
                }



            }

            //Display content in console
            row = 0;
            int column;
            foreach (var item in parent)
            {
                //Console.WriteLine($"Row : {row}");
                column = 0;
                foreach (var innerItem in item)
                {
                    //Console.WriteLine($"Column: {column} | Value : {innerItem}");
                    column++;
                }
                row++;
            }


            return parent;

        }

       public List<List<string>> GetListPopulationContentAsync(List<List<string>> listPop)
       {

            List<List<string>> parent = new List<List<string>>();
            if (listPop != null)
            {
                listPop.RemoveAt(0);
            }
            return listPop;
        }

        public Task<List<List<string>>> RemoveItemListUnique(List<List<string>> listIndex, int index)
        {
            //var itemToRemove = listIndex.SingleOrDefault(r => r.Index == index);
            //if (itemToRemove != null)
            //    listIndex.Remove(itemToRemove);

            //return Task.FromResult(listIndex);

            var itemToRemove = listIndex.SingleOrDefault(r => r.Count - 1 == index);
            if (itemToRemove != null)
                listIndex.Remove(itemToRemove);

            return Task.FromResult(listIndex);
        }

        public async Task<List<SampleSize>> SetSampleSizeAsync(HttpClient Http)
        {
            List<SampleSize> listSample = null;
            SampleSize[] sampleSize;

            try
            {

                //store to Array and convert to List
                //sampleSize = await Http.GetFromJsonAsync<SampleSize[]>("api/sampleSelection/data/samplesize");
                //if(sampleSize != null)
                //{
                //    listSample = sampleSize.ToList();
                //}

                //Updated 5.13.2021 | Levin jay Tagapan
                var response = await GetSampleSelectionData(Http, $"api/sampleSelection/data/samplesize");
                if (response.StatusCode == System.Net.HttpStatusCode.OK)
                {
                    string result = response.Content.ReadAsStringAsync().Result.ToString();
                    listSample = JsonConvert.DeserializeObject<List<SampleSize>>(result);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }

            return listSample;

        }

        public async Task<List<SampleSize>> SetClientSampleSizeAsync(HttpClient Http, string clientName)
        {
            List<SampleSize> listSample = null;
            SampleSize[] sampleSize;

            try
            {

                //Updated 5.13.2021 | Levin jay Tagapan
                var response = await GetSampleSelectionData(Http, $"api/sampleSelection/data/clientsamplesize/{clientName}");
                if (response.StatusCode == System.Net.HttpStatusCode.OK)
                {
                    string result = response.Content.ReadAsStringAsync().Result.ToString();
                    listSample = JsonConvert.DeserializeObject<List<SampleSize>>(result);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }

            return listSample;

        }


        public async Task<HttpResponseMessage> SaveSampleSelectionAsync(SampleSelection sampleSelection, HttpClient Http)
        {

            //ClientSettings settings = new ClientSettings();
            //string server = settings.GetApiServer();
            
            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/SampleSelection/data/create"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(sampleSelection));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                System.Diagnostics.Debug.WriteLine(response.Content.ReadAsStringAsync().Result);
                System.Diagnostics.Debug.WriteLine(response.StatusCode.ToString());
                return response;
            }
            
        }


        //This function will Open excel file and get dropdown
        public async Task<List<DropDown>> SetDropDownAsync(HttpClient Http)
        {
            //HttpClient Http = new HttpClient();
            ////ClientSettings settings = new ClientSettings();
            //string server = settings.GetApiServer();
            List<DropDown> listDropdown = new List<DropDown>();

            listDropdown.Add(new DropDown { ExternalAuditor = string.Empty, Percent = 0 });
            //var resultValue = await Http.GetFromJsonAsync<List<DropDown>>($"api/sampleSelection/data/dropdown");
            //if (resultValue != null)
            //{
            //    listDropdown.AddRange(resultValue);
            //}

            //Updated 5.13.2021 | Levin jay Tagapan
            var response = await GetSampleSelectionData(Http, $"api/sampleSelection/data/dropdown");
            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                string result = response.Content.ReadAsStringAsync().Result.ToString();
                var listSampleTemp = JsonConvert.DeserializeObject<List<DropDown>>(result);
                if(listSampleTemp != null && listSampleTemp.Any())
                {
                    listDropdown.AddRange(listSampleTemp);
                }
            }

            return listDropdown;

        }

        //This function will set client name selection
        //Note: once we have database setup, will connect it to database
        public async Task<List<ClientSs>> SetClientAsync(HttpClient Http)
        {

            List<ClientSs> listClient = new List<ClientSs>();
            listClient.Add(new ClientSs { ClientName = string.Empty, ExternalAuditor = string.Empty });

            //var returnValue = await Http.GetFromJsonAsync<List<ClientSs>>($"api/sampleSelection/data/client");
            //if (returnValue != null)
            //{
            //    listClient.AddRange(returnValue.ToList());
            //}

            //Updated 5.13.2021 | Levin jay Tagapan
            var response = await GetSampleSelectionData(Http, $"api/sampleSelection/data/client");
            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                string result = response.Content.ReadAsStringAsync().Result.ToString();
                var tempClient = JsonConvert.DeserializeObject<List<ClientSs>>(result);
                if (tempClient != null && tempClient.Any())
                {
                    listClient.AddRange(tempClient);
                }
            }

            return listClient;

        }

        public async Task<List<ClientSs>> SetClient2Async(HttpClient Http, string clientName)
        {

            List<ClientSs> listClient = new List<ClientSs>();
            listClient.Add(new ClientSs { ClientName = string.Empty, ExternalAuditor = string.Empty });

            //var returnValue = await Http.GetFromJsonAsync<List<ClientSs>>($"api/sampleSelection/data/client");
            //if (returnValue != null)
            //{
            //    listClient.AddRange(returnValue.ToList());
            //}

            //Updated 5.13.2021 | Levin jay Tagapan
            var response = await GetSampleSelectionData(Http, $"api/sampleSelection/data/clientdetails/{clientName}");
            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                string result = response.Content.ReadAsStringAsync().Result.ToString();
                var tempClient = JsonConvert.DeserializeObject<List<ClientSs>>(result);
                if (tempClient != null && tempClient.Any())
                {
                    listClient.AddRange(tempClient);
                }
            }

            return listClient;

        }


        //This function will set frequency selection
        //Note: once we have database setup, will connect it to database
        public async Task<List<Frequency>> SetFrequencyAsync(HttpClient Http)
        {
            //HttpClient Http = new HttpClient();
            ////ClientSettings settings = new ClientSettings();
            //string server = settings.GetApiServer();
            List<Frequency> listFrequency = null ;
            //listFrequency = await Http.GetFromJsonAsync<List<Frequency>>($"api/sampleSelection/data/frequency");

            //Updated 5.13.2021 | Levin jay Tagapan
            var response = await GetSampleSelectionData(Http, $"api/sampleSelection/data/frequency");
            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                string result = response.Content.ReadAsStringAsync().Result.ToString();
                listFrequency = JsonConvert.DeserializeObject<List<Frequency>>(result);
            }
            return listFrequency;
        }

        //This function will set risk selection
        //Note: once we have database setup, will connect it to database
        public async Task<List<string>> SetRiskAsync(HttpClient Http)
        {

            //HttpClient Http = new HttpClient();
            ////ClientSettings settings = new ClientSettings();
            //string server = settings.GetApiServer();
            List<string> listRisk = null;
            //listRisk = await Http.GetFromJsonAsync<List<string>>($"api/sampleSelection/data/risk");


            //Updated 5.13.2021 | Levin jay Tagapan
            var response = await GetSampleSelectionData(Http, $"api/sampleSelection/data/risk");
            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                string result = response.Content.ReadAsStringAsync().Result.ToString();
                listRisk = JsonConvert.DeserializeObject<List<string>>(result);
            }

            return listRisk;
        }

        //This function will set Q4 R3 selection
        //Note: once we have database setup, will connect it to database
        public async Task<List<string>> SetQ4R3Async(HttpClient Http)
        {
            //HttpClient Http = new HttpClient();
            ////ClientSettings settings = new ClientSettings();
            //string server = settings.GetApiServer();
            List<string> listQ4Q3 = null;
            //listQ4Q3 = await Http.GetFromJsonAsync<List<string>>($"api/sampleSelection/data/q4r3");

            //Updated 5.13.2021 | Levin jay Tagapan
            var response = await GetSampleSelectionData(Http, $"api/sampleSelection/data/q4r3");
            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                string result = response.Content.ReadAsStringAsync().Result.ToString();
                listQ4Q3 = JsonConvert.DeserializeObject<List<string>>(result);
            }

            return listQ4Q3;
        }

        //Upload file to API
        public async Task<HttpResponseMessage> UploadFileAsync(MemoryStream ms, string fileName, HttpClient Http)
        {

            //ClientSettings settings = new ClientSettings();
            //string server = settings.GetApiServer();
           

                var content = new MultipartFormDataContent();
                content.Headers.ContentDisposition = new ContentDispositionHeaderValue("form-data");
                content.Add(new ByteArrayContent(ms.GetBuffer()), "file", fileName);

                var response = await Http.PostAsync($"api/fileupload/upload", content);
                System.Diagnostics.Debug.WriteLine(response.Content.ReadAsStringAsync().Result);
                System.Diagnostics.Debug.WriteLine(response.StatusCode.ToString());

                return response;

        }

        public async Task<List<SampleSelection>> GetListSampleSelectionAsync(HttpClient Http)
        {

            //HttpClient Http = new HttpClient();
            ////ClientSettings settings = new ClientSettings();
            //string server = settings.GetApiServer();
            List<SampleSelection> listSampleSelection = null;
            //listSampleSelection = await Http.GetFromJsonAsync<List<SampleSelection>>($"api/SampleSelection/data/listsampleselection");

            //Updated 5.13.2021 | Levin jay Tagapan
            var response = await GetSampleSelectionData(Http, $"api/SampleSelection/data/listsampleselection");
            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                string result = response.Content.ReadAsStringAsync().Result.ToString();
                listSampleSelection = JsonConvert.DeserializeObject<List<SampleSelection>>(result);
            }

            return listSampleSelection;

        }
        public async Task<HttpResponseMessage> CreateNewClient(ClientSs client, HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/ClientUIProcess/create"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(client));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                System.Diagnostics.Debug.WriteLine(response.Content.ReadAsStringAsync().Result);
                System.Diagnostics.Debug.WriteLine(response.StatusCode.ToString());
                return response;
            }

        }

        public async Task<HttpResponseMessage> Search(ClientSs client, HttpClient Http)
        {

            using (var request = new HttpRequestMessage(new HttpMethod("POST"), $"api/ClientUIProcess/search"))
            {
                request.Headers.TryAddWithoutValidation("accept", "text/plain");

                request.Content = new StringContent(Newtonsoft.Json.JsonConvert.SerializeObject(client));
                request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");

                var response = await Http.SendAsync(request);
                System.Diagnostics.Debug.WriteLine(response.Content.ReadAsStringAsync().Result);
                System.Diagnostics.Debug.WriteLine(response.StatusCode.ToString());
                return response;
            }

        }

        public async Task<List<string>> GetClients(HttpClient Http)
        {
            List<string> Client = null;
            //Client = await Http.GetFromJsonAsync<List<string>>($"api/ClientUIProcess/clients");

            //Updated 5.13.2021 | Levin jay Tagapan
            var response = await GetSampleSelectionData(Http, $"api/ClientUIProcess/clients");
            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                string result = response.Content.ReadAsStringAsync().Result.ToString();
                Client = JsonConvert.DeserializeObject<List<string>>(result);
            }
            return Client;
        }

        
        public async Task<List<string>> GetClientCode(HttpClient Http)
        {
            List<string> Code = null ;
            Http.DefaultRequestHeaders.TryAddWithoutValidation("Content-Type", "application/json; charset=utf-8");
            //Code = await Http.GetFromJsonAsync<List<string>>($"api/ClientUIProcess/clientscode");
            //Updated 5.13.2021 | Levin jay Tagapan
            var response = await GetSampleSelectionData(Http, $"api/ClientUIProcess/clientscode");
            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                string result = response.Content.ReadAsStringAsync().Result.ToString();
                Code = JsonConvert.DeserializeObject<List<string>>(result);
            }
            return Code;
        }
        
    }
}
