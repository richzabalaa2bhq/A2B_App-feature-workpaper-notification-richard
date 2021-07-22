using A2B_App.Server.Data;
using A2B_App.Server.Services;
using A2B_App.Shared.Podio;
//using A2B_App.Server.Log;
using A2B_App.Shared.Sox;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Newtonsoft.Json;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using PodioAPI;
using PodioAPI.Models;
using PodioAPI.Utils.ItemFields;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace A2B_App.Server.Controllers
{
    [Authorize]
    [Route("api/[controller]")]
    [ApiController]
    
    public class KeyReportController : ControllerBase
    {
        private readonly IConfiguration _config;
        private readonly ILogger<KeyReportController> _logger;
        private readonly IWebHostEnvironment _environment;
        private readonly SoxContext _soxContext;

        public KeyReportController(IConfiguration config,
            ILogger<KeyReportController> logger,
            IWebHostEnvironment environment,
            SoxContext soxContext)
        {
            _config = config;
            _logger = logger;
            _environment = environment;
            _soxContext = soxContext;
        }



        //[AllowAnonymous]
        [HttpGet("download/{filename}")]
        public async Task<IActionResult> GetKeyReportDownloadAsync(string filename)
        {

            try
            {
                string startupPath = Directory.GetCurrentDirectory();
                //string path = Path.Combine(startupPath, "include", "keyreport", filename); 031421
                //string path = Path.Combine(startupPath, "include","upload","keyreport","leadsheet", filename);
                string path = Path.Combine(startupPath, "include", "upload", "keyreport", "kr tracker", filename);
                var memory = new MemoryStream();
                using (var stream = new FileStream(path, FileMode.Open))
                {
                    await stream.CopyToAsync(memory);
                }
                memory.Position = 0;
                var ext = Path.GetExtension(path).ToLowerInvariant();

                return File(memory, GetMimeTypes()[ext], Path.GetFileName(path));
            }
            catch (Exception ex)
            {
                FileLog.Write($"Error GetKeyReportDownloadAsync {ex}", "ErrorGetKeyReportDownloadAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetKeyReportDownloadAsync");
                if (_environment.IsDevelopment())
                {
                    return BadRequest("Oops something when wrong!." + ex.ToString());
                }
                else
                {
                    return BadRequest("Oops something when wrong!.");
                }

            }
        }


        [HttpGet("keyReport/clients/")]
        public IActionResult GetListClientAsync()
       {
            List<string> _clients = new List<string>();
            try
            {
                using (var context = _soxContext.Database.BeginTransaction())
                {
                    var listClient = _soxContext.ClientSs .OrderBy(x => x.ClientName).Select(x => x.ClientName);

                    if (listClient != null)
                    {
                        _clients = listClient.ToList();
                    }
                }
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetListClientAsync {ex}", "ErrorGetListClientAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetListClientAsync");
            }

            if (_clients != null)
            {
                return Ok(_clients.ToArray());
            }
            else
            {
                return NoContent();
            }

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="filter"></param>
        /// <returns>List<string></returns>
        //[AllowAnonymous]
        [HttpPost("clientbyyear")]
        public IActionResult GetListClientAsync([FromBody] RcmQuestionnaireFilter filter)
        {
            List<KeyReportFileProperties> _listClient = new List<KeyReportFileProperties>();
            try
            {//RCM SF Temp 012621
                //032221
                //string rcmSharefile = _config.GetSection("SharefileApi").GetSection("SoxRcmFolder").GetSection("Link").Value;
                string keyLeadsheetSharefile = _config.GetSection("SharefileApi").GetSection("SoxKeyReportLeadSheetFolder").GetSection("Link").Value;
                using (var context = _soxContext.Database.BeginTransaction())
                {
                    var listKeyReport = _soxContext.KeyReport
                        .Where(x => 
                            
                          //  x.FY.Equals(filter.FY) &&
                            x.ClientName != string.Empty &&
                            x.ClientName != null
                        )
                        .Select(x => x.ClientName);

                    if (listKeyReport != null)
                    {
                        var listClient = listKeyReport.Distinct().OrderBy(x => x).ToList();
                        foreach (var item in listClient)
                        {
                            _listClient.Add(new KeyReportFileProperties //012621
                            {
                                ClientName = item,
                                LoadingStatus = string.Empty,
                                //SharefileLink = rcmSharefile,
                                SharefileLink = keyLeadsheetSharefile,
                                FileName = string.Empty
                            });
                        }
                        //_listClient = listRcm.Distinct().OrderBy(x => x).ToList();
                    }
                }
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetListClientAsync {ex}", "ErrorGetListClientAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetListClientAsync");
            }

            if (_listClient != null)
            {
                return Ok(_listClient.ToArray());
            }
            else
            {
                return NoContent();
            }

        }


        [AllowAnonymous]
        [HttpPost("leadsheet/generate")]
        public IActionResult GenerateKeyReportTemplateFile([FromBody] KeyReportScreenshot filter)
        {
            //List<string> excelFilename = new List<string>();
            string excelFilename = string.Empty;

            try
            {

                List<int> listItemIdAll = new List<int>();
                List<int> listItemIdUnique = new List<int>();
                List<int> listItemIdUniqueTestStatus = new List<int>();
                List<int> listItemIdUniqueException = new List<int>();
                ExcelService xlsService = new ExcelService();
                string appId = _config.GetSection("KeyReportApp").GetSection("AllIUCId").Value;
                string appIdTestStat = _config.GetSection("KeyReportApp").GetSection("TestStatusTrackerId").Value;

                var inputItemId = _soxContext.KeyReportUserInput.Where(x =>
                        x.TagFY.ToLower().Equals(filter.Filter.FY.ToLower())
                        && x.TagClientName.ToLower().Equals(filter.Filter.ClientName.ToLower())
                        && x.TagReportName.ToLower().Equals(filter.Filter.KeyReportName.ToLower())
                        && x.AppId.Equals(appId)
                        && x.TagStatus.ToLower() != "inactive"
                    )
                    .AsNoTracking()
                    .Select(x => x.ItemId)
                    .Distinct()
                    .ToList();
                foreach (var itemId in inputItemId)
                {
                    if (!listItemIdAll.Contains(itemId) && itemId != 0)
                    {
                        listItemIdAll.Add(itemId);
                    }
                }

                //Get all screenshot base from Client, ReportName, FY and ControlId
                List<KeyReportScreenshotUpload> listScreenshot = new List<KeyReportScreenshotUpload>();
                var allScreenshot = _soxContext.KeyReportScreenshotUpload
                    .Where(sc =>
                        sc.Client.Equals(filter.Filter.ClientName)
                        && sc.ReportName.Equals(filter.Filter.KeyReportName)
                        && sc.Fy.Equals(filter.Filter.FY))
                    .AsNoTracking()
                    .ToList();
                if(allScreenshot != null && allScreenshot.Any())
                {
                    listScreenshot = allScreenshot;
                   
                }

                //Get all screenshot base from Client, ReportName, FY and ControlId
                List<KeyReportFile> listReportFile = new List<KeyReportFile>();
                var allReportFile = _soxContext.KeyReportFile
                    .Where(rf =>
                        rf.Client.Equals(filter.Filter.ClientName)
                        && rf.ReportName.Equals(filter.Filter.KeyReportName)
                        && rf.Fy.Equals(filter.Filter.FY))
                    .AsNoTracking()
                    .ToList();
                if (allReportFile != null && allReportFile.Any())
                {
                    filter.ReportFilename = allReportFile[0].NewFilename;
                }


                if (listItemIdAll != null && listItemIdAll.Count > 0)
                {
                    //Get unique itemid
                    foreach (var item in listItemIdAll)
                    {
                        var inputItemId4 = _soxContext.KeyReportUserInput.Where(x =>
                                        x.ItemId.Equals(item)
                                        && x.StrQuestion.ToLower().Contains("unique key report")
                                        && x.StrAnswer.ToLower().Equals("yes")
                                        && x.TagStatus.ToLower() != "inactive"
                                    )
                                    .AsNoTracking()
                                    .Select(x => new { x.ItemId , x.TagControlId})
                                    .Distinct()
                                    .ToList();

                        foreach (var itemUnique in inputItemId4)
                        {
                            if (!listItemIdUnique.Contains(itemUnique.ItemId) && itemUnique.ItemId != 0)
                            {
                                listItemIdUnique.Add(itemUnique.ItemId);
                                filter.Filter.ControlId = itemUnique.TagControlId;
                            }
                        }
                    }

                    //Get Testing Status
                    var testingStatusItemId = _soxContext.KeyReportUserInput.Where(x =>
                                        x.TagClientName.Equals(filter.Filter.ClientName)
                                        && x.TagFY.Equals(filter.Filter.FY)
                                        && x.TagReportName.Equals(filter.Filter.KeyReportName)
                                        && x.TagControlId.Equals(filter.Filter.ControlId) 
                                        && x.TagStatus.ToLower() != "inactive"
                                        && x.AppId.Equals(appIdTestStat)
                                        && x.StrQuestion.ToLower().Contains("unique key report")
                                        && x.StrAnswer.ToLower().Equals("yes")
                                    )
                                    .AsNoTracking()
                                    .Select(x =>  x.ItemId)
                                    .Distinct()
                                    .FirstOrDefault();

                    List<KeyReportUserInput> testingStatus = new List<KeyReportUserInput>();
                    if (testingStatusItemId != 0)
                    {
                        
                        testingStatus = _soxContext.KeyReportUserInput.Where(x =>
                                                            x.ItemId.Equals(testingStatusItemId)
                                                        )
                                                        .AsNoTracking()
                                                        .Distinct()
                                                        .ToList();
                               
                    }


                    //Processing excel output
                    if (listItemIdUnique != null && listItemIdUnique.Count > 0)
                    {
                        //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        using (ExcelPackage xls = new ExcelPackage())
                        {
                            string screenshotDirectory = Directory.GetCurrentDirectory();
                            //string screenshotPath = Path.Combine(screenshotDirectory, "include", "sharefile", "download");
                            string screenshotPath = Path.Combine(screenshotDirectory, "include", "upload", "image");
                            ExcelPackage xlsReport = null;

                            filter.ListScreenshotName = new List<string>();
                            if(listScreenshot != null && listScreenshot.Any())
                            {
                                foreach (var item in listScreenshot)
                                {
                                    filter.ListScreenshotName.Add(item.NewFilename);
                                }
                            }


                            //check if report filename is not null, then we open the file
                            if (filter.ReportFilename != null && filter.ReportFilename != string.Empty)
                            {
                                //string fileReportPath = Path.Combine(screenshotPath, $"{filter.Filter.ClientName}", $"{filter.Filter.KeyReportName}", filter.ReportFilename);
                                string fileReportPath = Path.Combine(screenshotPath, filter.ReportFilename);
                                //Check if file not exists
                                if (System.IO.File.Exists(fileReportPath))
                                {
                                    FileInfo reportFileInfo = new FileInfo(fileReportPath);
                                    xlsReport = new ExcelPackage(reportFileInfo);
                                }
                                
                            }


                            //set sheet name
                            var ws = xls.Workbook.Worksheets.Add("LeadSheet");
                            var ws1 = xls.Workbook.Worksheets.Add($"1. {filter.Filter.KeyReportName} - Screenshot");
                            //var ws2 = xls.Workbook.Worksheets.Add($"2. {filter.Filter.KeyReportName}");

                            bool isImageAddedTab2 = false;
                            bool isImageAddedTab4 = false;
                            bool isImageAddedTab5 = false;

                            foreach (var itemId in listItemIdUnique)
                            {
                                //Get All IUC
                                var keyReportItem = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(itemId)).ToList();

                                #region Tab 1
                                ws.View.ZoomScale = 90;
                                //set column width
                                ws.Column(1).Width = 9;     //step
                                ws.Column(2).Width = 88;    //procedure    
                                ws.Column(3).Width = 77;    //results
                                ws.Column(4).Width = 50;    //notes

                                //disable grid
                                ws.View.ShowGridLines = false;

                                //set table header
                                //Tab1 Leadsheet
                                ws.Cells[1, 1].Value = ""; //Step #
                                ws.Cells[1, 2].Value = "Procedures";
                                ws.Cells[1, 3].Value = "Results";

                                ws.Cells[2, 2].Value = "Key Report Information";

                                ws.Cells[3, 1].Value = "1";
                                ws.Cells[4, 1].Value = "2";
                                ws.Cells[5, 1].Value = "3";
                                ws.Cells[6, 1].Value = "4";
                                ws.Cells[7, 1].Value = "5";
                                ws.Cells[8, 1].Value = "6";
                                ws.Cells[9, 1].Value = "7";
                                ws.Cells[10, 1].Value = "8";
                                ws.Cells[11, 1].Value = "9";
                                ws.Cells[12, 1].Value = "10";
                                ws.Cells[13, 1].Value = "11";
                                ws.Cells[14, 1].Value = "12";
                                ws.Cells[15, 1].Value = "13";
                                ws.Cells[16, 1].Value = "14";
                                ws.Cells[17, 1].Value = "15";
                                ws.Cells["A17:A20"].Merge = true;
                                ws.Cells[21, 1].Value = "16";
                                ws.Cells[22, 1].Value = "17";
                                ws.Cells[23, 1].Value = "18";
                                ws.Cells[24, 1].Value = "19";
                                ws.Cells[25, 1].Value = "20";
                                ws.Cells[26, 1].Value = "";
                                ws.Cells[27, 1].Value = "21";
                                ws.Cells[28, 1].Value = "22";
                                ws.Cells[29, 1].Value = "23";
                                ws.Cells[30, 1].Value = "24";
                                ws.Cells[31, 1].Value = "25";
                                ws.Cells[32, 1].Value = "26";
                                ws.Cells[33, 1].Value = "";
                                ws.Cells[34, 1].Value = "27";
                                ws.Cells[35, 1].Value = "28";
                                ws.Cells[36, 1].Value = "29";
                               // ws.Range(ws.Cells[33, 1], ws.Cells[34, 1]).Merge();
                               // ws.Cells[34, 1].Merge;
                               
                                String accuracy = "Accuracy";
                                String accuracyB = "<b>" + accuracy + "</b>";
                                String netsuite = "NetSuite";
                                String netsuiteU = "<U>" + netsuite + "</U>";

                                ws.Cells[3, 2].Value = "What is the key report name?";
                                //ws.Cells[4, 2].Value = "Was this report previously evaluated for reliability? (If so, link to those procedures)";
                                //
                                ws.Cells[4, 2].Value = "Was this report previously evaluated for reliability? (If so, link to those procedures) " +
                                    Environment.NewLine + "" +
                                    Environment.NewLine + "If previously base line: 1. Review last modified date: " +
                                    Environment.NewLine + "a. if unchanged from base line testing, no additional testing required " +
                                    Environment.NewLine + "b. if changed, follow testing below";
                                    

                                ws.Cells[5, 2].Value = "What is the purpose of the report?" +
                                    Environment.NewLine + "" ;
                                ws.Cells[6, 2].Value = "What are the key data fields used from this report?";
                                ws.Cells[7, 2].Value = "What type of report is this? [System Generated or Non-system Generated] ";
                                ws.Cells[8, 2].Value = "Define a condition that would represent an exception on this report";
                                ws.Cells[9, 2].Value = "What is the application/system name? (source of the data used to generate the report)";

                                ws.Cells[10, 2].Value = "Is this report: custom, canned, custom query?" +
                                     Environment.NewLine + "" +
                                     Environment.NewLine + "1.	Standard/Canned/Out of Box – a report designed by " +
                                     Environment.NewLine + "software developer and has not been modified or customized by the entity." +
                                     Environment.NewLine + "" + 
                                     Environment.NewLine + "2.	Custom – a modified standard report or a report developed " +
                                     Environment.NewLine + "to meet specific needs of the end user. Reports generated from" +
                                     Environment.NewLine + "to meet specific needs of the end user. Reports generated from" +
                                     Environment.NewLine + "a client-developed (‘in house’) system are considered custom reports";

                                ws.Cells[11, 2].Value = "When was the last time this report was modified?";
                                ws.Cells[12, 2].Value = "What are the key report parameters?";
                                ws.Cells[13, 2].Value = "Are parameters input into the report each time it is run?";
                                ws.Cells[14, 2].Value = "How does the report user verify the report is complete?";
                                ws.Cells[15, 2].Value = "How does the report user verify the report is accurate?";
                                ws.Cells[16, 2].Value = "How does the report user verify the report data has integrity?";
                                ws.Cells[17, 2].Value = "Procedures performed to assess the accuracy, completeness, and validity of the source data." +
                                   Environment.NewLine + "" +
                                   Environment.NewLine + "All Applications" +
                                   Environment.NewLine + "Accuracy - trace one item to the outside source" +
                                   Environment.NewLine + "Completeness - identify an outside data point and trace back to the report" +
                                   Environment.NewLine + "" +
                                   Environment.NewLine + "" +
                                   Environment.NewLine + "" +
                                   Environment.NewLine + "Canned Report - Review report name to ensure it is still unmodified. If true, no additional testing needed." +
                                   Environment.NewLine + "Customized Report - Review report builder for customization and reconcile against the report requirement / purpose " +
                                   Environment.NewLine + "Search - Review search criteria and results tab and reconcile against requirement / purpose ";
                                ws.Cells["B17:B20"].Merge = true;
                                ws.Cells[21, 2].Value = "Who is the report user (report runner)? Name and Title, date of observation.";
                                ws.Cells[22, 2].Value = "When did the tester observe the report being run?";
                                ws.Cells[23, 2].Value = "Frequency of the report/query generation";
                                ws.Cells[24, 2].Value = "Are there other UAT, Change Management testing, ITGC that support the reliability of this Key Report?";
                                ws.Cells[25, 2].Value = "Is the report output modifiable? (Y/N?)";
                                ws.Cells[26, 2].Value = "Testing Information ";
                                ws.Cells[27, 2].Value = "What date was testing performed?";
                                ws.Cells[28, 2].Value = "Who performed the testing?";
                                ws.Cells[29, 2].Value = "What date was testing reviewed?";
                                ws.Cells[30, 2].Value = "Who performed the review?";
                                ws.Cells[31, 2].Value = "What period did the report cover?";
                                ws.Cells[32, 2].Value = "Was testing performed during another period? If yes, when?";
                                ws.Cells[33, 2].Value = "Conclusion";
                                ws.Cells[34, 2].Value = "Any exceptions noted?";
                                ws.Cells[35, 2].Value = "Is the report complete and accurate?";
                                ws.Cells[36, 2].Value = "Notes";

                                StringBuilder sbPurpose = new StringBuilder();
                                StringBuilder sbControlId = new StringBuilder();
                                List<string> listPurpose = new List<string>();
                                List<string> listControlId = new List<string>();
                                int count = 0;

                                //consolidate value from key and non-key and store in list
                                if (listItemIdAll != null && listItemIdAll.Count > 0)
                                {
                                    //here we loop through all the key and non-key value
                                    foreach (var item in listItemIdAll)
                                    {
                                        var leadsheetItem = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(item)).ToList();
                                        if (leadsheetItem != null && leadsheetItem.Count > 0)
                                        {
                                            //Debug.WriteLine($"ItemId : {item}");
                                            var leadsheetPurpose = leadsheetItem.Where(x => x.StrQuestion.ToLower().Contains("what is the purpose of the report")).Select(x => x.StrAnswer).FirstOrDefault();
                                            if (leadsheetPurpose != null)
                                            {
                                                //Debug.WriteLine($"Purpose : {purpose}");
                                                if (!listPurpose.Contains(leadsheetPurpose))
                                                    listPurpose.Add(leadsheetPurpose);
                                            }

                                            var leadsheetControlid = leadsheetItem.Where(x => x.StrQuestion.ToLower().Contains("key control using iuc")).Select(x => x.StrAnswer).FirstOrDefault();
                                            if (leadsheetControlid != null)
                                            {
                                                if (!listControlId.Contains(leadsheetControlid))
                                                    listControlId.Add(leadsheetControlid);
                                            }
                                        }
                                    }
                                }

                                //process list of consolidated data from key and non key
                                if (listPurpose.Count > 0)
                                {
                                    count = 0;
                                    foreach (var item in listPurpose)
                                    {
                                        sbPurpose.Append($"{GetColumnName(count)}: {item}");
                                        //sbPurpose.Append(Environment.NewLine);
                                        count++;
                                    }
                                }

                                if (listControlId.Count > 0)
                                {
                                    foreach (var item in listControlId)
                                    {
                                        sbControlId.Append($"{item}, ");
                                    }
                                    sbControlId = sbControlId.Remove(sbControlId.Length - 2, 2);
                                }

                                var controlId = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("key control using iuc")).Select(x => x.StrAnswer).FirstOrDefault();

                                var controlFrequency = _soxContext.Rcm.Where(x =>
                                        x.FY.ToLower().Equals(filter.Filter.FY.ToLower())
                                        && x.ClientName.ToLower().Equals(filter.Filter.ClientName.ToLower())
                                        && x.ControlId.ToLower().Equals(controlId.ToLower()))
                                    .Select(x => x.ControlFrequency)
                                    .AsNoTracking()
                                    .FirstOrDefault();


                                ws.Cells[3, 3].Value = filter.Filter.KeyReportName;
                                ws.Cells[4, 3].Value = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("was this report previously evaluated for reliability")).Select(x => x.StrAnswer).FirstOrDefault();
                                ws.Cells[5, 3].Value = sbPurpose; //"What is the purpose of the report?";
                                ws.Cells[6, 3].Value = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("what are the key data fields used from this report")).Select(x => x.StrAnswer).FirstOrDefault();
                                ws.Cells[7, 3].Value = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("iuc/key report type")).Select(x => x.StrAnswer).FirstOrDefault();
                                ws.Cells[8, 3].Value = string.Empty; //"Define a condition that would represent an exception on this report";
                                ws.Cells[9, 3].Value = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("key report system name")).Select(x => x.StrAnswer).FirstOrDefault(); //"What is the application/system name? (source of the data used to generate the report)";
                                ws.Cells[10, 3].Value = //keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("key report type")).Select(x => x.StrAnswer).FirstOrDefault(); //"Is this report: custom, canned, custom query?";
                                ws.Cells[11, 3].Value = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("when was the report last modified")).Select(x => x.StrAnswer).FirstOrDefault(); //"When was the last time this report was modified?";
                                ws.Cells[12, 3].Value = //"What are the key report parameters?";
                                ws.Cells[13, 3].Value = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("are parameters input into the report each time it is run")).Select(x => x.StrAnswer).FirstOrDefault();  //"Are parameters input into the report each time it is run?";
                                ws.Cells[14, 3].Value = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("how does the report user verify the report is complete")).Select(x => x.StrAnswer).FirstOrDefault(); //"How does the report user verify the report is complete?";
                                ws.Cells[15, 3].Value = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("how does the report user verify the report is accurate")).Select(x => x.StrAnswer).FirstOrDefault(); //"How does the report user verify the report is accurate?";
                                ws.Cells[16, 3].Value = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("how does the report user verify the report data has integrity")).Select(x => x.StrAnswer).FirstOrDefault(); //"How does the report user verify the report data has integrity?";
                                //ws.Cells[17, 3].Value = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("what are the procedures performed to assess the accuracy, completeness, and validity of the source data")).Select(x => x.StrAnswer).FirstOrDefault(); //"Procedures performed to assess the accuracy, completeness, and validity of the source data.";
                                int counter = 0;
                                var overrideSampleProc = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("override sample procedures")).Select(x => x.StrAnswer).FirstOrDefault();
                                var completenessMethods = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("key report completeness methods")).Select(x => x.StrAnswer).FirstOrDefault(); //"Procedures performed to assess the accuracy, completeness, and validity of the source data.";
                                var completenessAnswer = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("key report completeness answer")).Select(x => x.StrAnswer).FirstOrDefault();
                                var accuracyMethods = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("key report accuracy methods")).Select(x => x.StrAnswer).FirstOrDefault();
                                var accuracyAnswer = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("key report accuracy answer")).Select(x => x.StrAnswer).FirstOrDefault();
                                var parameterMethods = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("key report parameter methods")).Select(x => x.StrAnswer).FirstOrDefault();
                                var parameterAnswer = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("key report parameter answer")).Select(x => x.StrAnswer).FirstOrDefault();
                                var reportMethods = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("key report report methods")).Select(x => x.StrAnswer).FirstOrDefault();
                                var reportAnswer = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("key report report answer")).Select(x => x.StrAnswer).FirstOrDefault();
                                var sampleProcedures = "";
                                var sampleProc_parameter = "";
                                var sampleProc_report = "";
                                var sampleProc_completeness = "";
                                var sampleProc_accuracy = "";
                                //string[] a1, a2, a3, a4, a5, b1, b2, b3, b4, b5, c1, c2, c3, c4, c5, d1, d2, d3, d4, d5, d6;
                                string[] parameterMethodArray = !string.IsNullOrEmpty(parameterMethods) ? parameterMethods.Split(";") : null;
                                string[] reportMethodArray = !string.IsNullOrEmpty(reportMethods) ? reportMethods.Split(";") : null;
                                string[] completenessMethodArray = !string.IsNullOrEmpty(completenessMethods) ?  completenessMethods.Split(";") : null;
                                string[] accuracyMethodArray = !string.IsNullOrEmpty(accuracyMethods) ? accuracyMethods.Split(";") : null;
                                string[] parameterAnswerArray = !string.IsNullOrEmpty(parameterAnswer) ? parameterAnswer.Split("////") : null;
                                string[] reportAnswerArray = !string.IsNullOrEmpty(reportAnswer) ? reportAnswer.Split("////") : null;
                                string[] completenessAnswerArray = !string.IsNullOrEmpty(completenessAnswer) ? completenessAnswer.Split("////") : null;
                                string[] accuracyAnswerArray = !string.IsNullOrEmpty(accuracyAnswer) ? accuracyAnswer.Split("////") : null;
                                if (!string.IsNullOrEmpty(overrideSampleProc))
                                {

                                    sampleProcedures = overrideSampleProc;
                                }
                                else
                                {


                                    counter = 0;
                                    if (parameterMethodArray != null && parameterMethodArray.Any())
                                    {
                                        foreach (var method in parameterMethodArray)
                                        {
                                            if (method != null && method != string.Empty)
                                            {

                                                int result = Int32.Parse(method);
                                                var tempMethod = _soxContext.CAMethodLibrary.Where(x => x.Id.Equals(result)).FirstOrDefault();
                                                if (tempMethod != null)
                                                {
                                                    var parameter_lib = _soxContext.ParametersLibrary.Where(x =>
                                                                            x.KeyReportName.Equals(filter.Filter.KeyReportName)
                                                                            && x.Method.Equals(tempMethod.MethodName)).FirstOrDefault();
                                                    if (parameter_lib != null)
                                                    {

                                                        if (counter == 0)
                                                        {
                                                            sampleProc_parameter += "A. ";
                                                        }
                                                        sampleProc_parameter += parameter_lib.Parameter + "\n \n";
                                                        sampleProc_parameter = sampleProc_parameter.Replace("<p>", string.Empty)
                                                                                                   .Replace("</p>", string.Empty)
                                                                                                   .Replace("</strong", string.Empty)
                                                                                                   .Replace("<br>", "\n")
                                                                                                   .Replace("<strong class=\"text-bold\">", string.Empty)
                                                                                                   .Replace("<br/>", "\n");
                                                        if (parameterAnswerArray != null && parameterAnswerArray[counter] != null)
                                                        {
                                                            string[] tempParamAnswers = parameterAnswerArray[counter].Split(";");
                                                            if (tempParamAnswers.Count() > 0)
                                                            {
                                                                sampleProc_parameter = sampleProc_parameter.Replace("(A1)", tempParamAnswers[0])
                                                                                               .Replace("(A2)", tempParamAnswers[1])
                                                                                               .Replace("(A3)", tempParamAnswers[2])
                                                                                               .Replace("(A4)", tempParamAnswers[3])
                                                                                               .Replace("(A5)", tempParamAnswers[4])
                                                                                               .Replace("(A6)", tempParamAnswers[5])
                                                                                               .Replace("(A7)", tempParamAnswers[6])
                                                                                               .Replace("(A8)", tempParamAnswers[7])
                                                                                               .Replace("(A9)", tempParamAnswers[8])
                                                                                               .Replace("(A10)", tempParamAnswers[9]);
                                                            }

                                                        }


                                                    }
                                                }
                                                counter++;
                                            }

                                        }
                                    }

                                    counter = 0;
                                    if (reportMethodArray != null && reportMethodArray.Any())
                                    {
                                        foreach (var method in reportMethodArray)
                                        {
                                            if (method != null && method != string.Empty)
                                            {
                                                int result = Int32.Parse(method);
                                                var tempMethod = _soxContext.CAMethodLibrary.Where(x => x.Id.Equals(result)).FirstOrDefault();
                                                if (tempMethod != null)
                                                {
                                                    var report_lib = _soxContext.ReportsLibrary.Where(x =>
                                                                            x.KeyReportName.Equals(filter.Filter.KeyReportName)
                                                                            && x.Method.Equals(tempMethod.MethodName)).FirstOrDefault();
                                                    if (report_lib != null)
                                                    {
                                                        if (counter == 0)
                                                        {
                                                            sampleProc_report += "B. ";
                                                        }
                                                        sampleProc_report += report_lib.Report + "\n \n";
                                                        sampleProc_report = sampleProc_report.Replace("<p>", string.Empty)
                                                                                                   .Replace("</p>", string.Empty)
                                                                                                   .Replace("</strong", string.Empty)
                                                                                                   .Replace("<br>", "\n")
                                                                                                   .Replace("<strong class=\"text-bold\">", string.Empty)
                                                                                                   .Replace("<br/>", "\n");
                                                        if (reportAnswerArray != null && reportAnswerArray[counter] != null)
                                                        {
                                                            string[] tempReportAnswers = reportAnswerArray[counter].Split(";");
                                                            if (tempReportAnswers.Count() > 0)
                                                            {
                                                                sampleProc_report = sampleProc_report.Replace("(B1)", tempReportAnswers[0])
                                                                                               .Replace("(B2)", tempReportAnswers[1])
                                                                                               .Replace("(B3)", tempReportAnswers[2])
                                                                                               .Replace("(B4)", tempReportAnswers[3])
                                                                                               .Replace("(B5)", tempReportAnswers[4])
                                                                                               .Replace("(B6)", tempReportAnswers[5])
                                                                                               .Replace("(B7)", tempReportAnswers[6])
                                                                                               .Replace("(B8)", tempReportAnswers[7])
                                                                                               .Replace("(B9)", tempReportAnswers[8])
                                                                                               .Replace("(B10)", tempReportAnswers[9]);
                                                            }



                                                        }

                                                    }
                                                }
                                                counter++;
                                            }

                                        }
                                    }

                                    counter = 0;
                                    if (completenessMethodArray != null && completenessMethodArray.Any())
                                    {
                                        foreach (var method in completenessMethodArray)
                                        {
                                            if (method != null && method != string.Empty)
                                            {
                                                int result = Int32.Parse(method);
                                                var tempMethod = _soxContext.CAMethodLibrary.Where(x => x.Id.Equals(result)).FirstOrDefault();
                                                if (tempMethod != null)
                                                {
                                                    //sampleProc += tempMethod.MethodName + "\n";
                                                    var completeness = _soxContext.CompletenessLibrary.Where(x =>
                                                                            x.KeyReportName.Equals(filter.Filter.KeyReportName)
                                                                            && x.Method.Equals(tempMethod.MethodName)).FirstOrDefault();
                                                    if (completeness != null)
                                                    {

                                                        if (counter == 0)
                                                        {
                                                            sampleProc_completeness += "C. ";
                                                        }
                                                        sampleProc_completeness += completeness.Completeness + "\n \n";
                                                        sampleProc_completeness = sampleProc_completeness.Replace("<p>", string.Empty)
                                                                                                           .Replace("</p>", string.Empty)
                                                                                                           .Replace("</strong", string.Empty)
                                                                                                           .Replace("<br>", "\n")
                                                                                                           .Replace("<strong class=\"text-bold\">", string.Empty)
                                                                                                           .Replace("<br/>", "\n");
                                                        if (completenessAnswerArray != null && completenessAnswerArray[counter] != null)
                                                        {
                                                            string[] tempCompletenessAnswers = completenessAnswerArray[counter].Split(";");
                                                            if (tempCompletenessAnswers.Count() > 0)
                                                            {
                                                                sampleProc_completeness = sampleProc_completeness.Replace("(C1)", tempCompletenessAnswers[0])
                                                                                               .Replace("(C2)", tempCompletenessAnswers[1])
                                                                                               .Replace("(C3)", tempCompletenessAnswers[2])
                                                                                               .Replace("(C4)", tempCompletenessAnswers[3])
                                                                                               .Replace("(C5)", tempCompletenessAnswers[4])
                                                                                               .Replace("(C6)", tempCompletenessAnswers[5])
                                                                                               .Replace("(C7)", tempCompletenessAnswers[6])
                                                                                               .Replace("(C8)", tempCompletenessAnswers[7])
                                                                                               .Replace("(C9)", tempCompletenessAnswers[8])
                                                                                               .Replace("(C10)", tempCompletenessAnswers[9]);
                                                            }


                                                        }


                                                    }


                                                }
                                                counter++;

                                            }

                                        }
                                    }

                                    counter = 0;
                                    if (accuracyMethodArray != null && accuracyMethodArray.Any())
                                    {
                                        foreach (var method in accuracyMethodArray)
                                        {
                                            if (method != null && method != string.Empty)
                                            {
                                                int result = Int32.Parse(method);
                                                var tempMethod = _soxContext.CAMethodLibrary.Where(x => x.Id.Equals(result)).FirstOrDefault();
                                                if (tempMethod != null)
                                                {
                                                    var accuracy_lib = _soxContext.AccuracyLibrary.Where(x =>
                                                                            x.KeyReportName.Equals(filter.Filter.KeyReportName)
                                                                            && x.Method.Equals(tempMethod.MethodName)).FirstOrDefault();
                                                    if (accuracy_lib != null)
                                                    {
                                                        if (counter == 0)
                                                        {
                                                            sampleProc_accuracy += "D. ";
                                                        }
                                                        sampleProc_accuracy += accuracy_lib.Accuracy + "\n \n";
                                                        sampleProc_accuracy = sampleProc_accuracy.Replace("<p>", string.Empty)
                                                                                                   .Replace("</p>", string.Empty)
                                                                                                   .Replace("</strong", string.Empty)
                                                                                                   .Replace("<br>", "\n")
                                                                                                   .Replace("<strong class=\"text-bold\">", string.Empty)
                                                                                                   .Replace("<br/>", "\n");
                                                        if (accuracyAnswerArray != null && accuracyAnswerArray[counter] != null)
                                                        {
                                                            string[] tempAccuracyAnswers = accuracyAnswerArray[counter].Split(";");
                                                            if (tempAccuracyAnswers.Count() > 0)
                                                            {
                                                                sampleProc_accuracy = sampleProc_accuracy.Replace("(D1)", tempAccuracyAnswers[0])
                                                                                               .Replace("(D2)", tempAccuracyAnswers[1])
                                                                                               .Replace("(D3)", tempAccuracyAnswers[2])
                                                                                               .Replace("(D4)", tempAccuracyAnswers[3])
                                                                                               .Replace("(D5)", tempAccuracyAnswers[4])
                                                                                               .Replace("(D6)", tempAccuracyAnswers[5])
                                                                                               .Replace("(D7)", tempAccuracyAnswers[6])
                                                                                               .Replace("(D8)", tempAccuracyAnswers[7])
                                                                                               .Replace("(D9)", tempAccuracyAnswers[8])
                                                                                               .Replace("(D10)", tempAccuracyAnswers[9]);
                                                            }


                                                        }

                                                    }
                                                    counter++;
                                                }
                                            }

                                        }
                                    }

                                    sampleProcedures = sampleProc_parameter + "\n \n" + sampleProc_report + "\n \n" + sampleProc_completeness + "\n \n" + sampleProc_accuracy;
                                    if (string.IsNullOrEmpty(sampleProc_parameter) && string.IsNullOrEmpty(sampleProc_report) && string.IsNullOrEmpty(sampleProc_completeness) && string.IsNullOrEmpty(sampleProc_accuracy))
                                    {
                                        sampleProcedures = string.Empty;
                                    }
                                }


                                ws.Cells[17, 3].Value = sampleProcedures;

                                ws.Cells["C17:C20"].Merge = true;
                                ws.Cells[21, 3].Value = string.Empty; //Who is the report user (report runner)? Name and Title, date of observation.
                                ws.Cells[22, 3].Value = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("meeting date")).Select(x => x.StrAnswer).FirstOrDefault(); //"When did the tester observe the report being run?";
                                ws.Cells[23, 3].Value = //controlFrequency; //Frequency of the report/query generation
                                ws.Cells[24, 3].Value = string.Empty; //"Are there other UAT, Change Management testing, ITGC that support the reliability of this Key Report?";
                                ws.Cells[25, 3].Value = string.Empty; //Is the report output modifiable? (Y/N?)
                                ws.Cells[26, 3].Value = string.Empty; //"Testing Information ";
                                ws.Cells[27, 3].Value = string.Empty; //"What date was testing performed?";
                                ws.Cells[28, 3].Value = string.Empty; //"Who performed the testing?";
                                ws.Cells[29, 3].Value = string.Empty; //"What date was testing reviewed?";
                                ws.Cells[30, 3].Value = keyReportItem.Where(x => x.StrQuestion.ToLower().Equals("10. Reviewer")).Select(x => x.StrAnswer).FirstOrDefault(); //"Who performed the review?";
                                ws.Cells[31, 3].Value = string.Empty; //"What period did the report cover?";
                                ws.Cells[32, 3].Value = string.Empty; //"Was testing performed during another period? If yes, when? ";
                                ws.Cells[33, 3].Value = string.Empty; //"Conclusion";
                                ws.Cells[34, 3].Value = string.Empty; //"Any exceptions noted?";
                                ws.Cells[35, 3].Value = string.Empty; //"Is the report complete and accurate?";
                                ws.Cells[36, 3].Value = keyReportItem.Where(x => x.StrQuestion.ToLower().Equals("28. Notes")).Select(x => x.StrAnswer).FirstOrDefault(); //"Notes";

                                xlsService.ExcelSetBorder(ws, 1, 2, 36, 3);

                                ws.Cells["A" + 1 + ":A" + 2].Merge = true;
                                ws.Cells["B" + 2 + ":C" + 2].Merge = true;


                                xlsService.ExcelSetBackgroundColorAshGray(ws, 1, 2, 2, 3);
                                xlsService.ExcelSetFontColorWhite(ws, 1, 2, 2, 3);

                                //Procedures performed to access accuracy, completeness and validity
                                ws.Row(17).Height = 40;
                                ws.Row(18).Height = 40;
                                ws.Row(19).Height = 150;
                                ws.Row(20).Height = 150;
                                xlsService.ExcelSetBackgroundColorMintGreen(ws, 17, 2, 20, 3);
                                xlsService.ExcelSetFontItalic(ws, 17, 2, 20, 3);

                                xlsService.ExcelSetFontItalic(ws, 10, 2, 10, 3);

                                xlsService.ExcelWrapText(ws, 3, 2, 36, 3); //(worksheet, from row, from column, to row, to column)
                                xlsService.ExcelSetVerticalAlignCenter(ws, 3, 2, 36, 3);


                                //Testing Information
                                xlsService.ExcelSetBackgroundColorAshGray(ws, 26, 2, 26, 3);
                                xlsService.ExcelSetFontColorWhite(ws, 26, 2, 26, 3);
                               

                                //Conclusion
                                xlsService.ExcelSetBackgroundColorAshGray(ws, 33, 2, 33, 3);
                                xlsService.ExcelSetFontColorWhite(ws, 33, 2, 33, 3);
                                xlsService.ExcelSetFontBold(ws, 33, 2, 33, 3);

                                for (int row = 1; row <= 36; row++)
                                {
                                    xlsService.ExcelSetVerticalAlignCenter(ws, row, 1, row, 1);
                                    xlsService.ExcelSetHorizontalAlignCenter(ws, row, 1, row, 1);
                                }


                                xlsService.ExcelSetCalibriLight10(ws, 1, 1, 36, 3);
                                xlsService.ExcelSetCalibriLight12(ws, 1, 1, 2, 3);
                                xlsService.ExcelSetCalibriLight12(ws, 23, 1, 23, 3);
                                
                                xlsService.ExcelSetFontBold(ws, 1, 1, 2, 3);
                                xlsService.ExcelSetFontItalic(ws, 4, 2, 4, 2);
                                xlsService.ExcelSetFontItalic(ws, 10, 2, 10, 2);
                                xlsService.ExcelSetFontItalic(ws, 17, 2, 20, 2);
                                xlsService.ExcelSetFontBold(ws, 26, 2, 26, 3);// Testing Information
                                xlsService.ExcelSetFontBold(ws, 33, 2, 33, 3);//Conclusion


                                //end Tab1 Leadsheet
                                #endregion



                                #region Tab 2
                                //Tab2 (KeyReportName) - Screenshot
                                ws1.View.ZoomScale = 80;
                                //set column width
                                ws1.Column(1).Width = 16.43;
                                ws1.Column(2).Width = 20;
                                ws1.Column(3).Width = 20;
                                ws1.Column(4).Width = 20;
                                ws1.Column(5).Width = 20;

                                //disable grid
                                ws1.View.ShowGridLines = false;

                                ws1.Cells[1, 1].Value = "Key Report:";
                                ws1.Cells[2, 1].Value = "Control #:";
                                ws1.Cells[3, 1].Value = "Provided By:";
                                ws1.Cells[4, 1].Value = "Tested by:";
                                ws1.Cells[5, 1].Value = "Description:";


                                ws1.Cells[1, 2].Value = filter.Filter.KeyReportName; //"Key Report:";
                                ws1.Cells[2, 2].Value = sbControlId; //"Control #:";
                                ws1.Cells[3, 2].Value = //"Provided By:";
                                ws1.Cells[4, 2].Value = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("9. preparer")).Select(x => x.StrAnswer).FirstOrDefault(); //"Tested by:";
                                ws1.Cells[5, 2].Value = //"Description:";
                                ws1.Cells[5, 2].Value = $"(System) Screenshot of {filter.Filter.KeyReportName}";




                                xlsService.ExcelSetBackgroundColorDarkBlue(ws1, 1, 1, 5, 1);

                                xlsService.ExcelSetFontColorWhite(ws1, 1, 1, 5, 1);

                                xlsService.ExcelSetBorder(ws1, 1, 1, 5, 5);
                                ws1.Cells["B" + 1 + ":E" + 1].Merge = true;
                                ws1.Cells["B" + 2 + ":E" + 2].Merge = true;
                                ws1.Cells["B" + 3 + ":E" + 3].Merge = true;
                                ws1.Cells["B" + 4 + ":E" + 4].Merge = true;
                                ws1.Cells["B" + 5 + ":E" + 5].Merge = true;
                                xlsService.ExcelSetArialSize10(ws1, 1, 1, 10, 14);
                                xlsService.ExcelSetFontBold(ws1, 1, 1, 5, 1);

                                //add screenshot to excel tab 2
                                if(filter.ListScreenshotName != null && filter.ListScreenshotName.Count > 0 && !isImageAddedTab2)
                                {
                                    AddExcelImage(ws1, screenshotPath, filter.Filter.ClientName, filter.Filter.KeyReportName, filter.ListScreenshotName, 7);
                                    isImageAddedTab2 = true;
                                }
                                //end of Tab2
                                #endregion



                                #region Tab 3

                                string wsName2 = $"2. {filter.Filter.KeyReportName}";
                                //check if xlsReport is not null, if not null then it has report file
                                if (xlsReport != null)
                                {
                                    //check if worksheet is valid
                                    if (xlsReport.Workbook?.Worksheets?.Count > 0)
                                    {
                                        if (xls.Workbook?.Worksheets?.Count > 0)
                                        {
                                            var checkWsExists = xls.Workbook.Worksheets.Where(name => name.Name.Equals(wsName2)).FirstOrDefault();
                                            if (checkWsExists == null)
                                            {
                                                //copy worksheet from report workbook to leadsheet output file
                                                var ws2 = xls.Workbook.Worksheets.Add(wsName2, xlsReport.Workbook.Worksheets[0]);
                                                //then insert blank row which will be use in the header
                                                ws2.InsertRow(1, 7);
                                                LeadsheetOutputFileTab3(ws2, filter, keyReportItem, sbControlId.ToString());
                                            }
                                        }
                                        ////copy worksheet from report workbook to leadsheet output file
                                        //var ws2 = xls.Workbook.Worksheets.Add($"2. {filter.Filter.KeyReportName}", xlsReport.Workbook.Worksheets[0]);
                                        ////then insert blank row which will be use in the header
                                        //ws2.InsertRow(1, 7);
                                        //LeadsheetOutputFileTab3(ws2, filter, keyReportItem, sbControlId.ToString());
                                    }
                                }
                                else
                                {
                                    if (xls.Workbook?.Worksheets?.Count > 0)
                                    {
                                        var checkWsExists = xls.Workbook.Worksheets.Where(name => name.Name.Equals(wsName2)).FirstOrDefault();
                                        if(checkWsExists == null )
                                        {
                                            var ws2 = xls.Workbook.Worksheets.Add(wsName2);
                                            LeadsheetOutputFileTab3(ws2, filter, keyReportItem, sbControlId.ToString());
                                        }
                                        else
                                        {
                                            var ws2 = xls.Workbook.Worksheets[wsName2];
                                            LeadsheetOutputFileTab3(ws2, filter, keyReportItem, sbControlId.ToString());
                                        }
                                            
                                    } 
                                }
                                //end of Tab3
                                #endregion



                                #region Tab 4
                                string wsName3 = $"3. Sample Testing_Completeness";
                                if (xlsReport != null)
                                {
                                    //check if worksheet is valid
                                    if (xlsReport.Workbook?.Worksheets?.Count > 0)
                                    {
                                        if (xls.Workbook?.Worksheets?.Count > 0)
                                        {
                                            var checkWsExists = xls.Workbook.Worksheets.Where(name => name.Name.Equals(wsName3)).FirstOrDefault();
                                            if (checkWsExists == null)
                                            {
                                                //copy worksheet from report workbook to leadsheet output file
                                                var ws3 = xls.Workbook.Worksheets.Add(wsName3, xlsReport.Workbook.Worksheets[0]);
                                                //then insert blank row which will be use in the header and column
                                                ws3.InsertRow(1, 13);
                                                ws3.InsertColumn(1, 25);
                                                LeadsheetOutputFileTab4(ws3, filter, keyReportItem, sbControlId.ToString(), screenshotPath);

                                                //add screenshot to excel tab 4
                                                if (filter.ListScreenshotName != null && filter.ListScreenshotName.Count > 0 && !isImageAddedTab4)
                                                {
                                                    AddExcelImage(ws3, screenshotPath, filter.Filter.ClientName, filter.Filter.KeyReportName, filter.ListScreenshotName, 13);
                                                    isImageAddedTab4 = true;
                                                }
                                            }
                                        }
                                        ////copy worksheet from report workbook to leadsheet output file
                                        //var ws3 = xls.Workbook.Worksheets.Add("3. Sample Testing_Completeness", xlsReport.Workbook.Worksheets[0]);
                                        ////then insert blank row which will be use in the header and column
                                        //ws3.InsertRow(1, 13);
                                        //ws3.InsertColumn(1, 25);
                                        //LeadsheetOutputFileTab4(ws3, filter, keyReportItem, sbControlId.ToString(), screenshotPath);
                                    }
                                }
                                else
                                {
                                    if (xls.Workbook?.Worksheets?.Count > 0)
                                    {
                                        ExcelWorksheet ws3 = null;
                                        var checkWsExists = xls.Workbook.Worksheets.Where(name => name.Name.Equals(wsName3)).FirstOrDefault();
                                        if (checkWsExists == null)
                                        {
                                            ws3 = xls.Workbook.Worksheets.Add(wsName3);
                                        }
                                        else
                                        {
                                            ws3 = xls.Workbook.Worksheets[wsName3];
                                            
                                        }

                                        LeadsheetOutputFileTab4(ws3, filter, keyReportItem, sbControlId.ToString(), screenshotPath);
                                        //add screenshot to excel tab 4
                                        if (filter.ListScreenshotName != null && filter.ListScreenshotName.Count > 0 && !isImageAddedTab4)
                                        {
                                            AddExcelImage(ws3, screenshotPath, filter.Filter.ClientName, filter.Filter.KeyReportName, filter.ListScreenshotName, 13);
                                            isImageAddedTab4 = true;
                                        }
                                    }
                                    //var ws3 = xls.Workbook.Worksheets.Add("3. Sample Testing_Completeness");
                                    //LeadsheetOutputFileTab4(ws3, filter, keyReportItem, sbControlId.ToString(), screenshotPath);
                                }
                                #endregion



                                #region Tab 5
                                string wsName4 = $"4.Sample Testing_Accuracy";
                                if (xlsReport != null)
                                {
                                    //check if worksheet is valid
                                    if (xlsReport.Workbook?.Worksheets?.Count > 0)
                                    {
                                        if (xls.Workbook?.Worksheets?.Count > 0)
                                        {
                                            var checkWsExists = xls.Workbook.Worksheets.Where(name => name.Name.Equals(wsName4)).FirstOrDefault();
                                            if (checkWsExists == null)
                                            {
                                                //copy worksheet from report workbook to leadsheet output file
                                                var ws4 = xls.Workbook.Worksheets.Add(wsName4, xlsReport.Workbook.Worksheets[0]);
                                                //then insert blank row which will be use in the header and column
                                                ws4.InsertRow(1, 13);
                                                ws4.InsertColumn(1, 25);
                                                LeadsheetOutputFileTab5(ws4, filter, keyReportItem, sbControlId.ToString(), screenshotPath);
                                                //add screenshot to excel tab 5
                                                if (filter.ListScreenshotName != null && filter.ListScreenshotName.Count > 0 && !isImageAddedTab5)
                                                {
                                                    AddExcelImage(ws4, screenshotPath, filter.Filter.ClientName, filter.Filter.KeyReportName, filter.ListScreenshotName, 13);
                                                    isImageAddedTab5 = true;
                                                }
                                            }
                                        }
                                        ////copy worksheet from report workbook to leadsheet output file
                                        //var ws4 = xls.Workbook.Worksheets.Add("4.Sample Testing_Accuracy", xlsReport.Workbook.Worksheets[0]);
                                        ////then insert blank row which will be use in the header and column
                                        //ws4.InsertRow(1, 13);
                                        //ws4.InsertColumn(1, 25);
                                        //LeadsheetOutputFileTab5(ws4, filter, keyReportItem, sbControlId.ToString(), screenshotPath);
                                    }
                                }
                                else
                                {
                                    if (xls.Workbook?.Worksheets?.Count > 0)
                                    {
                                        ExcelWorksheet ws4 = null;
                                        var checkWsExists = xls.Workbook.Worksheets.Where(name => name.Name.Equals(wsName4)).FirstOrDefault();
                                        if (checkWsExists == null)
                                        {
                                            ws4 = xls.Workbook.Worksheets.Add(wsName4);
                                        }
                                        else
                                        {
                                            ws4 = xls.Workbook.Worksheets[wsName4];  
                                        }

                                        LeadsheetOutputFileTab5(ws4, filter, keyReportItem, sbControlId.ToString(), screenshotPath);
                                        //add screenshot to excel tab 5
                                        if (filter.ListScreenshotName != null && filter.ListScreenshotName.Count > 0 && !isImageAddedTab5)
                                        {
                                            AddExcelImage(ws4, screenshotPath, filter.Filter.ClientName, filter.Filter.KeyReportName, filter.ListScreenshotName, 13);
                                            isImageAddedTab5 = true;
                                        }

                                    }
                                    //var ws4 = xls.Workbook.Worksheets.Add("4.Sample Testing_Accuracy");
                                    //LeadsheetOutputFileTab5(ws4, filter, keyReportItem, sbControlId.ToString(), screenshotPath);
                                }



                                //end of Tab5
                                #endregion


                                xls.Workbook.Worksheets.First().Select();
                            }



                            //save file
                            string startupPath = Environment.CurrentDirectory;
                            //string strSourceDownload = startupPath + "\\include\\sampleselection\\download\\";

                            //string strSourceDownload = Path.Combine(startupPath, "include", "keyreport"); 031421
                            string strSourceDownload = Path.Combine(startupPath, "include", "upload", "keyreport", "leadsheet");

                            if (!Directory.Exists(strSourceDownload))
                            {
                                Directory.CreateDirectory(strSourceDownload);
                            }
                            var ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                            //string filename = $"KeyReport-{filter.Filter.ClientName}-{ts}.xlsx";  031221
                            string filename = $"KeyReport-{filter.Filter.KeyReportName}-{filter.Filter.ClientName}-{ts}.xlsx";
                            string strOutput = Path.Combine(strSourceDownload, filename);

                            //Check if file not exists
                            if (System.IO.File.Exists(strOutput))
                            {
                                System.IO.File.Delete(strOutput);
                            }

                            xls.SaveAs(new FileInfo(strOutput));
                            excelFilename = filename;

                            //return status "Ok" with filename
                            return Ok(excelFilename);

                        }

                    }


                }

                return NoContent();

            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GenerateKeyReportTemplateFile {ex}", "ErrorGenerateKeyReportTemplateFile");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GenerateKeyReportTemplateFile");
                if (_environment.IsDevelopment())
                {
                    return BadRequest(ex.ToString());
                }
                else
                {
                    return BadRequest();
                }
            }

            
            //return Ok(excelFilename);
        }


        //[AllowAnonymous]
        [HttpPost("leadsheetPWC/generatePWC")]
        public IActionResult GenerateKeyReportTemplateFilePWC([FromBody] KeyReportScreenshot filter)
        {

            //test 062521========================================================================================
            //List<string> excelFilename = new List<string>();
            string excelFilename = string.Empty;

            try
            {

                List<int> listItemIdAll = new List<int>();
                List<int> listItemIdUnique = new List<int>();
                List<int> listItemIdUniqueTestStatus = new List<int>();
                List<int> listItemIdUniqueException = new List<int>();
                ExcelService xlsService = new ExcelService();
                string appId = _config.GetSection("KeyReportApp").GetSection("AllIUCId").Value;
                string appIdTestStat = _config.GetSection("KeyReportApp").GetSection("TestStatusTrackerId").Value;

                var inputItemId = _soxContext.KeyReportUserInput.Where(x =>
                        x.TagFY.ToLower().Equals(filter.Filter.FY.ToLower())
                        && x.TagClientName.ToLower().Equals(filter.Filter.ClientName.ToLower())
                        && x.TagReportName.ToLower().Equals(filter.Filter.KeyReportName.ToLower())
                        && x.AppId.Equals(appId)
                        && x.TagStatus.ToLower() != "inactive"
                    )
                    .AsNoTracking()
                    .Select(x => x.ItemId)
                    .Distinct()
                    .ToList();
                foreach (var itemId in inputItemId)
                {
                    if (!listItemIdAll.Contains(itemId) && itemId != 0)
                    {
                        listItemIdAll.Add(itemId);
                    }
                }

                if (listItemIdAll != null && listItemIdAll.Count > 0)
                {
                    //Get unique itemid
                    foreach (var item in listItemIdAll)
                    {
                        var inputItemId4 = _soxContext.KeyReportUserInput.Where(x =>
                                        x.ItemId.Equals(item)
                                        && x.StrQuestion.ToLower().Contains("unique key report")
                                        && x.StrAnswer.ToLower().Equals("yes")
                                        && x.TagStatus.ToLower() != "inactive"
                                    )
                                    .AsNoTracking()
                                    .Select(x => new { x.ItemId, x.TagControlId })
                                    .Distinct()
                                    .ToList();

                        foreach (var itemUnique in inputItemId4)
                        {
                            if (!listItemIdUnique.Contains(itemUnique.ItemId) && itemUnique.ItemId != 0)
                            {
                                listItemIdUnique.Add(itemUnique.ItemId);
                                filter.Filter.ControlId = itemUnique.TagControlId;
                            }
                        }
                    }

                    #region TesttingStatus-Orig KeyReport
                   //=========================================//commented not needed for now
                    //Get Testing Status
                    var testingStatusItemId = _soxContext.KeyReportUserInput.Where(x =>
                                        x.TagClientName.Equals(filter.Filter.ClientName)
                                        && x.TagFY.Equals(filter.Filter.FY)
                                        && x.TagReportName.Equals(filter.Filter.KeyReportName)
                                        && x.TagControlId.Equals(filter.Filter.ControlId)
                                        && x.TagStatus.ToLower() != "inactive"
                                        && x.AppId.Equals(appIdTestStat)
                                        && x.StrQuestion.ToLower().Contains("unique key report")
                                        && x.StrAnswer.ToLower().Equals("yes")
                                    )
                                    .AsNoTracking()
                                    .Select(x => x.ItemId)
                                    .Distinct()
                                    .FirstOrDefault();

                    List<KeyReportUserInput> testingStatus = new List<KeyReportUserInput>();
                    if (testingStatusItemId != 0)
                    {

                        testingStatus = _soxContext.KeyReportUserInput.Where(x =>
                                                            x.ItemId.Equals(testingStatusItemId)
                                                        )
                                                        .AsNoTracking()
                                                        .Distinct()
                                                        .ToList();

                    }
                    //====================================================================================
                    
                    #endregion TesttingStatus-Orig KeyReport 

                    //Processing excel output
                    if (listItemIdUnique != null && listItemIdUnique.Count > 0)
                    {
                        //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        using (ExcelPackage xls = new ExcelPackage())
                        {
                            string screenshotDirectory = Directory.GetCurrentDirectory();
                            string screenshotPath = Path.Combine(screenshotDirectory, "include", "sharefile", "download");
                            ExcelPackage xlsReport = null;
                            //check if report filename is not null, then we open the file
                            if (filter.ReportFilename != null && filter.ReportFilename != string.Empty)
                            {
                                string fileReportPath = Path.Combine(screenshotPath, $"{filter.Filter.ClientName}", $"{filter.Filter.KeyReportName}", filter.ReportFilename);
                                //Check if file not exists
                                if (System.IO.File.Exists(fileReportPath))
                                {
                                    FileInfo reportFileInfo = new FileInfo(fileReportPath);
                                    xlsReport = new ExcelPackage(reportFileInfo);
                                }

                            }
                            //end test 062521====================================================================================



                            //List<string> excelFilename = new List<string>();
                            //string excelFilename = string.Empty;
                            //==================================================062421
                            #region create excel_pwc_keyreport
                            //using (ExcelPackage xls = new ExcelPackage()) uncomment after test
                            {
                                // ExcelService xlsService = new ExcelService(); uncomment after test
                                //set sheet name
                                var ws = xls.Workbook.Worksheets.Add("Accept-Reject Testing RuleSet");
                                var ws1 = xls.Workbook.Worksheets.Add("Key Report Template (1)");
                                var ws2 = xls.Workbook.Worksheets.Add("Document A1.1");
                                var ws3 = xls.Workbook.Worksheets.Add("Document A1.2");
                                var ws4 = xls.Workbook.Worksheets.Add("Document A2.1");
                                var ws5 = xls.Workbook.Worksheets.Add("Document A2.2");
                                var ws6 = xls.Workbook.Worksheets.Add("Document A3");
                                var ws7 = xls.Workbook.Worksheets.Add("Document B1");
                                var ws8 = xls.Workbook.Worksheets.Add("Document C1");
                                var ws9 = xls.Workbook.Worksheets.Add("Document C2");
                                var ws10 = xls.Workbook.Worksheets.Add("Document D1");

                                #region Tab 1 Accept-Reject Testing RuleSet
                                ws.View.ZoomScale = 90;
                                //set column width
                                ws.Column(1).Width = 26; //Desired Level of Evidence
                                ws.Column(2).Width = 23;    //No exceptions tolerated    
                                ws.Column(3).Width = 23;    //1 exception tolerated
                                ws.Column(4).Width = 23;    //2 exceptions tolerated

                                //disable grid
                                ws.View.ShowGridLines = false;

                                //Format Cells
                                //merge cells

                                ws.Cells["A5:A6"].Merge = true;
                                ws.Cells["B5:D5"].Merge = true;

                                xlsService.ExcelSetAlignCenter(ws, "A5:A6");
                                xlsService.ExcelSetAlignCenter(ws, "B5:D5");
                                xlsService.ExcelSetAlignCenter(ws, "B19:B22");

                                xlsService.ExcelSetBackgroundColorLightRed(ws, 5, 1, 6, 4);
                                xlsService.ExcelSetFontColorWhite(ws, 5, 1, 6, 4);
                                xlsService.ExcelSetBackgroundColorLightRed(ws, 18, 1, 18, 2);
                                xlsService.ExcelSetFontColorWhite(ws, 18, 1, 18, 2);


                                xlsService.ExcelSetBorderRed(ws, 5, 1, 9, 4);
                                xlsService.ExcelSetBorderRed(ws, 18, 1, 22, 2);

                                xlsService.ExcelSetArialSize10(ws, 1, 1, 37, 20);
                                xlsService.ExcelSetFontBold(ws, 1, 1, 1, 1);
                                xlsService.ExcelSetFontBold(ws, 12, 1, 12, 1);
                                xlsService.ExcelSetFontBold(ws, 27, 1, 34, 1);


                                //set table header
                                //Tab1 Accept-Reject Testing RuleSet
                                ws.Cells[1, 1].Value = "Population of 200 items or more";
                                ws.Cells[3, 1].Value = "When applying accept-reject testing to populations of 200 or more items, we apply the sample sizes as shown in the table below.";
                                ws.Cells[5, 1].Value = "Desired Level of Evidence";
                                ws.Cells[5, 2].Value = "Number of items to Test";
                                ws.Cells[6, 2].Value = "No exceptions tolerated";
                                ws.Cells[6, 3].Value = "1 exception tolerated";
                                ws.Cells[6, 4].Value = "2 exceptions tolerated";
                                ws.Cells[7, 1].Value = "Low";
                                ws.Cells[7, 2].Value = "16";
                                ws.Cells[7, 3].Value = "32";
                                ws.Cells[7, 4].Value = "52";
                                ws.Cells[8, 1].Value = "Moderate";
                                ws.Cells[8, 2].Value = "30";
                                ws.Cells[8, 3].Value = "55";
                                ws.Cells[8, 4].Value = "80";
                                ws.Cells[9, 1].Value = "High";
                                ws.Cells[9, 2].Value = "55";
                                ws.Cells[9, 3].Value = "85";
                                ws.Cells[9, 4].Value = "115";
                                ws.Cells[12, 1].Value = "Populations of fewer than 200 items";
                                ws.Cells[14, 1].Value = "When applying accept-reject testing to populations fewer than 200 items, the sample sizes as shown in the table below are appplied. ";
                                ws.Cells[15, 1].Value = "These sample sizes are minimum sample sizes to obtain a low level of evidence. ";
                                ws.Cells[16, 1].Value = "If a higher level of evidence is required, we test more than the minimum number of items. Judgment is required to determine the appropriate sample size.";

                                ws.Cells[18, 1].Value = "Population range";
                                ws.Cells[18, 2].Value = "Number of items to test";
                                ws.Cells[19, 1].Value = "Between 100 and 199 items";
                                ws.Cells[19, 2].Value = "10";
                                ws.Cells[20, 1].Value = "Between 50 and 99 items";
                                ws.Cells[20, 2].Value = "5";
                                ws.Cells[21, 1].Value = "Between 20 and 49 items";
                                ws.Cells[21, 2].Value = "3";
                                ws.Cells[22, 1].Value = "Fewer than 20 items";
                                ws.Cells[22, 2].Value = "2";

                                ws.Cells[24, 1].Value = "When testing 10 or fewer items (i.e., for populations of fewer than 200 items), no exceptions can be expected.";
                                ws.Cells[27, 1].Value = "PWC Accept-Reject testing";
                                ws.Cells[29, 1].Value = "Accuracy:";
                                ws.Cells[29, 2].Value = "Trace 16 samples from report to system";
                                ws.Cells[30, 1].Value = "Completeness:";
                                ws.Cells[30, 2].Value = "1) count of rows; OR";
                                ws.Cells[31, 2].Value = "2) total amounts; OR";
                                ws.Cells[32, 2].Value = "3) sum of key fields BETWEEN the system and report (not just counting rows on report output).";
                                ws.Cells[33, 2].Value = "Last resort is select 16 transactions from the system directly and trace to report";
                                ws.Cells[34, 1].Value = "System:";
                                ws.Cells[34, 2].Value = "Looking up transaction record in system (not agreeing back to report shown on the screen, not going back to orig invoice)";
                                #endregion Tab1

                                #region Tab 2 Key Report Template (1)
                                ws1.View.ZoomScale = 90;
                                //set column width
                                ws1.Column(1).Width = 1.5;
                                ws1.Column(2).Width = 5;
                                ws1.Column(3).Width = 5;
                                ws1.Column(4).Width = 38;
                                ws1.Column(5).Width = 15;
                                ws1.Column(6).Width = 5;
                                ws1.Column(7).Width = 5;
                                ws1.Column(8).Width = 55;
                                ws1.Column(9).Width = 25;
                                ws1.Column(67).Width = 15;


                                //disable grid
                                ws1.View.ShowGridLines = false;

                                ws1.View.FreezePanes(6, 1);

                                //format
                                xlsService.ExcelSetBackgroundColorMidGray(ws1, 1, 1, 5, 1);
                                //xlsService.ExcelSetBorder(ws1, 1, 2, 349, 9);
                                xlsService.ExcelSetArialSize10(ws1, 1, 1, 10, 14);
                                xlsService.ExcelSetFontBold(ws1, 1, 1, 5, 1);


                                xlsService.ExcelWrapText(ws1, 2, 2, 342, 5); //(worksheet, from row, from column, to row, to column)
                                xlsService.ExcelSetBackgroundColorDarkRed(ws1, 1, 2, 1, 8); // row1
                                xlsService.ExcelSetBackgroundColorDarkRed(ws1, 6, 2, 6, 8); //row6 General Info
                                xlsService.ExcelSetBackgroundColorDarkRed(ws1, 54, 2, 54, 8); //row54 Generated by
                                xlsService.ExcelSetBackgroundColorDarkRed(ws1, 67, 2, 67, 8); //row67 
                                xlsService.ExcelSetBackgroundColorDarkRed(ws1, 144, 2, 144, 8); //row144 Benchmarking
                                xlsService.ExcelSetBackgroundColorDarkRed(ws1, 237, 2, 237, 8); //row237 PWC Procedures
                                xlsService.ExcelSetBackgroundColorDarkRed(ws1, 330, 2, 330, 8); //row330 Conclusion

                                xlsService.ExcelSetFontColorWhite(ws1, 1, 2, 1, 8); // row1
                                xlsService.ExcelSetFontColorWhite(ws1, 6, 2, 6, 8); //row6 General Info
                                xlsService.ExcelSetFontColorWhite(ws1, 54, 2, 54, 8); //row54 Generated by
                                xlsService.ExcelSetFontColorWhite(ws1, 67, 2, 67, 8); //row67 
                                xlsService.ExcelSetFontColorWhite(ws1, 144, 2, 144, 8); //row144 Benchmarking
                                xlsService.ExcelSetFontColorWhite(ws1, 237, 2, 237, 8);//row237 PWC Procedures
                                xlsService.ExcelSetFontColorWhite(ws1, 330, 2, 330, 8);//row330 Conclusion

                                xlsService.ExcelSetFontColorCustom(ws1, 5, 2, 5, 9, "7A1818");//row 5
                                xlsService.ExcelSetBackgroundColorMidGray(ws1, 5, 2, 5, 9);

                                xlsService.ExcelSetBackgroundColorMidGray(ws1, 343, 2, 343, 9); //row 343

                                xlsService.ExcelSetFontColorCustom(ws1, 345, 2, 346, 9, "7A1818");//row 345,346
                                xlsService.ExcelSetFontColorCustom(ws1, 347, 2, 347, 9, "DC6900");//row 347,349
                                xlsService.ExcelSetBackgroundColorMidGray(ws1, 345, 2, 346, 9);

                                //BGCOLOR for answer column
                                xlsService.ExcelSetBackgroundColorSkinTone(ws1, 2, 6, 4, 9);
                                xlsService.ExcelSetBackgroundColorSkinTone(ws1, 7, 6, 9, 8);
                                xlsService.ExcelSetBackgroundColorSkinTone(ws1, 10, 6, 12, 8);
                                xlsService.ExcelSetBackgroundColorSkinTone(ws1, 16, 6, 18, 8);
                                xlsService.ExcelSetBackgroundColorSkinTone(ws1, 39, 6, 46, 8);
                                xlsService.ExcelSetBackgroundColorSkinTone(ws1, 50, 6, 53, 8);
                                xlsService.ExcelSetBackgroundColorSkinTone(ws1, 68, 6, 70, 8);
                                xlsService.ExcelSetBackgroundColorSkinTone(ws1, 91, 6, 93, 8);
                                xlsService.ExcelSetBackgroundColorSkinTone(ws1, 106, 6, 109, 8);
                                xlsService.ExcelSetBackgroundColorSkinTone(ws1, 119, 6, 120, 8);
                                xlsService.ExcelSetBackgroundColorSkinTone(ws1, 131, 6, 143, 8);
                                xlsService.ExcelSetBackgroundColorSkinTone(ws1, 238, 6, 255, 8);
                                xlsService.ExcelSetBackgroundColorSkinTone(ws1, 268, 6, 269, 8);
                                xlsService.ExcelSetBackgroundColorSkinTone(ws1, 301, 6, 302, 8);
                                xlsService.ExcelSetBackgroundColorSkinTone(ws1, 315, 6, 316, 8);
                                xlsService.ExcelSetBackgroundColorSkinTone(ws1, 324, 6, 329, 8);
                                xlsService.ExcelSetBackgroundColorSkinTone(ws1, 334, 6, 342, 8);
                                xlsService.ExcelSetBackgroundColorSkinTone(ws1, 347, 2, 349, 9);


                                xlsService.ExcelSetVerticalAlignTop(ws1, "C2:D349");
                                xlsService.ExcelSetVerticalAlignTop(ws1, "D347:D349");

                                xlsService.ExcelSetArialSize10(ws1, 1, 1, 349, 9); // font for all except those with specific format
                                xlsService.ExcelSetFontBold(ws1, 1, 1, 1, 9);//row1
                                xlsService.ExcelSetArialSize16(ws1, 1, 1, 1, 9); //row1
                                xlsService.ExcelSetArialSize16(ws1, 345, 1, 345, 9); //row345
                                xlsService.ExcelSetFontBold(ws1, 1, 1, 5, 1); //row1
                                xlsService.ExcelSetArialSize12(ws1, 5, 1, 5, 9); //row5
                                xlsService.ExcelSetArialSize12(ws1, 346, 1, 346, 9); //row346
                                xlsService.ExcelSetFontBold(ws1, 5, 1, 5, 9); //row5
                                xlsService.ExcelSetFontBold(ws1, 6, 1, 6, 9); //row6
                                xlsService.ExcelSetFontBold(ws1, 54, 1, 54, 9); //row54
                                xlsService.ExcelSetFontBold(ws1, 67, 1, 67, 9); //row67
                                xlsService.ExcelSetFontBold(ws1, 144, 1, 144, 9); //row144
                                xlsService.ExcelSetFontBold(ws1, 237, 1, 237, 9); //row237
                                xlsService.ExcelSetFontBold(ws1, 330, 1, 330, 9); //row330
                                xlsService.ExcelSetFontBold(ws1, 345, 1, 346, 9); //row330

                                xlsService.ExcelSetFontItalic(ws1, 345, 1, 347, 9); //row330




                                //merge cells
                                //merge red bg
                                //ws1.Cells["B1:I1"].Style.Border.BorderAround;

                                ws1.Cells["B1:I343"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["F2:H343"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["B1:I1"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["B5:I5"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["B6:I6"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["B7:I9"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["B10:I12"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["B13:I15"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["B16:I18"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["B19:I32"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["B35:I38"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["B39:I42"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["B43:I46"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["B47:I49"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["B50:I53"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["B54:I54"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);

                                ws1.Cells["B67:I67"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["B68:I70"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["B71:I79"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["B82:114"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["B115:I130"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["B131:I134"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["B135:I143"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);

                                ws1.Cells["B144:I144"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["B145:I147"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);


                                ws1.Cells["B237:I237"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["B238:I240"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["B241:I243"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["B237:I237"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["B244:I246"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["B247:I249"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["B250:I252"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["B253:I255"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["B256:I295"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["B296:I303"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["B306:I323"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["B324:I326"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["B327:I329"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);

                                ws1.Cells["B330:I330"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin); //CONCLUSION
                                ws1.Cells["B331:I333"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["B334:I336"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["B337:I339"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["B340:I342"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);

                                ws1.Cells["B343:I343"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);

                                ws1.Cells["B345:I345"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);//Additional engagement
                                ws1.Cells["B346:E346"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["F346:I346"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["B347:E349"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                ws1.Cells["F347:I349"].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);


                                ws1.Cells["B1:I1"].Merge = true;  // row1
                                ws1.Cells["B6:I6"].Merge = true; //row6 General Info
                                ws1.Cells["B54:I54"].Merge = true;//row54 Generated by
                                ws1.Cells["B67:I67"].Merge = true;//row67
                                ws1.Cells["B144:I144"].Merge = true;//row144 Benchmarking
                                ws1.Cells["B237:I237"].Merge = true;//row237 PWC Procedures
                                ws1.Cells["B330:I330"].Merge = true;//row330 Conclusion
                                ws1.Cells["B343:I343"].Merge = true;//row343
                                ws1.Cells["B345:I345"].Merge = true;//row345 
                                //ws1.Cells["B346:E346"].Merge = true;//row346 
                                ws1.Cells["F346:I346"].Merge = true;//row346 

                                ws1.Cells["B2:E4"].Merge = true;
                                ws1.Cells["F2:I4"].Merge = true;
                                ws1.Cells["B5:E5"].Merge = true;
                                ws1.Cells["F5:H5"].Merge = true;
                                ws1.Cells["B6:I6"].Merge = true;

                                ws1.Cells["C7:D9"].Merge = true;
                                ws1.Cells["E7:E9"].Merge = true;
                                ws1.Cells["F7:H9"].Merge = true;

                                //update merge till colE
                                ws1.Cells["C10:E12"].Merge = true;
                                ws1.Cells["C13:E15"].Merge = true;
                                ws1.Cells["C16:E18"].Merge = true;
                                ws1.Cells["C19:E20"].Merge = true;
                                ws1.Cells["C35:E38"].Merge = true;
                                ws1.Cells["C39:E42"].Merge = true;
                                ws1.Cells["C43:E46"].Merge = true;
                                ws1.Cells["C47:E49"].Merge = true;
                                ws1.Cells["C50:E53"].Merge = true;
                                ws1.Cells["C55:E57"].Merge = true;
                                ws1.Cells["C68:E70"].Merge = true;
                                ws1.Cells["C71:E78"].Merge = true;
                                ws1.Cells["C82:E114"].Merge = true;
                                ws1.Cells["C115:E130"].Merge = true;
                                ws1.Cells["C131:E134"].Merge = true;
                                ws1.Cells["C135:E143"].Merge = true;
                                ws1.Cells["C145:E147"].Merge = true;
                                //hide row 148-236
                                ws1.Row(148).Height = 0;
                                ws1.Cells["C238:E240"].Merge = true;
                                ws1.Cells["C241:E243"].Merge = true;
                                ws1.Cells["C244:E246"].Merge = true;
                                ws1.Cells["C247:E249"].Merge = true;
                                ws1.Cells["C250:E252"].Merge = true;
                                ws1.Cells["C253:E255"].Merge = true;
                                ws1.Cells["C256:E295"].Merge = true;
                                ws1.Cells["C296:E303"].Merge = true;
                                ws1.Cells["C306:E323"].Merge = true;
                                ws1.Cells["C324:E326"].Merge = true;
                                ws1.Cells["C327:E329"].Merge = true;
                                ws1.Cells["C331:E333"].Merge = true;
                                ws1.Cells["C334:E336"].Merge = true;
                                ws1.Cells["C337:E339"].Merge = true;
                                ws1.Cells["C340:E342"].Merge = true;
                                ws1.Cells["C348:E349"].Merge = true;

                                //ROW height
                                ws1.Row(67).Height = 37;
                                ws1.Row(67).Height = 37;

                                //Hide Rows
                                for (int hideRow = 21; hideRow < 29; hideRow++)
                                {
                                    ws1.Row(hideRow).Height = 0;
                                }

                                for (int hideRow = 33; hideRow < 35; hideRow++)
                                {
                                    ws1.Row(hideRow).Height = 0;
                                }

                                for (int hideRow = 58; hideRow < 67; hideRow++)
                                {
                                    ws1.Row(hideRow).Height = 0;
                                }

                                for (int hideRow = 80; hideRow < 82; hideRow++)
                                {
                                    ws1.Row(hideRow).Height = 0;
                                }

                                for (int hideRow = 85; hideRow < 89; hideRow++)
                                {
                                    ws1.Row(hideRow).Height = 0;
                                }

                                for (int hideRow = 111; hideRow < 115; hideRow++)
                                {
                                    ws1.Row(hideRow).Height = 0;
                                }


                                for (int hideRow = 122; hideRow < 124; hideRow++)
                                {
                                    ws1.Row(hideRow).Height = 0;
                                }


                                for (int hideRow = 148; hideRow < 237; hideRow++)
                                {
                                    ws1.Row(hideRow).Height = 0;
                                }

                                for (int hideRow = 265; hideRow < 267; hideRow++)
                                {
                                    ws1.Row(hideRow).Height = 0;
                                }

                                for (int hideRow = 274; hideRow < 276; hideRow++)
                                {
                                    ws1.Row(hideRow).Height = 0;
                                }

                                for (int hideRow = 277; hideRow < 280; hideRow++)
                                {
                                    ws1.Row(hideRow).Height = 0;
                                }

                                for (int hideRow = 282; hideRow < 285; hideRow++)
                                {
                                    ws1.Row(hideRow).Height = 0;
                                }

                                //set table header
                                //Tab2 Key Report Template (1)
                                ws1.Cells[1, 2].Value = "Access the Reliability of Information Used by Management or the Auditor";
                                ws1.Cells[2, 2].Value = "When more than one preparer was involved in the completion of this." +
                                                   Environment.NewLine + "EGA, document the names of the team members involved and the" +
                                                   Environment.NewLine + "procedures performed.";

                                ws1.Cells[5, 2].Value = "Procedures";
                                ws1.Cells[5, 6].Value = "Results";
                                ws1.Cells[5, 9].Value = "Links";
                                ws1.Cells[6, 2].Value = "General Information";
                                ws1.Cells[7, 2].Value = "1.";
                                ws1.Cells[7, 3].Value = "Key Report";
                                ws1.Cells[10, 2].Value = "2.";
                                ws1.Cells[10, 3].Value = "Application / system name - source of the data used to generate the report?";
                                ws1.Cells[13, 2].Value = "3.";
                                ws1.Cells[13, 3].Value = "Is a reporting tool being used to generate the report?";
                                ws1.Cells[16, 2].Value = "4.";
                                ws1.Cells[16, 3].Value = "What reporting tool is used to generate the report/perform the query?";
                                ws1.Cells[19, 2].Value = "5.";
                                ws1.Cells[19, 3].Value = "What type of report was used?";
                                ws1.Cells[35, 2].Value = "6.";
                                ws1.Cells[35, 3].Value = "Is the report used by management in the execution of a relevant control or used by the engagement team in substantive procedures?";
                                ws1.Cells[39, 2].Value = "7.";
                                ws1.Cells[39, 3].Value = "Control(s) / substantive procedure(s) dependent on the reliability of the report";
                                ws1.Cells[43, 2].Value = "8.";
                                ws1.Cells[43, 3].Value = "Describe the purpose of the report and relevance to the audit, including what details in the report are being relied upon";
                                ws1.Cells[47, 2].Value = "9.";
                                ws1.Cells[47, 3].Value = "Do we consider the system-generated information in this report to be appropriate for its intended uses?";
                                ws1.Cells[50, 2].Value = "10.";
                                ws1.Cells[50, 3].Value = "Description of procedures performed to assess the accuracy, completeness, and validity of the source data.";
                                ws1.Cells[54, 2].Value = "Report Generated by Third-Party Service Provider ";
                                ws1.Cells[55, 2].Value = "11.";
                                ws1.Cells[55, 3].Value = "Is the system-generated information (key report) designed and produced by a third-party service provider?";
                                ws1.Cells[67, 2].Value = "If the report is used in the execution of a relevant control, understand management control(s) to determine the information is reliable for its purpose (Design & Implementation)";
                                ws1.Cells[68, 2].Value = "12.";
                                ws1.Cells[68, 3].Value = "Control Operator - Name and Title";
                                ws1.Cells[71, 2].Value = "13.";
                                ws1.Cells[71, 3].Value = "Frequency of the report/query generation";
                                ws1.Cells[82, 2].Value = "14.";
                                ws1.Cells[82, 3].Value = "Describe the management's control(s) designed to asses the reliability of key report (i.e., accurate, complete and reliable for its purpose).";
                                ws1.Cells[115, 2].Value = "15.";
                                ws1.Cells[115, 3].Value = "If parameters are input each time the report is run," +
                                                   Environment.NewLine + "these parameters (e.g., date ranges) should be" +
                                                   Environment.NewLine + "assessed each time the report is used." +
                                                   Environment.NewLine + "" +
                                                   Environment.NewLine + "If applicable, describe management control(s) to" +
                                                   Environment.NewLine + "assess the appropriateness of the parameters each time " +
                                                   Environment.NewLine + "the report is used";
                                ws1.Cells[131, 2].Value = "16.";
                                ws1.Cells[131, 3].Value = "Is the system generated output modifiable?";
                                ws1.Cells[135, 2].Value = "17.";
                                ws1.Cells[135, 3].Value = "An audit of ICFR, or a financial statement only" +
                                                   Environment.NewLine + "audit where we plan to rely on controls and modify" +
                                                   Environment.NewLine + "the nature, timing, and extent of planned " +
                                                   Environment.NewLine + "substantive procedures, describe the nature, timing," +
                                                   Environment.NewLine + "and extent of controls testing procedures to assess the" +
                                                   Environment.NewLine + "operating effectiveness of management's control(s)" +
                                                   Environment.NewLine + "noted above (i.e., management's control(s) designed" +
                                                   Environment.NewLine + "to assess: the reliability of the key report (i.e.," +
                                                   Environment.NewLine + "noted above (accurate, complete and reliable for its purpose);" +
                                                   Environment.NewLine + "parameters, if applicable; and modifiable outputs, if" +
                                                   Environment.NewLine + "applicable).";
                                ws1.Cells[144, 2].Value = "Benchmarking Considerations (if applicable)";
                                ws1.Cells[145, 2].Value = "18.";
                                ws1.Cells[145, 3].Value = "Has the engagement team previously evaluated the reliability of this report?";
                                ws1.Cells[237, 2].Value = "PwC Procedures over Reliability of Information";
                                ws1.Cells[238, 2].Value = "19.";
                                ws1.Cells[238, 3].Value = "Test Date";
                                ws1.Cells[241, 2].Value = "20.";
                                ws1.Cells[241, 3].Value = "Period Covered";
                                ws1.Cells[244, 2].Value = "21.";
                                ws1.Cells[244, 3].Value = "If testing is performed at an interim date, document considerations or link to update testing procedures.";
                                ws1.Cells[247, 2].Value = "22.";
                                ws1.Cells[247, 3].Value = "Extent of Testing";
                                ws1.Cells[250, 2].Value = "23.";
                                ws1.Cells[250, 3].Value = "Nature of Testing";
                                ws1.Cells[253, 2].Value = "24.";
                                ws1.Cells[253, 3].Value = "Client Contact(s)";
                                ws1.Cells[256, 2].Value = "25.";
                                ws1.Cells[256, 3].Value = "The testing objectives we perform to assess accuracy and completeness may vary depending on the complexity of the system, the type of report, and how the report is used in the audit." +
                                                 Environment.NewLine + "Describe testing procedures performed.";
                                ws1.Cells[296, 2].Value = "26.";
                                ws1.Cells[296, 3].Value = "If parameters are input each time the report is run, these parameters (e.g., date ranges) should be assessed each time the report is used to support our testing.";
                                ws1.Cells[306, 2].Value = "27.";
                                ws1.Cells[306, 3].Value = "There are multiple factors to consider when assessing the mathematical accuracy of reports.  These factors vary based on the type of report (i.e., Standard, Custom, or Query),  as well as the complexity of the underlying calculations being performed by the report.";
                                ws1.Cells[324, 2].Value = "28.";
                                ws1.Cells[324, 3].Value = "Define condition(s) that represent an exception";
                                ws1.Cells[327, 2].Value = "29.";
                                ws1.Cells[327, 3].Value = "Results of PwC testing performed (including assessment of the completeness and accuracy of the report/query/ data)";
                                ws1.Cells[330, 2].Value = "Conclusion";
                                ws1.Cells[331, 2].Value = "30.";
                                ws1.Cells[331, 3].Value = "Exceptions Noted?";
                                ws1.Cells[334, 2].Value = "31.";
                                ws1.Cells[334, 3].Value = "Control deficiency explanation (if an exception has been noted)";
                                ws1.Cells[337, 2].Value = "32.";
                                ws1.Cells[337, 3].Value = "Root Cause Analysis (if applicable)";
                                ws1.Cells[340, 2].Value = "33.";
                                ws1.Cells[340, 3].Value = "Audit Response (including identification of compensating controls and/or mitigating information) (if applicable)";
                                ws1.Cells[345, 2].Value = "Additional engagement specific procedures, if necessary";
                                ws1.Cells[346, 2].Value = "Procedures";
                                ws1.Cells[346, 6].Value = "Results";
                                ws1.Cells[347, 2].Value = "[Document engagement specific additional procedures, if necessary]";
                                ws1.Cells[347, 6].Value = "[Document results of additional procedures]";

                                #region Tab2 Data Population

                                //===========================================

                                foreach (var itemId in listItemIdUnique)
                                {
                                    //Get All IUC
                                    var keyReportItem = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(itemId)).ToList();

                                    StringBuilder sbPurpose = new StringBuilder();
                                    StringBuilder sbControlId = new StringBuilder();
                                    List<string> listPurpose = new List<string>();
                                    List<string> listControlId = new List<string>();
                                    int count = 0;

                                    //consolidate value from key and non-key and store in list
                                    if (listItemIdAll != null && listItemIdAll.Count > 0)
                                    {
                                        //here we loop through all the key and non-key value
                                        foreach (var item in listItemIdAll)
                                        {
                                            var leadsheetItem = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(item)).ToList();
                                            if (leadsheetItem != null && leadsheetItem.Count > 0)
                                            {
                                                //Debug.WriteLine($"ItemId : {item}");
                                                var leadsheetPurpose = leadsheetItem.Where(x => x.StrQuestion.ToLower().Contains("what is the purpose of the report")).Select(x => x.StrAnswer).FirstOrDefault();
                                                if (leadsheetPurpose != null)
                                                {
                                                    //Debug.WriteLine($"Purpose : {purpose}");
                                                    if (!listPurpose.Contains(leadsheetPurpose))
                                                        listPurpose.Add(leadsheetPurpose);
                                                }

                                                var leadsheetControlid = leadsheetItem.Where(x => x.StrQuestion.ToLower().Contains("key control using iuc")).Select(x => x.StrAnswer).FirstOrDefault();
                                                if (leadsheetControlid != null)
                                                {
                                                    if (!listControlId.Contains(leadsheetControlid))
                                                        listControlId.Add(leadsheetControlid);
                                                }
                                            }
                                        }
                                    }

                                    //process list of consolidated data from key and non key
                                    if (listPurpose.Count > 0)
                                    {
                                        count = 0;
                                        foreach (var item in listPurpose)
                                        {
                                            sbPurpose.Append($"{GetColumnName(count)}: {item}");
                                            //sbPurpose.Append(Environment.NewLine);
                                            count++;
                                        }
                                    }

                                    if (listControlId.Count > 0)
                                    {
                                        foreach (var item in listControlId)
                                        {
                                            sbControlId.Append($"{item}, ");
                                        }
                                        sbControlId = sbControlId.Remove(sbControlId.Length - 2, 2);
                                    }

                                    var controlId = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("key control using iuc")).Select(x => x.StrAnswer).FirstOrDefault();

                                    var controlFrequency = _soxContext.Rcm.Where(x =>
                                            x.FY.ToLower().Equals(filter.Filter.FY.ToLower())
                                            && x.ClientName.ToLower().Equals(filter.Filter.ClientName.ToLower())
                                            && x.ControlId.ToLower().Equals(controlId.ToLower()))
                                        .Select(x => x.ControlFrequency)
                                        .AsNoTracking()
                                        .FirstOrDefault();

                                    ws1.Cells[7, 6].Value = filter.Filter.KeyReportName; //1."Key Report:";
                                    ws1.Cells[10, 6].Value = "";//2. source of data;
                                    ws1.Cells[13, 6].Value = "";//3. Yes/No;Is used to generate report?
                                    ws1.Cells[16, 6].Value = "";//4. reporting toolused to generate report?
                                    ws1.Cells[19, 6].Value = "";//5. What type of report
                                    ws1.Cells[35, 6].Value = "";//6. What type of report
                                    ws1.Cells[399, 6].Value = "";//7. What type of report
                                    ws1.Cells[43, 6].Value = sbPurpose;//8. purpose of the report
                                    ws1.Cells[71, 6].Value = controlFrequency;//13. frequency
                                    ws1.Cells[324, 6].Value = "";//28. exception
                                                                            

                                    //===========================================


                                    #endregion Tab2 Data Population

                                    #endregion Tab 2

                                    #region Tab 3 Document A1.1
                                    ws2.View.ZoomScale = 90;
                                    //set column width
                                    ws2.Column(1).Width = 16.43;
                                    ws2.Column(2).Width = 20;
                                    ws2.Column(3).Width = 20;
                                    ws2.Column(4).Width = 20;
                                    ws2.Column(5).Width = 20;

                                    //disable grid
                                    ws2.View.ShowGridLines = false;
                                    ws2.View.FreezePanes(8, 2);

                                    ws2.Cells[1, 1].Value = "Key Report:";
                                    ws2.Cells[2, 1].Value = "Control #:";
                                    ws2.Cells[3, 1].Value = "Provided By:";
                                    ws2.Cells[4, 1].Value = "Tested by:";
                                    ws2.Cells[5, 1].Value = "Description:";

                                    //populate field values
                                    ws2.Cells[1, 2].Value = filter.Filter.KeyReportName; //"Key Report:";

                                    ws2.Cells[2, 2].Value = sbControlId; //"Control #:";
                                    ws2.Cells[3, 2].Value = //"Provided By:";
                                    ws2.Cells[4, 2].Value = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("14. preparer")).Select(x => x.StrAnswer).FirstOrDefault(); //"Tested by:";
                                    ws2.Cells[5, 2].Value = //"Description:";

                                    ws2.Cells[5, 2].Value = $"(System) Screenshot of {filter.Filter.KeyReportName}";


                                    xlsService.ExcelSetBackgroundColorLightGray(ws2, 1, 1, 5, 1);

                                    xlsService.ExcelSetBorder(ws2, 1, 1, 5, 5);
                                    ws2.Cells["B" + 1 + ":E" + 1].Merge = true;
                                    ws2.Cells["B" + 2 + ":E" + 2].Merge = true;
                                    ws2.Cells["B" + 3 + ":E" + 3].Merge = true;
                                    ws2.Cells["B" + 4 + ":E" + 4].Merge = true;
                                    ws2.Cells["B" + 5 + ":E" + 5].Merge = true;
                                    xlsService.ExcelSetArialSize10(ws2, 1, 1, 10, 14);
                                    xlsService.ExcelSetFontBold(ws2, 1, 1, 5, 1);
                                    #endregion

                                    #region Tab 4 Document A1.2
                                    ws3.View.ZoomScale = 90;
                                    //set column width
                                    ws3.Column(1).Width = 16.43;
                                    ws3.Column(2).Width = 20;
                                    ws3.Column(3).Width = 20;
                                    ws3.Column(4).Width = 20;
                                    ws3.Column(5).Width = 20;

                                    //disable grid
                                    ws3.View.ShowGridLines = false;
                                    ws3.View.FreezePanes(8, 2);

                                    ws3.Cells[1, 1].Value = "Key Report:";
                                    ws3.Cells[2, 1].Value = "Control #:";
                                    ws3.Cells[3, 1].Value = "Provided By:";
                                    ws3.Cells[4, 1].Value = "Tested by:";
                                    ws3.Cells[5, 1].Value = "Description:";

                                    //populate field values
                                    ws3.Cells[1, 2].Value = filter.Filter.KeyReportName; //"Key Report:";
                                    
                                    ws3.Cells[2, 2].Value = sbControlId; //"Control #:";
                                    ws3.Cells[3, 2].Value = //"Provided By:";
                                    ws3.Cells[4, 2].Value = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("14. preparer")).Select(x => x.StrAnswer).FirstOrDefault(); //"Tested by:";
                                    ws3.Cells[5, 2].Value = //"Description:";
                                    
                                    ws3.Cells[5, 2].Value = $"(System) Screenshot of {filter.Filter.KeyReportName}";

                                    xlsService.ExcelSetBackgroundColorLightGray(ws3, 1, 1, 5, 1);

                                    xlsService.ExcelSetBorder(ws3, 1, 1, 5, 5);
                                    ws3.Cells["B" + 1 + ":E" + 1].Merge = true;
                                    ws3.Cells["B" + 2 + ":E" + 2].Merge = true;
                                    ws3.Cells["B" + 3 + ":E" + 3].Merge = true;
                                    ws3.Cells["B" + 4 + ":E" + 4].Merge = true;
                                    ws3.Cells["B" + 5 + ":E" + 5].Merge = true;
                                    xlsService.ExcelSetArialSize10(ws3, 1, 1, 10, 14);
                                    xlsService.ExcelSetFontBold(ws3, 1, 1, 5, 1);
                                    #endregion

                                    #region Tab 5 Document A2.1
                                    ws4.View.ZoomScale = 90;
                                    //set column width
                                    ws4.Column(1).Width = 16.43;
                                    ws4.Column(2).Width = 20;
                                    ws4.Column(3).Width = 20;
                                    ws4.Column(4).Width = 20;
                                    ws4.Column(5).Width = 20;

                                    //disable grid
                                    ws4.View.ShowGridLines = false;
                                    ws4.View.FreezePanes(8, 2);

                                    ws4.Cells[1, 1].Value = "Key Report:";
                                    ws4.Cells[2, 1].Value = "Control #:";
                                    ws4.Cells[3, 1].Value = "Provided By:";
                                    ws4.Cells[4, 1].Value = "Tested by:";
                                    ws4.Cells[5, 1].Value = "Description:";

                                    //populate field values
                                    ws4.Cells[1, 2].Value = filter.Filter.KeyReportName; //"Key Report:";
                                    
                                    ws4.Cells[2, 2].Value = sbControlId; //"Control #:";
                                    ws4.Cells[3, 2].Value = //"Provided By:";
                                    ws4.Cells[4, 2].Value = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("14. preparer")).Select(x => x.StrAnswer).FirstOrDefault(); //"Tested by:";
                                    ws4.Cells[5, 2].Value = //"Description:";
                                    
                                    ws4.Cells[5, 2].Value = $"(System) Screenshot of {filter.Filter.KeyReportName}";

                                    xlsService.ExcelSetBackgroundColorLightGray(ws4, 1, 1, 5, 1);

                                    xlsService.ExcelSetBorder(ws4, 1, 1, 5, 5);
                                    ws4.Cells["B" + 1 + ":E" + 1].Merge = true;
                                    ws4.Cells["B" + 2 + ":E" + 2].Merge = true;
                                    ws4.Cells["B" + 3 + ":E" + 3].Merge = true;
                                    ws4.Cells["B" + 4 + ":E" + 4].Merge = true;
                                    ws4.Cells["B" + 5 + ":E" + 5].Merge = true;
                                    xlsService.ExcelSetArialSize10(ws4, 1, 1, 10, 14);
                                    xlsService.ExcelSetFontBold(ws4, 1, 1, 5, 1);
                                    #endregion

                                    #region Tab 6 Document A2.2
                                    ws5.View.ZoomScale = 90;
                                    //set column width
                                    ws5.Column(1).Width = 16.43;
                                    ws5.Column(2).Width = 20;
                                    ws5.Column(3).Width = 20;
                                    ws5.Column(4).Width = 20;
                                    ws5.Column(5).Width = 20;

                                    //disable grid
                                    ws5.View.ShowGridLines = false;
                                    ws5.View.FreezePanes(8, 2);

                                    ws5.Cells[1, 1].Value = "Key Report:";
                                    ws5.Cells[2, 1].Value = "Control #:";
                                    ws5.Cells[3, 1].Value = "Provided By:";
                                    ws5.Cells[4, 1].Value = "Tested by:";
                                    ws5.Cells[5, 1].Value = "Description:";

                                    //populate field values
                                    ws5.Cells[1, 2].Value = filter.Filter.KeyReportName; //"Key Report:";
                                    
                                    ws5.Cells[2, 2].Value = sbControlId; //"Control #:";
                                    ws5.Cells[3, 2].Value = //"Provided By:";
                                    ws5.Cells[4, 2].Value = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("14. preparer")).Select(x => x.StrAnswer).FirstOrDefault(); //"Tested by:";
                                    ws5.Cells[5, 2].Value = //"Description:";
                                    
                                    ws5.Cells[5, 2].Value = $"(System) Screenshot of {filter.Filter.KeyReportName}";

                                    xlsService.ExcelSetBackgroundColorLightGray(ws5, 1, 1, 5, 1);

                                    xlsService.ExcelSetBorder(ws5, 1, 1, 5, 5);
                                    ws5.Cells["B" + 1 + ":E" + 1].Merge = true;
                                    ws5.Cells["B" + 2 + ":E" + 2].Merge = true;
                                    ws5.Cells["B" + 3 + ":E" + 3].Merge = true;
                                    ws5.Cells["B" + 4 + ":E" + 4].Merge = true;
                                    ws5.Cells["B" + 5 + ":E" + 5].Merge = true;
                                    xlsService.ExcelSetArialSize10(ws5, 1, 1, 10, 14);
                                    xlsService.ExcelSetFontBold(ws5, 1, 1, 5, 1);
                                    #endregion

                                    #region Tab 7 Document A3
                                    ws6.View.ZoomScale = 90;
                                    //set column width
                                    ws6.Column(1).Width = 16.43;
                                    ws6.Column(2).Width = 20;
                                    ws6.Column(3).Width = 20;
                                    ws6.Column(4).Width = 20;
                                    ws6.Column(5).Width = 20;

                                    //disable grid
                                    ws6.View.ShowGridLines = false;
                                    ws6.View.FreezePanes(8, 2);

                                    ws6.Cells[1, 1].Value = "Key Report:";
                                    ws6.Cells[2, 1].Value = "Control #:";
                                    ws6.Cells[3, 1].Value = "Provided By:";
                                    ws6.Cells[4, 1].Value = "Tested by:";
                                    ws6.Cells[5, 1].Value = "Description:";

                                    //populate field values
                                    ws6.Cells[1, 2].Value = filter.Filter.KeyReportName; //"Key Report:";
                                    
                                    ws6.Cells[2, 2].Value = sbControlId; //"Control #:";
                                    ws6.Cells[3, 2].Value = //"Provided By:";
                                    ws6.Cells[4, 2].Value = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("14. preparer")).Select(x => x.StrAnswer).FirstOrDefault(); //"Tested by:";
                                    ws6.Cells[5, 2].Value = //"Description:";
                                    
                                    ws6.Cells[5, 2].Value = $"(System) Screenshot of {filter.Filter.KeyReportName}";


                                    xlsService.ExcelSetBackgroundColorLightGray(ws6, 1, 1, 5, 1);

                                    xlsService.ExcelSetBorder(ws6, 1, 1, 5, 5);
                                    ws6.Cells["B" + 1 + ":E" + 1].Merge = true;
                                    ws6.Cells["B" + 2 + ":E" + 2].Merge = true;
                                    ws6.Cells["B" + 3 + ":E" + 3].Merge = true;
                                    ws6.Cells["B" + 4 + ":E" + 4].Merge = true;
                                    ws6.Cells["B" + 5 + ":E" + 5].Merge = true;
                                    xlsService.ExcelSetArialSize10(ws6, 1, 1, 10, 14);
                                    xlsService.ExcelSetFontBold(ws6, 1, 1, 5, 1);
                                    #endregion

                                    #region Tab 8 Document B1
                                    ws7.View.ZoomScale = 90;
                                    //set column width
                                    ws7.Column(1).Width = 16.43;
                                    ws7.Column(2).Width = 20;
                                    ws7.Column(3).Width = 20;
                                    ws7.Column(4).Width = 20;
                                    ws7.Column(5).Width = 20;

                                    //disable grid
                                    ws7.View.ShowGridLines = false;
                                    ws7.View.FreezePanes(8, 2);

                                    ws7.Cells[1, 1].Value = "Key Report:";
                                    ws7.Cells[2, 1].Value = "Control #:";
                                    ws7.Cells[3, 1].Value = "Provided By:";
                                    ws7.Cells[4, 1].Value = "Tested by:";
                                    ws7.Cells[5, 1].Value = "Description:";

                                    //populate field values
                                    ws7.Cells[1, 2].Value = filter.Filter.KeyReportName; //"Key Report:";
                                    
                                    ws7.Cells[2, 2].Value = sbControlId; //"Control #:";
                                    ws7.Cells[3, 2].Value = //"Provided By:";
                                    ws7.Cells[4, 2].Value = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("14. preparer")).Select(x => x.StrAnswer).FirstOrDefault(); //"Tested by:";
                                    ws7.Cells[5, 2].Value = //"Description:";
                                    
                                    ws7.Cells[5, 2].Value = $"(System) Screenshot of {filter.Filter.KeyReportName}";

                                    xlsService.ExcelSetBackgroundColorLightGray(ws7, 1, 1, 5, 1);

                                    xlsService.ExcelSetBorder(ws7, 1, 1, 5, 5);
                                    ws7.Cells["B" + 1 + ":E" + 1].Merge = true;
                                    ws7.Cells["B" + 2 + ":E" + 2].Merge = true;
                                    ws7.Cells["B" + 3 + ":E" + 3].Merge = true;
                                    ws7.Cells["B" + 4 + ":E" + 4].Merge = true;
                                    ws7.Cells["B" + 5 + ":E" + 5].Merge = true;
                                    xlsService.ExcelSetArialSize10(ws7, 1, 1, 10, 14);
                                    xlsService.ExcelSetFontBold(ws7, 1, 1, 5, 1);
                                    #endregion

                                    #region Tab 9 Document C1
                                    ws8.View.ZoomScale = 90;
                                    //set column width
                                    ws8.Column(1).Width = 16.43;
                                    ws8.Column(2).Width = 20;
                                    ws8.Column(3).Width = 20;
                                    ws8.Column(4).Width = 20;
                                    ws8.Column(5).Width = 20;

                                    //disable grid
                                    ws8.View.ShowGridLines = false;
                                    ws8.View.FreezePanes(8, 2);

                                    ws8.Cells[1, 1].Value = "Key Report:";
                                    ws8.Cells[2, 1].Value = "Control #:";
                                    ws8.Cells[3, 1].Value = "Provided By:";
                                    ws8.Cells[4, 1].Value = "Tested by:";
                                    ws8.Cells[5, 1].Value = "Description:";

                                    //populate field values
                                    ws8.Cells[1, 2].Value = filter.Filter.KeyReportName; //"Key Report:";
                                    
                                    ws8.Cells[2, 2].Value = sbControlId; //"Control #:";
                                    ws8.Cells[3, 2].Value = //"Provided By:";
                                    ws8.Cells[4, 2].Value = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("14. preparer")).Select(x => x.StrAnswer).FirstOrDefault(); //"Tested by:";
                                    ws8.Cells[5, 2].Value = //"Description:";
                                    
                                    ws8.Cells[5, 2].Value = $"(System) Screenshot of {filter.Filter.KeyReportName}";

                                    xlsService.ExcelSetBackgroundColorLightGray(ws8, 1, 1, 5, 1);

                                    xlsService.ExcelSetBorder(ws8, 1, 1, 5, 5);
                                    ws8.Cells["B" + 1 + ":E" + 1].Merge = true;
                                    ws8.Cells["B" + 2 + ":E" + 2].Merge = true;
                                    ws8.Cells["B" + 3 + ":E" + 3].Merge = true;
                                    ws8.Cells["B" + 4 + ":E" + 4].Merge = true;
                                    ws8.Cells["B" + 5 + ":E" + 5].Merge = true;
                                    xlsService.ExcelSetArialSize10(ws8, 1, 1, 10, 14);
                                    xlsService.ExcelSetFontBold(ws8, 1, 1, 5, 1);
                                    #endregion

                                    #region Tab 10 Document C2
                                    ws9.View.ZoomScale = 90;
                                    //set column width
                                    ws9.Column(1).Width = 16.43;
                                    ws9.Column(2).Width = 20;
                                    ws9.Column(3).Width = 20;
                                    ws9.Column(4).Width = 20;
                                    ws9.Column(5).Width = 20;

                                    //disable grid
                                    ws9.View.ShowGridLines = false;
                                    ws9.View.FreezePanes(8, 2);

                                    ws9.Cells[1, 1].Value = "Key Report:";
                                    ws9.Cells[2, 1].Value = "Control #:";
                                    ws9.Cells[3, 1].Value = "Provided By:";
                                    ws9.Cells[4, 1].Value = "Tested by:";
                                    ws9.Cells[5, 1].Value = "Description:";

                                    //populate field values
                                    ws9.Cells[1, 2].Value = filter.Filter.KeyReportName; //"Key Report:";
                                    
                                    ws9.Cells[2, 2].Value = sbControlId; //"Control #:";
                                    ws9.Cells[3, 2].Value = //"Provided By:";
                                    ws9.Cells[4, 2].Value = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("14. preparer")).Select(x => x.StrAnswer).FirstOrDefault(); //"Tested by:";
                                    ws9.Cells[5, 2].Value = //"Description:";
                                   
                                    ws9.Cells[5, 2].Value = $"(System) Screenshot of {filter.Filter.KeyReportName}";

                                    xlsService.ExcelSetBackgroundColorLightGray(ws9, 1, 1, 5, 1);

                                    xlsService.ExcelSetBorder(ws9, 1, 1, 5, 5);
                                    ws9.Cells["B" + 1 + ":E" + 1].Merge = true;
                                    ws9.Cells["B" + 2 + ":E" + 2].Merge = true;
                                    ws9.Cells["B" + 3 + ":E" + 3].Merge = true;
                                    ws9.Cells["B" + 4 + ":E" + 4].Merge = true;
                                    ws9.Cells["B" + 5 + ":E" + 5].Merge = true;
                                    xlsService.ExcelSetArialSize10(ws9, 1, 1, 10, 14);
                                    xlsService.ExcelSetFontBold(ws9, 1, 1, 5, 1);
                                    #endregion

                                    #region Tab 11 Document D1
                                    ws10.View.ZoomScale = 90;
                                    //set column width
                                    ws10.Column(1).Width = 16.43;
                                    ws10.Column(2).Width = 20;
                                    ws10.Column(3).Width = 20;
                                    ws10.Column(4).Width = 20;
                                    ws10.Column(5).Width = 20;

                                    //disable grid
                                    ws10.View.ShowGridLines = false;
                                    ws10.View.FreezePanes(8, 2);

                                    ws10.Cells[1, 1].Value = "Key Report:";
                                    ws10.Cells[2, 1].Value = "Control #:";
                                    ws10.Cells[3, 1].Value = "Provided By:";
                                    ws10.Cells[4, 1].Value = "Tested by:";
                                    ws10.Cells[5, 1].Value = "Description:";

                                    //populate field values
                                    ws10.Cells[1, 2].Value = filter.Filter.KeyReportName; //"Key Report:";
                                    
                                    ws10.Cells[2, 2].Value = sbControlId; //"Control #:";
                                    ws10.Cells[3, 2].Value = //"Provided By:";
                                    ws10.Cells[4, 2].Value = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("14. preparer")).Select(x => x.StrAnswer).FirstOrDefault(); //"Tested by:";
                                    ws10.Cells[5, 2].Value = //"Description:";
                                    
                                    ws10.Cells[5, 2].Value = $"(System) Screenshot of {filter.Filter.KeyReportName}";

                                    xlsService.ExcelSetBackgroundColorLightGray(ws10, 1, 1, 5, 1);

                                    xlsService.ExcelSetBorder(ws10, 1, 1, 5, 5);
                                    ws10.Cells["B" + 1 + ":E" + 1].Merge = true;
                                    ws10.Cells["B" + 2 + ":E" + 2].Merge = true;
                                    ws10.Cells["B" + 3 + ":E" + 3].Merge = true;
                                    ws10.Cells["B" + 4 + ":E" + 4].Merge = true;
                                    ws10.Cells["B" + 5 + ":E" + 5].Merge = true;
                                    xlsService.ExcelSetArialSize10(ws10, 1, 1, 10, 14);
                                    xlsService.ExcelSetFontBold(ws10, 1, 1, 5, 1);
                                    #endregion
                                }
                                #endregion excel_pwc_keyreport
                                //==================================================
                                xls.Workbook.Worksheets.First().Select();

                                //save file
                                string startupPath = Environment.CurrentDirectory;
                                
                                string strSourceDownload = Path.Combine(startupPath, "include", "upload", "keyreport", "pwc"); //062521

                                if (!Directory.Exists(strSourceDownload))
                                {
                                    Directory.CreateDirectory(strSourceDownload);
                                }
                                var ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                               
                                string filename = $"KeyReport_PWC-{filter.Filter.KeyReportName}-{filter.Filter.ClientName}-{ts}.xlsx"; //062521
                                string strOutput = Path.Combine(strSourceDownload, filename);

                                //Check if file not exists
                                if (System.IO.File.Exists(strOutput))
                                {
                                    System.IO.File.Delete(strOutput);
                                }

                                xls.SaveAs(new FileInfo(strOutput));
                                excelFilename = filename;

                                //return status "Ok" with filename
                                return Ok(excelFilename);

                            }

                        }


                    }

                    return NoContent();

                }
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GenerateKeyReportTemplateFilePWC {ex}", "ErrorGenerateKeyReportTemplateFilePWC");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GenerateKeyReportTemplateFilePWC");
                if (_environment.IsDevelopment())
                {
                    return BadRequest(ex.ToString());
                }
                else
                {
                    return BadRequest();
                }
            }
            return Ok(excelFilename);
        }
        

        //[AllowAnonymous]
        [HttpPost("tracker/generate")]
        public IActionResult GenerateKeyReportTrackerTemplateFile([FromBody] KeyReportFilter filter)
        {
            string excelFilename = string.Empty;

            try
            {
                List<int> itemId = new List<int>();
                string appIdConsolOrig = _config.GetSection("KeyReportApp").GetSection("ConsolOrigFormatId").Value;
                string appIdAllIUC = _config.GetSection("KeyReportApp").GetSection("AllIUCId").Value;
                string appIdTestStatus = _config.GetSection("KeyReportApp").GetSection("TestStatusTrackerId").Value;
                string appIdException = _config.GetSection("KeyReportApp").GetSection("ExceptionId").Value;

                //get client name
                var checkClient = _soxContext.KeyReportUserInput
                    .Where(x => x.TagClientName.ToLower().Equals(filter.ClientName.ToLower()))
                    .AsNoTracking()
                    .Select(x => x.ItemId)
                    .Distinct()
                    .ToList();
                if (checkClient != null)
                {
                    foreach (var Id in checkClient)
                    {
                        var checkFY = _soxContext.KeyReportUserInput
                            .Where(x =>
                                x.TagFY.ToLower().Equals(filter.FY.ToLower())
                                && x.ItemId.Equals(Id)
                                && x.TagStatus.ToLower() != "inactive"
                                &&
                                    (
                                        x.AppId.Equals(appIdConsolOrig)
                                        || x.AppId.Equals(appIdAllIUC)
                                        || x.AppId.Equals(appIdTestStatus)
                                        || x.AppId.Equals(appIdException)
                                    )
                                )
                            .AsNoTracking()
                            .Select(x => x.ItemId)
                            .Distinct()
                            .ToList();
                        if (checkFY != null)
                        {
                            foreach (var item in checkFY)
                            {
                                if (!itemId.Contains(item))
                                    itemId.Add(item);
                            }

                        }
                    }

                }




                Debug.WriteLine($"ItemID Count: {itemId.Count}");


                #region Create Excel
                //Creating excel 
                //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage xls = new ExcelPackage())
                {
                    ExcelService xlsService = new ExcelService();
                    //set sheet name
                    var ws = xls.Workbook.Worksheets.Add("Navigation");
                    var ws1 = xls.Workbook.Worksheets.Add("1 Consol Original FormatIUC");
                    var ws2 = xls.Workbook.Worksheets.Add("2 All IUC- use this");
                    var ws3 = xls.Workbook.Worksheets.Add("3 Test Status Tracker");
                    var ws4 = xls.Workbook.Worksheets.Add("4 List of Exceptions");
                    var ws5 = xls.Workbook.Worksheets.Add("5 Status Summary");
                    var ws6 = xls.Workbook.Worksheets.Add("Drop down");



                    //Navigation
                    ws.View.ZoomScale = 85;

                    //disable grid
                    ws.View.ShowGridLines = false;
                    ws.Column(1).Width = 26.86;

                    ws.Cells[1, 1].Value = "Navigation";
                    ws.Cells[2, 1].Value = "1 Consol Original FormatIUC";
                    ws.Cells[3, 1].Value = "2 All IUC- use this";
                    ws.Cells[4, 1].Value = "3 Test Status Tracker";
                    ws.Cells[5, 1].Value = "4 List of Exceptions";
                    ws.Cells[6, 1].Value = "5 Status Summary";
                    ws.Cells[7, 1].Value = "Drop down";

                    xlsService.ExcelSetArialSize10(ws, 1, 1, 7, 1);
                    xlsService.ExcelSetFontBold(ws, 1, 1, 1, 1);

                    //freeze pane


                    //end Navigation




                    //1 Consol Original FormatIUC
                    ws1.View.ZoomScale = 85;
                    //set column width
                    ws1.Column(1).Width = 4;        //No.
                    ws1.Column(2).Width = 12.57;    //Key Control    
                    ws1.Column(3).Width = 55;       //Control Activity
                    ws1.Column(4).Width = 12.57;    //Key / Non-Key Control
                    ws1.Column(5).Width = 20;       //Name of IUC
                    ws1.Column(6).Width = 12.57;    // Source Process
                    ws1.Column(7).Width = 12.57;    //Key Report
                    ws1.Column(8).Width = 12.57;    //IUC Type
                    ws1.Column(9).Width = 8;        //System/Source
                    ws1.Column(10).Width = 16.71;   //Report Customized
                    ws1.Column(11).Width = 16.71;   //Controls Relying on IUC
                    ws1.Column(12).Width = 25;      //Preparer
                    ws1.Column(13).Width = 26.14;      //Reviewer
                    ws1.Column(14).Width = 42.29;   //Notes

                    //disable grid
                    ws1.View.ShowGridLines = false;

                    //set row
                    int row = 1;

                    //set title
                    ws1.Cells[row, 1].Value = "Back to Navigation";
                    row++;
                    ws1.Cells[row, 1].Value = filter.ClientName;
                    row++;
                    ws1.Cells[row, 1].Value = "Consolidated IUC List";
                    row++;
                    ws1.Cells[row, 1].Value = "SOX FY " + filter.FY;
                    xlsService.ExcelSetFontBold(ws1, 1, 1, row, 1); //(worksheet, from row, from column, to row, to column)
                    row += 3;

                    ws1.Cells[3, 10].Value = "Number of IUC (include duplicates) ";
                    ws1.Cells[4, 10].Value = "Total IUC (include duplicates)";
                    ws1.Cells[2, 12].Value = "Source in ELC";
                    ws1.Cells[2, 13].Value = "Source in Other Processes";

                    //set table header
                    ws1.Cells[row, 1].Value = "No.";
                    ws1.Cells[row, 2].Value = "Key Control";
                    ws1.Cells[row, 3].Value = "Control Activity";
                    ws1.Cells[row, 4].Value = "Key / Non-Key Control";
                    ws1.Cells[row, 5].Value = "Name of IUC";
                    ws1.Cells[row, 6].Value = "Source Process";
                    ws1.Cells[row, 7].Value = "Key Report";
                    ws1.Cells[row, 8].Value = "IUC Type";
                    ws1.Cells[row, 9].Value = "System/Source";
                    ws1.Cells[row, 10].Value = "Report Customized";
                    ws1.Cells[row, 11].Value = "Controls Relying on IUC";
                    ws1.Cells[row, 12].Value = "Preparer";
                    ws1.Cells[row, 13].Value = "Reviewer";
                    ws1.Cells[row, 14].Value = "Notes ";

                    //format cell
                    xlsService.ExcelWrapText(ws1, row, 1, row, 14); //(worksheet, from row, from column, to row, to column)
                    xlsService.ExcelSetHorizontalAlignCenter(ws1, 7, 1, 7, 14); //(worksheet, from row, from column, to row, to column)
                    xlsService.ExcelSetHorizontalAlignCenter(ws1, 7, 1, 7, 14); //(worksheet, from row, from column, to row, to column)
                    xlsService.ExcelSetVerticalAlignCenter(ws1, 7, 1, 7, 14); //(worksheet, from row, from column, to row, to column)
                    xlsService.ExcelSetHorizontalAlignCenter(ws1, 2, 12, 4, 13);

                    //font color and font background
                    //xlsService.ExcelSetFontColorRed(ws1, 6, 1, 6, 14);
                    xlsService.ExcelSetBackgroundColorGray(ws1, 7, 1, 7, 14);
                    xlsService.ExcelSetBackgroundColorGray(ws1, 3, 10, 4, 11);
                    xlsService.ExcelSetBackgroundColorGray(ws1, 2, 12, 2, 13);
                    xlsService.ExcelSetArialSize10(ws1, 1, 1, row, 14);
                    xlsService.ExcelSetFontBold(ws1, 2, 1, 7, 14);

                    //set border
                    xlsService.ExcelSetBorder(ws1, row, 1, row, 14);
                    xlsService.ExcelSetBorder(ws1, 2, 12, 4, 13);
                    xlsService.ExcelSetBorder(ws1, 3, 10, 4, 11);

                    //set auto filter
                    ws1.Cells[row, 1, row, 14].AutoFilter = true;

                    int num = 1;
                    foreach (var Id in itemId)
                    {
                        var checkConsolFormat = _soxContext.KeyReportUserInput.Where(x =>
                                x.ItemId.Equals(Id)
                                && x.AppId.Equals(appIdConsolOrig)
                                && (
                                    x.StrQuestion.ToLower().Contains("key control id")
                                    || x.StrQuestion.ToLower().Contains("control activity")
                                    || x.StrQuestion.ToLower().Contains("key/non-key control")
                                    || x.StrQuestion.ToLower().Contains("name of key report/iuc")
                                    || x.StrQuestion.ToLower().Contains("source process")
                                    || x.StrQuestion.ToLower().Contains("key/non-key report")
                                    || x.StrQuestion.ToLower().Contains("iuc type")
                                    || x.StrQuestion.ToLower().Contains("system / source")
                                    || x.StrQuestion.ToLower().Contains("report customized")
                                    || x.StrQuestion.ToLower().Contains("controls relying on key report")
                                    || x.StrQuestion.ToLower().Contains("preparer")
                                    || x.StrQuestion.ToLower().Contains("reviewer")
                                    || x.StrQuestion.ToLower().Contains("report notes")
                                   )
                                )
                            .AsNoTracking()
                            .ToList();

                        if(checkConsolFormat != null && checkConsolFormat.Count > 0)
                        {
                            row++;
                            ws1.Cells[row, 1].Value = num.ToString(); //"No.";
                            ws1.Cells[row, 2].Value = checkConsolFormat.Where(x => x.StrQuestion.ToLower().Contains("key control id")).Select(x => x.StrAnswer).FirstOrDefault(); //"Key Control";
                            ws1.Cells[row, 3].Value = checkConsolFormat.Where(x => x.StrQuestion.ToLower().Contains("control activity")).Select(x => x.StrAnswer).FirstOrDefault(); //"Control Activity";
                            ws1.Cells[row, 4].Value = checkConsolFormat.Where(x => x.StrQuestion.ToLower().Contains("key/non-key control")).Select(x => x.StrAnswer).FirstOrDefault(); //"Key / Non-Key Control";
                            ws1.Cells[row, 5].Value = checkConsolFormat.Where(x => x.StrQuestion.ToLower().Contains("name of key report/iuc")).Select(x => x.StrAnswer).FirstOrDefault(); //"Name of IUC";
                            ws1.Cells[row, 6].Value = checkConsolFormat.Where(x => x.StrQuestion.ToLower().Contains("source process")).Select(x => x.StrAnswer).FirstOrDefault(); //"Source Process";
                            ws1.Cells[row, 7].Value = checkConsolFormat.Where(x => x.StrQuestion.ToLower().Contains("key/non-key report")).Select(x => x.StrAnswer).FirstOrDefault(); //"Key Report";
                            ws1.Cells[row, 8].Value = checkConsolFormat.Where(x => x.StrQuestion.ToLower().Contains("iuc type")).Select(x => x.StrAnswer).FirstOrDefault(); //"IUC Type";
                            ws1.Cells[row, 9].Value = checkConsolFormat.Where(x => x.StrQuestion.ToLower().Contains("system / source")).Select(x => x.StrAnswer).FirstOrDefault(); //"System/Source";
                            ws1.Cells[row, 10].Value = checkConsolFormat.Where(x => x.StrQuestion.ToLower().Contains("report customized")).Select(x => x.StrAnswer).FirstOrDefault(); //"Report Customized";
                            ws1.Cells[row, 11].Value = checkConsolFormat.Where(x => x.StrQuestion.ToLower().Contains("controls relying on key report")).Select(x => x.StrAnswer).FirstOrDefault(); //"Controls Relying on IUC";
                            ws1.Cells[row, 12].Value = checkConsolFormat.Where(x => x.StrQuestion.ToLower().Contains("preparer")).Select(x => x.StrAnswer).FirstOrDefault(); //"Preparer";
                            ws1.Cells[row, 13].Value = checkConsolFormat.Where(x => x.StrQuestion.ToLower().Contains("reviewer")).Select(x => x.StrAnswer).FirstOrDefault();//"Reviewer";
                            ws1.Cells[row, 14].Value = checkConsolFormat.Where(x => x.StrQuestion.ToLower().Contains("report notes")).Select(x => x.StrAnswer).FirstOrDefault(); //"Notes ";
                            num++;
                        }

                    }

                    

                    //end 1 Consol Original FormatIUC 

                    //2 All IUC- use this
                    ws2.View.ZoomScale = 85;
                    //set column width
                    ws2.Column(1).Width = 4;        //No.
                    ws2.Column(2).Width = 12.57;    //Key Control    
                    ws2.Column(3).Width = 63.57;    //Control Activity
                    ws2.Column(4).Width = 15.71;    //Key / Non-Key Control
                    ws2.Column(5).Width = 23.86;    // Name of IUC
                    ws2.Column(6).Width = 12.57;    // Source Process
                    ws2.Column(7).Width = 12.57;    //Key Report
                    ws2.Column(8).Width = 12.57;    //Unique Key Report
                    ws2.Column(9).Width = 12.57;    //IUC Type
                    ws2.Column(10).Width = 12.57;   //System/Source
                    ws2.Column(11).Width = 14.57;   //Report Customized
                    ws2.Column(12).Width = 12.57;   //Controls Relying on IUC
                    ws2.Column(13).Width = 19.86;   //Preparer
                    ws2.Column(14).Width = 19.86;   //Reviewer
                    ws2.Column(15).Width = 20.57;   //Added to Key Report Tracker
                    ws2.Column(16).Width = 20.57;   //Report Notes
                    ws2.Column(17).Width = 33.29;   //Description of Key Report (key data fields used, purpose of report)
                    ws2.Column(18).Width = 16;      //Key Report Type (Canned, Customizable, Query, Excel, Custom Query, Saved Searches)
                    ws2.Column(19).Width = 16;      //How is the report generated? 
                    ws2.Column(20).Width = 35.86;   //How is the report used to support the control(s)?
                    ws2.Column(21).Width = 31;      //What steps are performed to validate the accuracy of the report?
                    ws2.Column(22).Width = 39.86;   //What steps are performed to validate the completeness of the report?
                    ws2.Column(23).Width = 16;      //What steps are performed to validate the source data used by the report?
                    ws2.Column(24).Width = 16;      //Are parameters (e.g. date ranges) input each time this report is run? 
                    ws2.Column(25).Width = 16;      //Who is authorized to make/request  changes to this report?
                    ws2.Column(26).Width = 16;      //Effective date
                    ws2.Column(27).Width = 15.57;   //Who has access to edit/modify this report? 
                    ws2.Column(28).Width = 18.29;   //Who has access to run the report? Restricted report access? 
                    ws2.Column(29).Width = 16.57;   //When was this report last modified? 
                    ws2.Column(30).Width = 15.57;   //How was it tested when last modified?
                    ws2.Column(31).Width = 15.57;   //IT Report Owner
                    ws2.Column(32).Width = 20.29;   //Questions
                    ws2.Column(33).Width = 14;      //Fastly Notes and Questions
                    ws2.Column(34).Width = 14;      //Meeting date
                    ws2.Column(35).Width = 13.71;   //Process
                    //disable grid
                    ws2.View.ShowGridLines = false;

                    ws2.Cells["D" + 2 + ":E" + 2].Merge = true;
                    ws2.Cells["D" + 3 + ":E" + 3].Merge = true;
                    ws2.Cells["D" + 4 + ":E" + 4].Merge = true;
                    ws2.Cells["Q" + 6 + ":Y" + 6].Merge = true;
                    ws2.Cells["AA" + 6 + ":AD" + 6].Merge = true;

                    //set row
                    int row2 = 1;

                    //set title
                    ws2.Cells[row2, 1].Value = "Back to Navigation";
                    row2++;
                    ws2.Cells[row2, 1].Value = filter.ClientName;
                    row2++;
                    ws2.Cells[row2, 1].Value = "Consolidated IUC List";
                    row2++;
                    ws2.Cells[row2, 1].Value = "SOX FY " + filter.FY;
                    xlsService.ExcelSetFontBold(ws2, 1, 1, row2, 1); //(worksheet, from row, from column, to row, to column)
                    row2 += 3;

                    ws2.Cells[2, 4].Value = "Number of IUC (include duplicates)";
                    ws2.Cells[3, 4].Value = "# of Key Reports (include dups)";
                    ws2.Cells[4, 4].Value = "# of UNIQUE Key Reports";
                    ws2.Cells[6, 17].Value = "Business Owners";
                    ws2.Cells[6, 27].Value = "Information Technology";


                    //set table header
                    ws2.Cells[row2, 1].Value = "No.";
                    ws2.Cells[row2, 2].Value = "Key Control";
                    ws2.Cells[row2, 3].Value = "Control Activity";
                    ws2.Cells[row2, 4].Value = "Key / Non-Key Control";
                    ws2.Cells[row2, 5].Value = "Name of IUC";
                    ws2.Cells[row2, 6].Value = "Source Process";
                    ws2.Cells[row2, 7].Value = "Key Report";
                    ws2.Cells[row2, 8].Value = "Unique Key Report";
                    ws2.Cells[row2, 9].Value = "IUC Type";
                    ws2.Cells[row2, 10].Value = "System/Source";
                    ws2.Cells[row2, 11].Value = "Report Customized";
                    ws2.Cells[row2, 12].Value = "Controls Relying on IUC";
                    ws2.Cells[row2, 13].Value = "Preparer";
                    ws2.Cells[row2, 14].Value = "Reviewer";
                    ws2.Cells[row2, 15].Value = "Added to Key Report Tracker";
                    ws2.Cells[row2, 16].Value = "Report Notes";
                    ws2.Cells[row2, 17].Value = "Description of Key Report (key data fields used, purpose of report)";
                    ws2.Cells[row2, 18].Value = "Key Report Type (Canned, Customizable, Query, Excel, Custom Query, Saved Searches)";
                    ws2.Cells[row2, 19].Value = "How is the report generated?";
                    ws2.Cells[row2, 20].Value = "How is the report used to support the control(s)?";
                    ws2.Cells[row2, 21].Value = "What steps are performed to validate the accuracy of the report?";
                    ws2.Cells[row2, 22].Value = "What steps are performed to validate the completeness of the report?";
                    ws2.Cells[row2, 23].Value = "What steps are performed to validate the source data used by the report?";
                    ws2.Cells[row2, 24].Value = "Are parameters (e.g. date ranges) input each time this report is run?";
                    ws2.Cells[row2, 25].Value = "Who is authorized to make/request  changes to this report?";
                    ws2.Cells[row2, 26].Value = "Effective date";
                    ws2.Cells[row2, 27].Value = "Who has access to edit/modify this report?";
                    ws2.Cells[row2, 28].Value = "Who has access to run the report? Restricted report access?";
                    ws2.Cells[row2, 29].Value = "When was this report last modified?";
                    ws2.Cells[row2, 30].Value = "How was it tested when last modified?";
                    ws2.Cells[row2, 31].Value = "IT Report Owner";
                    ws2.Cells[row2, 32].Value = "Questions";
                    ws2.Cells[row2, 33].Value = "Fastly Notes and Questions";
                    ws2.Cells[row2, 34].Value = "Meeting date";
                    ws2.Cells[row2, 35].Value = "Process";



                    //format cell
                    xlsService.ExcelWrapText(ws2, row2, 1, row2, 35); //(worksheet, from row, from column, to row, to column)
                    xlsService.ExcelSetHorizontalAlignCenter(ws2, 7, 1, 7, 35); //(worksheet, from row, from column, to row, to column)
                    xlsService.ExcelSetHorizontalAlignCenter(ws2, 7, 1, 7, 35); //(worksheet, from row, from column, to row, to column)
                    xlsService.ExcelSetVerticalAlignCenter(ws2, 7, 1, 7, 35); //(worksheet, from row, from column, to row, to column)
                    xlsService.ExcelSetHorizontalAlignCenter(ws2, 6, 17, 6, 17);
                    xlsService.ExcelSetHorizontalAlignCenter(ws2, 6, 27, 6, 27);

                    //font color and font background

                    xlsService.ExcelSetBackgroundColorGray(ws2, 2, 4, 4, 5);
                    xlsService.ExcelSetBackgroundColorGray(ws2, 7, 1, 7, 16);
                    xlsService.ExcelSetBackgroundColorAmber(ws2, 6, 17, 6, 26);
                    xlsService.ExcelSetBackgroundColorAmber(ws2, 7, 17, 7, 26);
                    xlsService.ExcelSetBackgroundColorLightBlue1(ws2, 6, 27, 6, 31);
                    xlsService.ExcelSetBackgroundColorLightBlue1(ws2, 7, 27, 7, 31);
                    xlsService.ExcelSetBackgroundColorLightGreen(ws2, 7, 32, 7, 32);
                    xlsService.ExcelSetBackgroundColorYellow(ws2, 7, 33, 7, 33);
                    xlsService.ExcelSetBackgroundColorLightGreen(ws2, 7, 34, 7, 34);
                    xlsService.ExcelSetBackgroundColorLightBlue1(ws2, 7, 35, 7, 35);

                    xlsService.ExcelSetArialSize10(ws2, 1, 1, row2, 35);
                    xlsService.ExcelSetFontBold(ws2, 2, 1, 7, 35);

                    //set border
                    xlsService.ExcelSetBorder(ws2, 6, 17, 6, 31);
                    xlsService.ExcelSetBorder(ws2, row2, 1, row2, 35);
                    xlsService.ExcelSetBorder(ws2, 2, 4, 4, 6);


                    //set auto filter
                    ws2.Cells[row2, 1, row2, 35].AutoFilter = true;


                    num = 1;
                    foreach (var Id in itemId)
                    {
                        var checkAllIUC = _soxContext.KeyReportUserInput.Where(x =>
                                x.ItemId.Equals(Id)
                                && x.AppId.Equals(appIdAllIUC))
                            .AsNoTracking()
                            .ToList();

                        if (checkAllIUC != null && checkAllIUC.Count > 0)
                        {
                            row2++;
                            
                            ws2.Cells[row2, 1].Value = num.ToString(); // "No.";
                            ws2.Cells[row2, 2].Value = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("key control id")).Select(x => x.StrAnswer).FirstOrDefault(); // "Key Control";
                            ws2.Cells[row2, 3].Value = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("control activity")).Select(x => x.StrAnswer).FirstOrDefault(); // "Control Activity";
                            ws2.Cells[row2, 4].Value = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("key/non-key control")).Select(x => x.StrAnswer).FirstOrDefault(); // "Key / Non-Key Control";
                            ws2.Cells[row2, 5].Value = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("name of key report/iuc")).Select(x => x.StrAnswer).FirstOrDefault(); // "Name of IUC";
                            ws2.Cells[row2, 6].Value = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("source process")).Select(x => x.StrAnswer).FirstOrDefault(); // "Source Process";
                            ws2.Cells[row2, 7].Value = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("key/non-key report")).Select(x => x.StrAnswer).FirstOrDefault(); // "Key Report";
                            ws2.Cells[row2, 8].Value = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("unique key report")).Select(x => x.StrAnswer).FirstOrDefault(); // "Unique Key Report";
                            ws2.Cells[row2, 9].Value = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("iuc type")).Select(x => x.StrAnswer).FirstOrDefault(); // "IUC Type";
                            ws2.Cells[row2, 10].Value = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("system / source")).Select(x => x.StrAnswer).FirstOrDefault(); // "System/Source";
                            ws2.Cells[row2, 11].Value = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("report customized")).Select(x => x.StrAnswer).FirstOrDefault(); // "Report Customized";
                            ws2.Cells[row2, 12].Value = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("controls relying on key report")).Select(x => x.StrAnswer).FirstOrDefault(); // "Controls Relying on IUC";
                            ws2.Cells[row2, 13].Value = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("preparer")).Select(x => x.StrAnswer).FirstOrDefault(); // "Preparer";
                            ws2.Cells[row2, 14].Value = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("16. reviewer")).Select(x => x.StrAnswer).FirstOrDefault(); // "Reviewer";
                            ws2.Cells[row2, 15].Value = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("added to key report tracker")).Select(x => x.StrAnswer).FirstOrDefault(); // "Added to Key Report Tracker";
                            ws2.Cells[row2, 16].Value = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("report notes")).Select(x => x.StrAnswer).FirstOrDefault(); // "Report Notes";
                            ws2.Cells[row2, 17].Value = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("description of key report (key data fields used, purpose of report)")).Select(x => x.StrAnswer).FirstOrDefault(); // "Description of Key Report (key data fields used, purpose of report)";
                            ws2.Cells[row2, 18].Value = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("key report type")).Select(x => x.StrAnswer).FirstOrDefault(); // "Key Report Type (Canned, Customizable, Query, Excel, Custom Query, Saved Searches)";
                            ws2.Cells[row2, 19].Value = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("how is the report generated?")).Select(x => x.StrAnswer).FirstOrDefault(); // "How is the report generated?";
                            ws2.Cells[row2, 20].Value = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("how is the report used to support the control(s)")).Select(x => x.StrAnswer).FirstOrDefault(); // "How is the report used to support the control(s)?";
                            ws2.Cells[row2, 21].Value = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("what steps are performed to validate the accuracy of the report")).Select(x => x.StrAnswer).FirstOrDefault(); ; // "What steps are performed to validate the accuracy of the report?";
                            ws2.Cells[row2, 22].Value = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("what steps are performed to validate the completeness of the report")).Select(x => x.StrAnswer).FirstOrDefault(); // "What steps are performed to validate the completeness of the report?";
                            ws2.Cells[row2, 23].Value = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("what steps are performed to validate the source data used by the report")).Select(x => x.StrAnswer).FirstOrDefault(); // "What steps are performed to validate the source data used by the report?";
                            ws2.Cells[row2, 24].Value = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("are parameters (e.g. date ranges) input each time this report is run")).Select(x => x.StrAnswer).FirstOrDefault(); // "Are parameters (e.g. date ranges) input each time this report is run?";
                            ws2.Cells[row2, 25].Value = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("who is authorized to make/request changes to this report")).Select(x => x.StrAnswer).FirstOrDefault(); // "Who is authorized to make/request  changes to this report?";
                            ws2.Cells[row2, 26].Value = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("effective date")).Select(x => x.StrAnswer).FirstOrDefault(); // "Effective date";
                            ws2.Cells[row2, 27].Value = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("who has access to edit/modify this report")).Select(x => x.StrAnswer).FirstOrDefault(); // "Who has access to edit/modify this report?";
                            ws2.Cells[row2, 28].Value = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("who has access to run the report? restricted report access")).Select(x => x.StrAnswer).FirstOrDefault(); // "Who has access to run the report? Restricted report access?";
                            ws2.Cells[row2, 29].Value = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("when was this report last modified")).Select(x => x.StrAnswer).FirstOrDefault(); // "When was this report last modified?";
                            ws2.Cells[row2, 30].Value = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("how was it tested when last modified")).Select(x => x.StrAnswer).FirstOrDefault(); // "How was it tested when last modified?";
                            ws2.Cells[row2, 31].Value = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("it report owner")).Select(x => x.StrAnswer).FirstOrDefault(); // "IT Report Owner";
                            ws2.Cells[row2, 32].Value = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("36. questions")).Select(x => x.StrAnswer).FirstOrDefault(); // "Questions";
                            ws2.Cells[row2, 33].Value = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("notes and questions")).Select(x => x.StrAnswer).FirstOrDefault(); // "Fastly Notes and Questions";
                            ws2.Cells[row2, 34].Value = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("meeting date")).Select(x => x.StrAnswer).FirstOrDefault(); // "Meeting date";
                            ws2.Cells[row2, 35].Value = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("40. process")).Select(x => x.StrAnswer).FirstOrDefault(); // "Process";
                            num++;
                        }

                    }


                    

                    //end 2 All IUC- use this

                    //3 Test Status Tracker
                    ws3.View.ZoomScale = 85;
                    //set column width
                    ws3.Column(1).Width = 5.71;     //No.
                    ws3.Column(2).Width = 34.71;    //Name of IUC
                    ws3.Column(3).Width = 34.71;    //Process
                    ws3.Column(4).Width = 12.57;    //Key Control
                    ws3.Column(5).Width = 49.71;    // Control Activity
                    ws3.Column(6).Width = 14;       // Key / Non-Key Control
                    ws3.Column(7).Width = 13.57;    //Key Report
                    ws3.Column(8).Width = 13.57;    //Unique Key Report
                    ws3.Column(9).Width = 13.57;    //Preparer
                    ws3.Column(10).Width = 13.57;   //Process Owner
                    ws3.Column(11).Width = 12;      //Key Report Owner
                    ws3.Column(12).Width = 12;      //Key Report IT Owner
                    ws3.Column(13).Width = 13.71;   //Set up report Lead Sheet for testing
                    ws3.Column(14).Width = 13.71;   //Schedule Process Owner meeting
                    ws3.Column(15).Width = 19;      //Report Received
                    ws3.Column(16).Width = 13.71;   //PBC Status
                    ws3.Column(17).Width = 12;      //Tester
                    ws3.Column(18).Width = 12;      //1st Reviewer
                    ws3.Column(19).Width = 12;      //2nd Reviewer
                    ws3.Column(20).Width = 13.71;   //Testing Status
                    ws3.Column(21).Width = 16.71;   //A2Q2 Due Date (Testing)
                    ws3.Column(22).Width = 15.71;   //Sent to Client Fastly
                    ws3.Column(23).Width = 13.71;   //Client Fastly Review Status
                    ws3.Column(24).Width = 13.71;   //Sent to Deloitte
                    ws3.Column(25).Width = 36.71;   //A2Q2 Notes

                    //disable grid
                    ws3.View.ShowGridLines = false;

                    //set auto filter
                    ws3.Cells[7, 1, 7, 25].AutoFilter = true;

                    //set row
                    int row3 = 1;

                    //set title
                    ws3.Cells[row3, 1].Value = "Back to Navigation";
                    row3++;
                    ws3.Cells[row3, 1].Value = filter.ClientName;
                    row3++;
                    ws3.Cells[row3, 1].Value = "Consolidated IUC List";
                    row3++;
                    ws3.Cells[row3, 1].Value = "SOX FY " + filter.FY;
                    xlsService.ExcelSetFontBold(ws3, 1, 1, row3, 1); //(worksheet, from row, from column, to row, to column)
                    row3 += 3;

                    //set table header
                    ws3.Cells[row3, 1].Value = "No.";
                    ws3.Cells[row3, 2].Value = "Name of IUC";
                    ws3.Cells[row3, 3].Value = "Process";
                    ws3.Cells[row3, 4].Value = "Key Control";
                    ws3.Cells[row3, 5].Value = "Control Activity";
                    ws3.Cells[row3, 6].Value = "Key / Non-Key Control";
                    ws3.Cells[row3, 7].Value = "Key Report";
                    ws3.Cells[row3, 8].Value = "Unique Key Report";
                    ws3.Cells[row3, 9].Value = "Preparer";
                    ws3.Cells[row3, 10].Value = "Process Owner";
                    ws3.Cells[row3, 11].Value = "Key Report Owner";
                    ws3.Cells[row3, 12].Value = "Key Report IT Owner";
                    ws3.Cells[row3, 13].Value = "Set up report Lead Sheet for testing";
                    ws3.Cells[row3, 14].Value = "Schedule Process Owner meeting";
                    ws3.Cells[row3, 15].Value = "Report Received";
                    ws3.Cells[row3, 16].Value = "PBC Status";
                    ws3.Cells[row3, 17].Value = "Tester";
                    ws3.Cells[row3, 18].Value = "1st Reviewer";
                    ws3.Cells[row3, 19].Value = "2nd Reviewer";
                    ws3.Cells[row3, 20].Value = "Testing Status";
                    ws3.Cells[row3, 21].Value = "A2Q2 Due Date (Testing)";
                    ws3.Cells[row3, 22].Value = "Sent to " + filter.ClientName;
                    ws3.Cells[row3, 23].Value = filter.ClientName + " Review Status";
                    ws3.Cells[row3, 24].Value = "Sent to Deloitte";
                    ws3.Cells[row3, 25].Value = "A2Q2 Notes";

                    xlsService.ExcelWrapText(ws3, row3, 1, row3, 25);
                    xlsService.ExcelSetFontBold(ws3, 7, 1, 7, 25);
                    xlsService.ExcelSetHorizontalAlignCenter(ws3, 7, 1, 7, 25);
                    xlsService.ExcelSetVerticalAlignCenter(ws3, 7, 1, 7, 25);
                    xlsService.ExcelSetBackgroundColorRed(ws3, 7, 1, 7, 12);
                    xlsService.ExcelSetBackgroundColorLightBlue1(ws3, 7, 13, 7, 25);
                    xlsService.ExcelSetFontColorWhite(ws3, 7, 1, 7, 12);
                    xlsService.ExcelSetBorder(ws3, 7, 1, 7, 25);
                    xlsService.ExcelSetArialSize10(ws3, 1, 1, row3, 25);


                    num = 1;
                    foreach (var Id in itemId)
                    {
                        var checkTestStatus = _soxContext.KeyReportUserInput.Where(x =>
                                x.ItemId.Equals(Id)
                                && x.AppId.Equals(appIdTestStatus))
                            .AsNoTracking()
                            .ToList();

                        if (checkTestStatus != null && checkTestStatus.Count > 0)
                        {
                            row3++;

                            ws3.Cells[row3, 1].Value = num.ToString(); // "No.";
                            ws3.Cells[row3, 2].Value = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("name of key report/iuc")).Select(x => x.StrAnswer).FirstOrDefault(); // "Name of IUC";
                            ws3.Cells[row3, 3].Value = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("source process")).Select(x => x.StrAnswer).FirstOrDefault(); // "Process";
                            ws3.Cells[row3, 4].Value = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("key control id")).Select(x => x.StrAnswer).FirstOrDefault(); // "Key Control";
                            ws3.Cells[row3, 5].Value = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("control activity")).Select(x => x.StrAnswer).FirstOrDefault(); // "Control Activity";
                            ws3.Cells[row3, 6].Value = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("key/non-key control")).Select(x => x.StrAnswer).FirstOrDefault(); // "Key / Non-Key Control";
                            ws3.Cells[row3, 7].Value = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("key/non-key report")).Select(x => x.StrAnswer).FirstOrDefault(); // "Key Report";
                            ws3.Cells[row3, 8].Value = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("unique key report")).Select(x => x.StrAnswer).FirstOrDefault(); // "Unique Key Report";
                            ws3.Cells[row3, 9].Value = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("preparer")).Select(x => x.StrAnswer).FirstOrDefault(); // "Preparer";
                            ws3.Cells[row3, 10].Value = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("process owner")).Select(x => x.StrAnswer).FirstOrDefault(); // "Process Owner";
                            ws3.Cells[row3, 11].Value = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("key report owner")).Select(x => x.StrAnswer).FirstOrDefault(); // "Key Report Owner";
                            ws3.Cells[row3, 12].Value = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("key report it owner")).Select(x => x.StrAnswer).FirstOrDefault(); // "Key Report IT Owner";
                            ws3.Cells[row3, 13].Value = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("set up report leadsheet for testing")).Select(x => x.StrAnswer).FirstOrDefault(); // "Set up report Lead Sheet for testing";
                            ws3.Cells[row3, 14].Value = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("schedule process owner meeting")).Select(x => x.StrAnswer).FirstOrDefault(); // "Schedule Process Owner meeting";
                            ws3.Cells[row3, 15].Value = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("reports received")).Select(x => x.StrAnswer).FirstOrDefault(); // "Report Received";
                            ws3.Cells[row3, 16].Value = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("pbc status")).Select(x => x.StrAnswer).FirstOrDefault(); // "PBC Status";
                            ws3.Cells[row3, 17].Value = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("tester")).Select(x => x.StrAnswer).FirstOrDefault(); // "Tester";
                            ws3.Cells[row3, 18].Value = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("1st reviewer")).Select(x => x.StrAnswer).FirstOrDefault(); // "1st Reviewer";
                            ws3.Cells[row3, 19].Value = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("2nd reviewer")).Select(x => x.StrAnswer).FirstOrDefault(); // "2nd Reviewer";
                            ws3.Cells[row3, 20].Value = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("testing status")).Select(x => x.StrAnswer).FirstOrDefault(); // "Testing Status";
                            ws3.Cells[row3, 21].Value = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("a2q2 due date (testing)")).Select(x => x.StrAnswer).FirstOrDefault(); // "A2Q2 Due Date (Testing)";
                            ws3.Cells[row3, 22].Value = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("sending to client status")).Select(x => x.StrAnswer).FirstOrDefault(); // "Sent to " + filter.ClientName;
                            ws3.Cells[row3, 23].Value = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("client review status")).Select(x => x.StrAnswer).FirstOrDefault(); // filter.ClientName + " Review Status";
                            ws3.Cells[row3, 24].Value = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("sending to auditor status")).Select(x => x.StrAnswer).FirstOrDefault(); // "Sent to Deloitte";
                            ws3.Cells[row3, 25].Value = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("a2q2 notes")).Select(x => x.StrAnswer).FirstOrDefault(); // "A2Q2 Notes";
                            num++;
                        }

                    }

                   
                    

                    //end 3 Test Status Tracker

                    //4 List of Exceptions Tab
                    ws4.View.ZoomScale = 85;
                    ws4.Column(1).Width = 8.43;
                    ws4.Column(2).Width = 18;
                    ws4.Column(3).Width = 20.29;
                    ws4.Column(4).Width = 47;
                    ws4.Column(5).Width = 20.86;
                    ws4.Column(6).Width = 46.57;
                    ws4.Column(7).Width = 33.57;

                    ws4.Cells[1, 1].Value = "No.";
                    ws4.Cells[1, 2].Value = "Name of IUC";
                    ws4.Cells[1, 3].Value = "Key Control";
                    ws4.Cells[1, 4].Value = "Control Activity";
                    ws4.Cells[1, 5].Value = "Exception noted";
                    ws4.Cells[1, 6].Value = "Reason for Exception";
                    ws4.Cells[1, 7].Value = "Remediation";

                    xlsService.ExcelSetHorizontalAlignCenter(ws4, 1, 1, 1, 7);
                    xlsService.ExcelSetBackgroundColorRed(ws4, 1, 1, 1, 7);
                    xlsService.ExcelSetFontColorWhite(ws4, 1, 1, 1, 7);
                    xlsService.ExcelSetBorder(ws4, 1, 1, 1, 7);
                    xlsService.ExcelSetFontBold(ws4, 1, 1, 1, 7);
                    xlsService.ExcelSetArialSize10(ws4, 1, 1, 1, 7);
                    row = 1;
                    num = 1;
                    foreach (var Id in itemId)
                    {
                        var checkException = _soxContext.KeyReportUserInput.Where(x =>
                                x.ItemId.Equals(Id)
                                && x.AppId.Equals(appIdException))
                            .AsNoTracking()
                            .ToList();

                        if (checkException != null && checkException.Count > 0)
                        {
                            row++;
                            
                            ws4.Cells[row, 1].Value = num.ToString(); // "No.";
                            ws4.Cells[row, 2].Value = checkException.Where(x => x.StrQuestion.ToLower().Contains("name of key report/iuc")).Select(x => x.StrAnswer).FirstOrDefault(); //"Name of IUC";
                            ws4.Cells[row, 3].Value = checkException.Where(x => x.StrQuestion.ToLower().Contains("key control id")).Select(x => x.StrAnswer).FirstOrDefault(); //"Key Control";
                            ws4.Cells[row, 4].Value = checkException.Where(x => x.StrQuestion.ToLower().Contains("control activity")).Select(x => x.StrAnswer).FirstOrDefault(); //"Control Activity";
                            ws4.Cells[row, 5].Value = checkException.Where(x => x.StrQuestion.ToLower().Contains("exceptions noted")).Select(x => x.StrAnswer).FirstOrDefault(); //"Exception noted";
                            ws4.Cells[row, 6].Value = checkException.Where(x => x.StrQuestion.ToLower().Contains("reasons for exceptions")).Select(x => x.StrAnswer).FirstOrDefault(); //"Reason for Exception";
                            ws4.Cells[row, 7].Value = checkException.Where(x => x.StrQuestion.ToLower().Contains("remediation")).Select(x => x.StrAnswer).FirstOrDefault(); //"Remediation";
                            num++;
                        }

                    }



                    //end 4 List of Exceptions

                    //Status Summary Tab
                    ws5.View.ZoomScale = 85;
                    ws5.Cells[1, 1].Value = "By Testing Instance";
                    xlsService.ExcelSetFontColorRed(ws5, 1, 1, 1, 1);
                    xlsService.ExcelSetArialSize10(ws5, 1, 1, 1, 1);

                    //End Status Summary


                    //Start Dropdown Tab
                    ws6.View.ZoomScale = 85;
                    //set column width
                    ws6.Column(1).Width = 26;  //IUC Type
                    ws6.Column(2).Width = 13; //Testing Phase
                    ws6.Column(3).Width = 31.57; //Nature of Testing
                    ws6.Column(4).Width = 26; //
                    ws6.Column(5).Width = 12.29; //
                    ws6.Column(11).Width = 14.14;
                    ws6.Column(14).Width = 13;
                    ws6.Column(16).Width = 11.14;
                    ws6.Column(17).Width = 11.14;

                    ws6.Cells[1, 1].Value = "IUC Type";
                    ws6.Cells[1, 2].Value = "Testing Phase";
                    ws6.Cells[1, 3].Value = "Nature of Testing";
                    ws6.Cells[1, 4].Value = "Effectiveness";
                    ws6.Cells[1, 5].Value = "Canned or Customized";
                    ws6.Cells[1, 6].Value = "Testing Status";
                    ws6.Cells[1, 7].Value = "PBC";
                    ws6.Cells[1, 8].Value = "Review Status";
                    ws6.Cells[1, 9].Value = "Lead Sheet";
                    ws6.Cells[1, 10].Value = "Meetings";
                    ws6.Cells[1, 11].Value = "Documentation";
                    ws6.Cells[1, 12].Value = "IUC Type";
                    ws6.Cells[1, 13].Value = "Accuracy";
                    ws6.Cells[1, 14].Value = "Completeness";
                    ws6.Cells[1, 15].Value = "Source Data";
                    ws6.Cells[1, 16].Value = "Parameters";
                    ws6.Cells[1, 17].Value = "Authorizer";
                    ws6.Cells[1, 18].Value = "Report Status";

                    //format cell
                    xlsService.ExcelWrapText(ws6, 1, 1, 1, 18);
                    xlsService.ExcelSetBackgroundColorGray(ws6, 1, 1, 1, 5);
                    xlsService.ExcelSetBackgroundColorGray(ws6, 1, 12, 1, 18);
                    xlsService.ExcelSetBackgroundColorLightBlue1(ws6, 1, 6, 1, 11);
                    xlsService.ExcelSetHorizontalAlignCenter(ws6, 1, 1, 1, 18); //(worksheet, from row, from column, to row, to column)
                    xlsService.ExcelSetVerticalAlignCenter(ws6, 1, 1, 1, 18); //(worksheet, from row, from column, to row, to column)

                    xlsService.ExcelSetBorder(ws6, 1, 1, 1, 18);
                    xlsService.ExcelSetFontBold(ws6, 1, 1, 1, 18);
                    xlsService.ExcelSetArialSize10(ws6, 1, 1, 1, 18);

                    //end Dropdown Tab
                    //save file
                    string startupPath = Environment.CurrentDirectory;

                    //string strSourceDownload = Path.Combine(startupPath, "include", "keyreport"); 0322
                    string strSourceDownload = Path.Combine(startupPath, "include","upload", "keyreport", "kr tracker");

                    if (!Directory.Exists(strSourceDownload))
                    {
                        Directory.CreateDirectory(strSourceDownload);
                    }
                    var ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    string filename = $"KeyReport-Tracker-{filter.ClientName}-{ts}.xlsx";
                    string strOutput = Path.Combine(strSourceDownload, filename);

                    //Check if file not exists
                    if (System.IO.File.Exists(strOutput))
                    {
                        System.IO.File.Delete(strOutput);
                    }

                    xls.SaveAs(new FileInfo(strOutput));
                    excelFilename = filename;

                    //return status "Ok" with filename
                    return Ok(excelFilename);

                }

                #endregion


            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GenerateKeyReportTracker {ex}", "ErrorGenerateKeyReport Tracker");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GenerateKeyReportTracker");
                if (_environment.IsDevelopment())
                {
                    return BadRequest(ex.ToString());
                }
                else
                {
                    return BadRequest();
                }
            }


            //return Ok(excelFilename);
        }


        [HttpGet("getDownloadableItem")]
        public IActionResult GetKeyReport([FromBody] KeyReportFilter filter)
        {
            try
            {
                string excelFilename = string.Empty;
                #region Create Excel
                //Creating excel 
                //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage xls = new ExcelPackage())
                {
                    ExcelService xlsService = new ExcelService();
                    //set sheet name
                    var ws = xls.Workbook.Worksheets.Add("LeadSheet");
                    var ws1 = xls.Workbook.Worksheets.Add("1. (KeyReportName) - Screenshot");
                    var ws2 = xls.Workbook.Worksheets.Add("2. (Key Report Name)");
                    var ws3 = xls.Workbook.Worksheets.Add("3. Sample Testing_Completeness");
                    var ws4 = xls.Workbook.Worksheets.Add("4.Sample Testing_Accuracy");



                    ws.View.ZoomScale = 90;
                    //set column width
                    ws.Column(1).Width = 9;     //step
                    ws.Column(2).Width = 88;    //procedure    
                    ws.Column(3).Width = 77;    //results
                    ws.Column(4).Width = 50;    //notes


                    //disable grid
                    ws.View.ShowGridLines = false;

                    //set row
                    //int row = 1;

                    //set title
                    //ws.Cells[row, 1].Value = clientName;
                    //row++;
                    //ws.Cells[row, 1].Value = "Key Report";
                    //row++;
                    //ws.Cells[row, 1].Value = Fy;
                    //xlsService.ExcelSetFontBold(ws, 1, 1, row, 1); //(workspace, from row, from column, to row, to column)
                    //row += 3;

                    //ws.Row(row).Height = 85;
                    //set table header
                    //Tab1 Leadsheet
                    ws.Cells[1, 1].Value = "Step #";
                    ws.Cells[1, 2].Value = "Procedures";
                    ws.Cells[1, 3].Value = "Results";

                    ws.Cells[2, 2].Value = "Key Report Information";

                    ws.Cells[3, 1].Value = "1";
                    ws.Cells[4, 1].Value = "2";
                    ws.Cells[5, 1].Value = "3";
                    ws.Cells[6, 1].Value = "4";
                    ws.Cells[7, 1].Value = "5";
                    ws.Cells[8, 1].Value = "6";
                    ws.Cells[9, 1].Value = "7";
                    ws.Cells[10, 1].Value = "8";
                    ws.Cells[11, 1].Value = "9";
                    ws.Cells[12, 1].Value = "10";
                    ws.Cells[13, 1].Value = "11";
                    ws.Cells[14, 1].Value = "12";
                    ws.Cells[15, 1].Value = "13";
                    ws.Cells[16, 1].Value = "14";
                    ws.Cells[17, 1].Value = "15";
                    ws.Cells[18, 1].Value = "16";
                    ws.Cells[19, 1].Value = "17";
                    ws.Cells[20, 1].Value = "18";
                    ws.Cells[21, 1].Value = "19";
                    ws.Cells[22, 1].Value = "20";
                    ws.Cells[23, 1].Value = "";
                    ws.Cells[24, 1].Value = "21";
                    ws.Cells[25, 1].Value = "22";
                    ws.Cells[26, 1].Value = "23";
                    ws.Cells[27, 1].Value = "24";
                    ws.Cells[28, 1].Value = "25";
                    ws.Cells[29, 1].Value = "26";
                    ws.Cells[30, 1].Value = "";
                    ws.Cells[31, 1].Value = "27";
                    ws.Cells[32, 1].Value = "28";
                    ws.Cells[33, 1].Value = "29";




                    String accuracy = "Accuracy";
                    String accuracyB = "<b>" + accuracy + "</b>";
                    String netsuite = "NetSuite";
                    String netsuiteU = "<U>" + netsuite + "</U>";





                    ws.Cells[3, 2].Value = "What is the key report name?";
                    ws.Cells[4, 2].Value = "Was this report previously evaluated for reliability? (If so, link to those procedures)";
                    ws.Cells[5, 2].Value = "What is the purpose of the report?";
                    ws.Cells[6, 2].Value = "What are the key data fields used from this report?";
                    ws.Cells[7, 2].Value = "What type of report is this?";
                    ws.Cells[8, 2].Value = "Define a condition that would represent an exception on this report";
                    ws.Cells[9, 2].Value = "What is the application/system name? (source of the data used to generate the report)";
                    ws.Cells[10, 2].Value = "Is this report: custom, canned, custom query?";
                    ws.Cells[11, 2].Value = "When was the last time this report was modified?";
                    ws.Cells[12, 2].Value = "What are the key report parameters?";
                    ws.Cells[13, 2].Value = "Are parameters input into the report each time it is run?";
                    ws.Cells[14, 2].Value = "How does the report user verify the report is complete?";
                    ws.Cells[15, 2].Value = "How does the report user verify the report is accurate?";
                    ws.Cells[16, 2].Value = "How does the report user verify the report data has integrity?";
                    ws.Cells[17, 2].Value = "Procedures performed to assess the accuracy, completeness, and validity of the source data." +
                       Environment.NewLine + "" +
                       Environment.NewLine + "All Applications" +
                       Environment.NewLine + "Accuracy - trace one item to the outside source" +
                       Environment.NewLine + "Completeness - identify an outside data point and trace back to the report" +
                       Environment.NewLine + "" +
                       Environment.NewLine + "" +
                       Environment.NewLine + "" +
                       Environment.NewLine + "NetSuite" +
                       Environment.NewLine + "Canned Report - Review report name to ensure it is still unmodified. If true, no additional testing needed." +
                       Environment.NewLine + "Customized Report - Review report builder for customization and reconcile against the report requirement / purpose " +
                       Environment.NewLine + "Search - Review search criteria and results tab and reconcile against requirement / purpose ";
                    ws.Cells[18, 2].Value = "Who is the report user (report runner)? Name and Title, date of observation.";
                    ws.Cells[19, 2].Value = "When did the tester observe the report being run?";
                    ws.Cells[20, 2].Value = "Frequency of the report/query generation";
                    ws.Cells[21, 2].Value = "Are there other UAT, Change Management testing, ITGC that support the reliability of this Key Report?";
                    ws.Cells[22, 2].Value = "Is the report output modifiable? (Y/N?)";


                    ws.Cells[23, 2].Value = "Testing Information ";

                    ws.Cells[24, 2].Value = "What date was testing performed?";
                    ws.Cells[25, 2].Value = "Who performed the testing?";
                    ws.Cells[26, 2].Value = "What date was testing reviewed?";
                    ws.Cells[27, 2].Value = "Who performed the review?";
                    ws.Cells[28, 2].Value = "What period did the report cover?";
                    ws.Cells[29, 2].Value = "Was testing performed during another period? If yes, when? ";

                    ws.Cells[30, 2].Value = "Conclusion";
                    ws.Cells[31, 2].Value = "Any exceptions noted?";
                    ws.Cells[32, 2].Value = "Is the report complete and accurate?";
                    ws.Cells[33, 2].Value = "Notes";






                    xlsService.ExcelSetBorder(ws, 1, 1, 33, 3);

                    ws.Cells["A" + 1 + ":A" + 2].Merge = true;
                    ws.Cells["B" + 2 + ":C" + 2].Merge = true;


                    xlsService.ExcelSetBackgroundColorDarkGray(ws, 1, 2, 2, 3);
                    xlsService.ExcelSetFontColorWhite(ws, 1, 2, 2, 3);

                    //Procedures performed to access accuracy, completeness and validity
                    ws.Row(17).Height = 213;
                    xlsService.ExcelWrapText(ws, 17, 2, 17, 3); //(worksheet, from row, from column, to row, to column)
                    xlsService.ExcelSetVerticalAlignCenter(ws, 17, 2, 17, 2);


                    //Testing Information
                    xlsService.ExcelSetBackgroundColorDarkGray(ws, 23, 2, 23, 3);
                    xlsService.ExcelSetFontColorWhite(ws, 23, 2, 23, 3);

                    //Conclusion
                    xlsService.ExcelSetBackgroundColorDarkGray(ws, 30, 2, 30, 3);
                    xlsService.ExcelSetFontColorWhite(ws, 30, 2, 30, 3);

                    for (int row = 1; row <= 33; row++)
                    {
                        xlsService.ExcelSetVerticalAlignCenter(ws, row, 1, row, 1);
                        xlsService.ExcelSetHorizontalAlignCenter(ws, row, 1, row, 1);
                    }


                    xlsService.ExcelSetArialSize10(ws, 1, 1, 33, 3);
                    xlsService.ExcelSetArialSize12(ws, 1, 1, 2, 3);
                    xlsService.ExcelSetArialSize12(ws, 23, 1, 23, 3);
                    xlsService.ExcelSetArialSize12(ws, 30, 1, 30, 3);
                    xlsService.ExcelSetFontBold(ws, 1, 1, 2, 3);




                    //end Tab1 Leadsheet

                    //set border
                    //xlsService.ExcelSetBorder(ws, row, 1, row, 23);

                    //set auto filter
                    //ws.Cells[row, 1, row, 23].AutoFilter = true;

                    //row++;

                    //Tab2 (KeyReportName) - Screenshot

                    //set row
                    //int row = 1;


                    ws1.View.ZoomScale = 90;
                    //set column width
                    ws1.Column(1).Width = 16.43;
                    ws1.Column(2).Width = 20;
                    ws1.Column(3).Width = 20;
                    ws1.Column(4).Width = 20;
                    ws1.Column(5).Width = 20;


                    ws1.Cells[1, 1].Value = "Key Report:";
                    ws1.Cells[2, 1].Value = "Control #:";
                    ws1.Cells[3, 1].Value = "Provided By:";
                    ws1.Cells[4, 1].Value = "Tested by:";
                    ws1.Cells[5, 1].Value = "Description:";
                    ws1.Cells[5, 2].Value = "(System) Screenshot of (Report Name)";




                    xlsService.ExcelSetBackgroundColorDarkBlue(ws1, 1, 1, 5, 1);

                    xlsService.ExcelSetFontColorWhite(ws1, 1, 1, 5, 1);

                    xlsService.ExcelSetBorder(ws1, 1, 1, 5, 5);
                    ws1.Cells["B" + 1 + ":E" + 1].Merge = true;
                    ws1.Cells["B" + 2 + ":E" + 2].Merge = true;
                    ws1.Cells["B" + 3 + ":E" + 3].Merge = true;
                    ws1.Cells["B" + 4 + ":E" + 4].Merge = true;
                    ws1.Cells["B" + 5 + ":E" + 5].Merge = true;
                    xlsService.ExcelSetArialSize10(ws1, 1, 1, 10, 14);
                    xlsService.ExcelSetFontBold(ws1, 1, 1, 5, 1);


                    //end of Tab2

                    //Tab3 (KeyReportName) - 

                    ws2.View.ZoomScale = 90;
                    //set column width
                    ws2.Column(1).Width = 16.43;
                    ws2.Column(2).Width = 20;
                    ws2.Column(3).Width = 20;
                    ws2.Column(4).Width = 20;
                    ws2.Column(5).Width = 20;

                    ws2.Cells[1, 1].Value = "Key Report:";
                    ws2.Cells[2, 1].Value = "Control #:";
                    ws2.Cells[3, 1].Value = "Provided By:";
                    ws2.Cells[4, 1].Value = "Tested by:";
                    ws2.Cells[5, 1].Value = "Description:";
                    ws2.Cells[5, 2].Value = "(Excel / Pdf / Text) Output of the(Report Name)";

                    xlsService.ExcelSetBackgroundColorDarkBlue(ws2, 1, 1, 5, 1);
                    xlsService.ExcelSetFontColorWhite(ws2, 1, 1, 5, 1);

                    xlsService.ExcelSetBorder(ws2, 1, 1, 5, 5);
                    xlsService.ExcelSetArialSize10(ws2, 1, 1, 10, 14);

                    ws2.Cells["B" + 1 + ":E" + 1].Merge = true;
                    ws2.Cells["B" + 2 + ":E" + 2].Merge = true;
                    ws2.Cells["B" + 3 + ":E" + 3].Merge = true;
                    ws2.Cells["B" + 4 + ":E" + 4].Merge = true;
                    ws2.Cells["B" + 5 + ":E" + 5].Merge = true;
                    xlsService.ExcelSetFontBold(ws2, 1, 1, 5, 1);

                    //end of Tab3

                    //Tab4 (3. Sample Testing_Completeness) - 

                    ws3.View.ZoomScale = 90;
                    //set column width
                    ws3.Column(1).Width = 16.43;
                    ws3.Column(2).Width = 20;
                    ws3.Column(3).Width = 20;
                    ws3.Column(4).Width = 20;
                    ws3.Column(5).Width = 20;

                    ws3.Cells[1, 1].Value = "Key Report:";
                    ws3.Cells[2, 1].Value = "Control #:";
                    ws3.Cells[3, 1].Value = "Provided By:";
                    ws3.Cells[4, 1].Value = "Tested by:";
                    ws3.Cells[5, 1].Value = "Description:";
                    ws3.Cells[5, 2].Value = "Comparison Between (System) Screenshot and (Excel/Pdf/Text) Output";


                    ws3.Cells[11, 1].Value = "From screenshots";




                    xlsService.ExcelSetBackgroundColorDarkBlue(ws3, 1, 1, 5, 1);
                    xlsService.ExcelSetFontColorWhite(ws3, 1, 1, 5, 1);
                    xlsService.ExcelSetFontColorRed(ws3, 11, 1, 11, 1);


                    //xlsService.ExcelSetFontBold(ws3, 1, 1, 5, 1);

                    xlsService.ExcelSetArialSize10(ws3, 1, 1, 11, 14);


                    xlsService.ExcelSetBorder(ws3, 1, 1, 5, 5);
                    ws3.Cells[11, 14].Value = "From report";
                    ws3.Cells["B" + 1 + ":E" + 1].Merge = true;
                    ws3.Cells["B" + 2 + ":E" + 2].Merge = true;
                    ws3.Cells["B" + 3 + ":E" + 3].Merge = true;
                    ws3.Cells["B" + 4 + ":E" + 4].Merge = true;
                    ws3.Cells["B" + 5 + ":E" + 5].Merge = true;

                    xlsService.ExcelSetFontBold(ws3, 1, 1, 5, 1);
                    xlsService.ExcelSetFontBold(ws3, 11, 1, 11, 14);
                    xlsService.ExcelSetFontColorRed(ws3, 11, 14, 11, 14);
                    //end of Tab4

                    //Tab5 (4. Sample Testing_Accuracy) - 

                    ws4.View.ZoomScale = 90;
                    //set column width
                    ws4.Column(1).Width = 16.43;
                    ws4.Column(2).Width = 20;
                    ws4.Column(3).Width = 20;
                    ws4.Column(4).Width = 20;
                    ws4.Column(5).Width = 20;

                    ws4.Cells[1, 1].Value = "Key Report:";
                    ws4.Cells[2, 1].Value = "Control #:";
                    ws4.Cells[3, 1].Value = "Provided By:";
                    ws4.Cells[4, 1].Value = "Tested by:";
                    ws4.Cells[5, 1].Value = "Description:";
                    ws4.Cells[5, 2].Value = "Comparison Between (System) Screenshot and (Excel/Pdf/Text) Output";

                    ws4.Cells[11, 1].Value = "From screenshots";


                    xlsService.ExcelSetBackgroundColorDarkBlue(ws4, 1, 1, 5, 1);
                    xlsService.ExcelSetFontColorWhite(ws4, 1, 1, 5, 1);

                    xlsService.ExcelSetFontColorRed(ws4, 11, 1, 11, 1);



                    xlsService.ExcelSetArialSize10(ws4, 1, 1, 11, 14);

                    xlsService.ExcelSetBorder(ws4, 1, 1, 5, 5);
                    ws4.Cells[11, 14].Value = "From report";
                    ws4.Cells["B" + 1 + ":E" + 1].Merge = true;
                    ws4.Cells["B" + 2 + ":E" + 2].Merge = true;
                    ws4.Cells["B" + 3 + ":E" + 3].Merge = true;
                    ws4.Cells["B" + 4 + ":E" + 4].Merge = true;
                    ws4.Cells["B" + 5 + ":E" + 5].Merge = true;
                    xlsService.ExcelSetFontBold(ws4, 1, 1, 5, 1);
                    xlsService.ExcelSetFontBold(ws4, 11, 1, 11, 14);
                    xlsService.ExcelSetFontColorRed(ws4, 11, 14, 11, 14);
                    //end of Tab5



                    //save file
                    string startupPath = Environment.CurrentDirectory;
                   

                    //string strSourceDownload = Path.Combine(startupPath, "include", "keyreport"); 031421
                    string strSourceDownload = Path.Combine(startupPath, "include", "upload","keyreport","leadsheet");

                    if (!Directory.Exists(strSourceDownload))
                    {
                        Directory.CreateDirectory(strSourceDownload);
                    }
                    var ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    string filename = $"KeyReport-{filter.ClientName}-{ts}.xlsx";
                    string strOutput = Path.Combine(strSourceDownload, filename);

                    //Check if file not exists
                    if (System.IO.File.Exists(strOutput))
                    {
                        System.IO.File.Delete(strOutput);
                    }

                    xls.SaveAs(new FileInfo(strOutput));
                    excelFilename = filename;

                    //return status "Ok" with filename
                    return Ok(excelFilename);

                }
                #endregion
            }
            catch (Exception ex)
            {
                Shared.Sox.FileLog.Write(ex.ToString(), "ErrorGetKeyReport");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetKeyReport");
                return BadRequest();
            }
            
        }

        //[AllowAnonymous]

        [HttpGet("getControlActivity")]
        public IActionResult GetKeyReportControlActivity()
        {
            try
            {
                //var controlName = _config.GetSection("ControlName").AsEnumerable().Where(x => x.Value != null).Select(x => x.Value);
                var item = _soxContext.KeyReportControlActivity.Select(x => x.ControlActivity).AsNoTracking().Distinct().ToList();
                if (item != null)
                {
                    return Ok(item);
                }

                return NoContent();
            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorGetKeyReportControlActivity");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetKeyReportControlActivity");
                return BadRequest();
            }

        }

        [HttpGet("getKeyControl")]
        public IActionResult GetKeyReportControl()
        {
            try
            {
                var item = _soxContext.KeyReportKeyControl.Select(x => x.Key).AsNoTracking().Distinct().ToList();
                if (item != null)
                {
                    return Ok(item);
                }

                return NoContent();
            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorGetKeyReportControl");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetKeyReportControl");
                return BadRequest();
            }

        }

        [HttpGet("getReportName")]
        public IActionResult GetKeyReportName()
        {
            try
            {
                var item = _soxContext.KeyReportName.Select(x => x.Name).AsNoTracking().Distinct().ToList();
                if (item != null)
                {
                    return Ok(item);
                }

                return NoContent();
            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorGetKeyReportName");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetKeyReportName");
                return BadRequest();
            }

        }

        [HttpGet("getSystemSource")]
        public IActionResult GetKeyReportSystemSource()
        {
            try
            {
                var item = _soxContext.KeyReportSystemSource.Select(x => x.SystemSource).Distinct().AsNoTracking().ToList();
                if (item != null)
                {
                    return Ok(item);
                }

                return NoContent();
            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorGetKeyReportSystemSource");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetKeyReportSystemSource");
                return BadRequest();
            }

        }

        [HttpGet("getKeyOrNonKeyReport")]
        public IActionResult GetKeyReportOrNonKeyReport()
        {
            try
            {
                var item = _soxContext.KeyReportNonKeyReport.Select(x => x.Report).Distinct().AsNoTracking().ToList();
                if (item != null)
                {
                    return Ok(item);
                }

                return NoContent();
            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorGetKeyReportOrNonKeyReport");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetKeyReportOrNonKeyReport");
                return BadRequest();
            }

        }

        [HttpGet("getReportCustomized")]
        public IActionResult GetKeyReportReportCustomized()
        {
            try
            {
                var item = _soxContext.KeyReportReportCustomized.Select(x => x.ReportCustomized).Distinct().AsNoTracking().ToList();
                if (item != null)
                {
                    return Ok(item);
                }

                return NoContent();
            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorGetKeyReportReportCustomized");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetKeyReportReportCustomized");
                return BadRequest();
            }

        }

        [HttpGet("getIUCType")]
        public IActionResult GetKeyReportIUCType()
        {
            try
            {
                var item = _soxContext.KeyReportIUCType.Select(x => x.IUCType).AsNoTracking().Distinct().ToList();
                if (item != null)
                {
                    return Ok(item);
                }

                return NoContent();
            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorGetKeyReportIUCType");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetKeyReportIUCType");
                return BadRequest();
            }

        }

        [HttpGet("getControlsRelyingIUC")]
        public IActionResult GetKeyReportControlsRelyingIUC()
        {
            try
            {
                var item = _soxContext.KeyReportControlsRelyingIUC.Select(x => x.ControlsRelyingIUC).Distinct().AsNoTracking().ToList();
                if (item != null)
                {
                    return Ok(item);
                }

                return NoContent();
            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorGetKeyReportControlsRelyingIUC");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetKeyReportControlsRelyingIUC");
                return BadRequest();
            }

        }

        [HttpGet("getPreparer")]
        public IActionResult GetKeyReportPreparer()
        {
            try
            {
                var item = _soxContext.KeyReportPreparer.Select(x => x.Preparer).AsNoTracking().Distinct().ToList();
                if (item != null)
                {
                    return Ok(item);
                }

                return NoContent();
            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorGetKeyReportPreparer");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetKeyReportPreparer");
                return BadRequest();
            }

        }

        [HttpGet("getUniqueKeyReport")]
        public IActionResult GetKeyReportUniqueKeyReport()
        {
            try
            {
                var item = _soxContext.KeyReportUniqueKeyReport.Select(x => x.UniqueKeyReport).Distinct().AsNoTracking().ToList();
                if (item != null)
                {
                    return Ok(item);
                }

                return NoContent();
            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorGetKeyReportUniqueKeyReport");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetKeyReportUniqueKeyReport");
                return BadRequest();
            }

        }

        [HttpGet("getNotes")]
        public IActionResult GetKeyReportNotes()
        {
            try
            {
                var item = _soxContext.KeyReportNotes.Select(x => x.ReportNotes).AsNoTracking().Distinct().ToList();
                if (item != null)
                {
                    return Ok(item);
                }

                return NoContent();
            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorGetKeyReportNotes");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetKeyReportNotes");
                return BadRequest();
            }

        }

        [HttpGet("getNumber")]
        public IActionResult GetKeyReportNumber()
        {
            try
            {
                var item = _soxContext.KeyReportNumber.Select(x => x.ReportNumber).Distinct().AsNoTracking().ToList();
                if(item != null)
                {
                    return Ok(item);
                }

                return NoContent();
            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorGetKeyReportNumber");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetKeyReportNumber");
                return BadRequest();
            }

        }

        [HttpGet("getTester")]
        public IActionResult GetKeyReportTester()
        {
            try
            {
                var item = _soxContext.KeyReportTester.Select(x => x.Tester).AsNoTracking().Distinct().ToList();
                if (item != null)
                {
                    return Ok(item);
                }

                return NoContent();
            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorGetKeyReportTester");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetKeyReportTester");
                return BadRequest();
            }

        }

        [HttpGet("getReviewer")]
        public IActionResult GetKeyReportReviewer()
        {
            try
            {
                var item = _soxContext.KeyReportReviewer.Select(x => x.Reviewer).Distinct().AsNoTracking().ToList();
                if (item != null)
                {
                    return Ok(item);
                }

                return NoContent();
            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorGetKeyReportReviewer");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetKeyReportReviewer");
                return BadRequest();
            }

        }

        //[AllowAnonymous]
        [HttpGet("getRcmControlId")]
        public IActionResult GetRcmControlId()
        {
            try
            {
                var item = _soxContext.RcmControlId.AsNoTracking().Select(x => x.ControlId).Distinct();
                if (item != null)
                {
                    return Ok(item.ToList());
                }

                return NoContent();
            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorGetRcmControlId");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetRcmControlId");
                return BadRequest();
            }

        }
        
        //[AllowAnonymous]
        [HttpGet("getRcmProcess")]
        public IActionResult GetRcmProcess()
        {
            try
            {
                var item = _soxContext.RcmProcess.AsNoTracking().Select(x => x.Process).Distinct();
                if (item != null)
                {
                    return Ok(item.ToList());
                }

                return NoContent();
            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorGetRcmProcess");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetRcmProcess");
                return BadRequest();
            }

        }

        //[AllowAnonymous]
        [HttpGet("getRcmControlOwner")]
        public IActionResult GetRcmControlOwner()
        {
            try
            {
                var item = _soxContext.RcmControlOwner.Select(x => x.ControlOwner).Distinct().AsNoTracking();
                if (item != null)
                {
                    //List<RcmControlOwner> listRcmControlOwner = new List<RcmControlOwner>();
                    //foreach (var control in item)
                    //{
                    //    listRcmControlOwner.Add(new RcmControlOwner { ControlOwner = control });
                    //}
                    
                              
                    return Ok(item.ToList());
                }

                return NoContent();
            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorGetRcmControlOwner");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetRcmControlOwner");
                return BadRequest();
            }

        }

        //[AllowAnonymous]
        [HttpGet("leadsheet/getfy")]
        public IActionResult GetLeadsheetFy()
        {
            try
            {
                //var item = _soxContext.KeyReportUserInput.Where(x => 
                //        x.StrQuestion.ToLower().Contains("1. what is the fy")
                //    )
                //    .Select(x => x.StrAnswer)
                //    .Distinct()
                //    .AsNoTracking()
                //    .ToList();
                var item = _soxContext.KeyReportUserInput
                    .Where(x => x.TagFY != null && x.TagFY != string.Empty && x.TagStatus.ToLower() != "inactive")
                    .Select(x => x.TagFY)
                    .Distinct()
                    .AsNoTracking()
                    .ToList();
                if (item != null)
                {
                    return Ok(item.OrderBy(x => x));
                }

                return NoContent();
            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorGetLeadsheetFy");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetLeadsheetFy");
                return BadRequest();
            }

        }

        //[AllowAnonymous]
        [HttpPost("leadsheet/getclient")]
        public IActionResult GetLeadsheetClient([FromBody] KeyReportFilter filter)
        {
            try
            {
                //var inputItemId = _soxContext.KeyReportUserInput.Where(x =>
                //        x.StrAnswer.ToLower().Equals(filter.FY)
                //    )
                //    .AsNoTracking()
                //    .Select(x => x.ItemId)                    
                //    .Distinct()                    
                //    .ToList();
                //if (inputItemId != null)
                //{
                //    List<string> clientName = new List<string>();
                //    foreach (var item in inputItemId)
                //    {
                //        var clientList = _soxContext.KeyReportUserInput.Where(x =>
                //            x.ItemId.Equals(item)
                //            && x.StrQuestion.ToLower().Contains("2. client name")
                //        )
                //        .AsNoTracking()
                //        .Select(x => x.StrAnswer)
                //        .Distinct()
                //        .ToList();

                //        foreach (var client in clientList)
                //        {
                //            if(!clientName.Contains(client) && client != string.Empty)
                //            {
                //                clientName.Add(client);
                //            }
                //        }

                //    }


                //    return Ok(clientName);
                //}

                var inputItemId = _soxContext.KeyReportUserInput.Where(x =>
                        x.TagFY.ToLower().Equals(filter.FY)
                        && x.TagClientName != null 
                        && x.TagClientName != string.Empty
                        && x.TagStatus.ToLower() != "inactive"
                    )
                    .AsNoTracking()
                    .Select(x => x.TagClientName)
                    .Distinct()
                    .ToList();
                if (inputItemId != null)
                {
                    List<string> clientName = new List<string>();
                    foreach (var item in inputItemId)
                    {
                        if (!clientName.Contains(item) && item != string.Empty)
                        {
                            clientName.Add(item);
                        }
                    }

                    return Ok(clientName.OrderBy(x => x));
                }

                return NoContent();
            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorGetLeadsheetClient");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetLeadsheetClient");
                return BadRequest();
            }

        }

        //[AllowAnonymous]
        [HttpPost("leadsheet/getreportname")]
        public IActionResult GetLeadsheetReportName([FromBody] KeyReportFilter filter)
        {
            try
            {
                //string appId = _config.GetSection("KeyReportApp").GetSection("AllIUCId").Value;
                //List<string> reportName = new List<string>();
                //var inputItemId = _soxContext.KeyReportUserInput.Where(x =>
                //        x.StrAnswer.ToLower().Equals(filter.FY)
                //        && x.AppId.Equals(appId)
                //    )
                //    .AsNoTracking()
                //    .Select(x => x.ItemId)
                //    .Distinct()
                //    .ToList();
                //if (inputItemId != null)
                //{

                //    foreach (var item in inputItemId)
                //    {
                //        var inputItemId2 = _soxContext.KeyReportUserInput.Where(x =>
                //            x.ItemId.Equals(item)
                //            && (
                //                x.StrAnswer.ToLower().Equals(filter.ClientName)
                //            )
                //        )
                //        .AsNoTracking()
                //        .Select(x => x.ItemId)
                //        .Distinct()
                //        .ToList();

                //        if(inputItemId2 != null)
                //        {
                //            foreach (var itemId2 in inputItemId2)
                //            {
                //                var clientList = _soxContext.KeyReportUserInput.Where(x =>
                //                        x.ItemId.Equals(itemId2)
                //                        && (
                //                            x.StrQuestion.ToLower().Contains("6. name of key report/iuc") ||
                //                            x.StrQuestion.ToLower().Contains("3. name of key report/iuc")
                //                    )
                //                )
                //                .Select(x => x.StrAnswer)
                //                .Distinct()
                //                .AsNoTracking()
                //                .ToList();

                //                foreach (var client in clientList)
                //                {
                //                    if (!reportName.Contains(client) && client != string.Empty && client != null)
                //                    {
                //                        reportName.Add(client);
                //                    }
                //                }
                //            }
                //        }
                //    }

                //    return Ok(reportName);
                //}

                //return NoContent();


                string appId = _config.GetSection("KeyReportApp").GetSection("AllIUCId").Value;
                List<string> reportName = new List<string>();
                //List<string> listReportName = new List<string>();
                //List<int> listItemId = new List<int>();

                var listReportName = _soxContext.KeyReportUserInput.Where(x =>
                           x.TagFY.ToLower().Equals(filter.FY.ToLower())
                           && x.TagClientName.ToLower().Equals(filter.ClientName.ToLower())
                           && x.TagReportName != null
                           && x.TagReportName != string.Empty
                           && x.TagStatus.ToLower() != "inactive"
                       )
                       .AsNoTracking()
                       .Select(x => new { x.TagReportName, x.ItemId})
                       .Distinct()
                       .ToList();


    
                if (listReportName != null)
                {
                    foreach (var item in listReportName)
                    {
                        if(filter.UniqueKey == "All")
                        {
                            if (!reportName.Contains(item.TagReportName) && item.TagReportName != string.Empty && item.TagReportName != null)
                            {
                                reportName.Add(item.TagReportName);
                            }
                        }
                        else if (filter.UniqueKey == "Yes")
                        {
                            var listUniqueReportName = _soxContext.KeyReportUserInput.Where(x =>
                                   x.ItemId.Equals(item.ItemId)
                                   && x.StrQuestion.ToLower().Contains("unique key report")
                                   && x.StrAnswer.ToLower().Equals("yes")
                               )
                               .AsNoTracking()
                               .Select(x =>  x.TagReportName)
                               .FirstOrDefault();
                            if(listUniqueReportName != null && listUniqueReportName != string.Empty && !reportName.Contains(listUniqueReportName))
                            {   
                                reportName.Add(item.TagReportName);
                            }
                        }
                        else if (filter.UniqueKey == "No")
                        {
                            var listUniqueReportName = _soxContext.KeyReportUserInput.Where(x =>
                                   x.ItemId.Equals(item.ItemId)
                                   && x.StrQuestion.ToLower().Contains("unique key report")
                                   && x.StrAnswer.ToLower().Equals("no")
                               )
                               .AsNoTracking()
                               .Select(x => x.TagReportName)
                               .FirstOrDefault();
                            if (listUniqueReportName != null && listUniqueReportName != string.Empty && !reportName.Contains(listUniqueReportName))
                            {
                                reportName.Add(item.TagReportName);
                            }
                        }
                    }
                    
                    //Return report name
                    return Ok(reportName.OrderBy(x => x));
                
                }
               
                




                return NoContent();


            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorGetLeadsheetReportName");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetLeadsheetReportName");
                return BadRequest();
            }

        }

        //[AllowAnonymous]
        [HttpPost("getreportname")]
        public IActionResult GetReportName([FromBody] KeyReportFilter filter)
        {
            try
            {
                //List<string> reportName = new List<string>();
                //var inputItemId = _soxContext.KeyReportUserInput.Where(x =>
                //        x.StrAnswer.ToLower().Equals(filter.FY)
                //    )
                //    .AsNoTracking()
                //    .Select(x => x.ItemId)
                //    .Distinct()
                //    .ToList();
                //if (inputItemId != null)
                //{
                //    foreach (var item in inputItemId)
                //    {
                //        var inputItemId2 = _soxContext.KeyReportUserInput.Where(x =>
                //            x.ItemId.Equals(item)
                //            && (
                //                x.StrAnswer.ToLower().Equals(filter.ClientName)
                //            )
                //        )
                //        .AsNoTracking()
                //        .Select(x => x.ItemId)
                //        .Distinct()
                //        .ToList();
                //        if (inputItemId2 != null)
                //        {
                //            foreach (var itemId2 in inputItemId2)
                //            {
                //                var clientList = _soxContext.KeyReportUserInput.Where(x =>
                //                        x.ItemId.Equals(itemId2)
                //                        && (
                //                            x.StrQuestion.ToLower().Contains("6. name of key report/iuc") ||
                //                            x.StrQuestion.ToLower().Contains("3. name of key report/iuc")
                //                    )
                //                )
                //                .Select(x => x.StrAnswer)
                //                .Distinct()
                //                .AsNoTracking()
                //                .ToList();
                //                foreach (var client in clientList)
                //                {
                //                    if (!reportName.Contains(client) && client != string.Empty && client != null)
                //                    {
                //                        reportName.Add(client);
                //                    }
                //                }
                //            }
                //        }
                //    }
                //    return Ok(reportName);
                //}

                List<string> reportName = new List<string>();
                var inputItemId = _soxContext.KeyReportUserInput.Where(x =>
                        x.TagFY.ToLower().Equals(filter.FY.ToLower())
                        && x.TagClientName.ToLower().Equals(filter.ClientName.ToLower())
                        && x.TagReportName != null
                        && x.TagReportName != string.Empty
                        && x.TagStatus.ToLower() != "inactive"
                    )
                    .AsNoTracking()
                    .Select(x => x.TagReportName)
                    .Distinct()
                    .ToList();
                if (inputItemId != null)
                {

                    foreach (var item in inputItemId)
                    {
                        if (!reportName.Contains(item) && item != string.Empty && item != null)
                        {
                            reportName.Add(item);
                        }
                    }
                    return Ok(reportName.OrderBy(x => x));
                }

                return NoContent();
            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorGetReportName");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetReportName");
                return BadRequest();
            }

        }

  
        //[AllowAnonymous]
        [HttpPost("getcontrolid")]
        public IActionResult GetControlId([FromBody] KeyReportFilter filter)
        {
            try
            {
                //List<string> listControlId = new List<string>();
                //var inputItemId = _soxContext.KeyReportUserInput.Where(x =>
                //        x.StrAnswer.ToLower().Equals(filter.FY)
                //    )
                //    .AsNoTracking()
                //    .Select(x => x.ItemId)
                //    .Distinct()
                //    .ToList();
                //if (inputItemId != null)
                //{
                //    foreach (var item in inputItemId)
                //    {
                //        var inputItemId2 = _soxContext.KeyReportUserInput.Where(x =>
                //            x.ItemId.Equals(item)
                //            && (
                //                x.StrAnswer.ToLower().Equals(filter.ClientName)
                //            )
                //        )
                //        .AsNoTracking()
                //        .Select(x => x.ItemId)
                //        .Distinct()
                //        .ToList();
                //        if (inputItemId2 != null)
                //        {
                //            foreach (var itemId2 in inputItemId2)
                //            {
                //                var clientList = _soxContext.KeyReportUserInput.Where(x =>
                //                        x.ItemId.Equals(itemId2)
                //                        && (
                //                            x.StrQuestion.ToLower().Contains("6. name of key report/iuc") ||
                //                            x.StrQuestion.ToLower().Contains("3. name of key report/iuc")
                //                    )
                //                )
                //                .AsNoTracking()
                //                .Select(x => x.ItemId)
                //                .Distinct()
                //                .ToList();
                //                foreach (var client in clientList)
                //                {
                //                    var controlList = _soxContext.KeyReportUserInput.Where(x =>
                //                            x.ItemId.Equals(itemId2)
                //                            && (
                //                                x.StrQuestion.ToLower().Contains("key control id")
                //                        )
                //                    )
                //                    .Select(x => x.StrAnswer)
                //                    .Distinct()
                //                    .AsNoTracking()
                //                    .ToList();
                //                    foreach (var itemControl in controlList)
                //                    {
                //                        if (!listControlId.Contains(itemControl) && itemControl != string.Empty && itemControl != null)
                //                        {
                //                            listControlId.Add(itemControl);
                //                        }
                //                    }                  
                //                }
                //            }
                //        }
                //    }
                //    return Ok(listControlId);
                //}

                List<string> listControlId = new List<string>();
                var inputItemId = _soxContext.KeyReportUserInput.Where(x =>
                        x.TagFY.ToLower().Equals(filter.FY.ToLower())
                        && x.TagClientName.ToLower().Equals(filter.ClientName.ToLower())
                        && x.TagReportName.ToLower().Equals(filter.KeyReportName.ToLower())
                        && x.TagControlId != null
                        && x.TagControlId != string.Empty
                        && x.TagStatus.ToLower() != "inactive"
                    )
                    .AsNoTracking()
                    .Select(x => x.TagControlId)
                    .Distinct()
                    .ToList();
                if (inputItemId != null)
                {
                    foreach (var item in inputItemId)
                    {
                        if (!listControlId.Contains(item) && item != string.Empty && item != null)
                        {
                            listControlId.Add(item);
                        }
                    }
                    return Ok(listControlId);
                }


                return NoContent();
            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorGetControlId");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetControlId");
                return BadRequest();
            }

        }

        //[AllowAnonymous]
        [HttpPost("getitemid")]
        public IActionResult GetItemId([FromBody] KeyReportFilter filter)
        {
            try
            {
                

                KeyReportItemId keyReportItemId = new KeyReportItemId();
                string consolOrigId = _config.GetSection("KeyReportApp").GetSection("ConsolOrigFormatId").Value;
                string allIUCId = _config.GetSection("KeyReportApp").GetSection("AllIUCId").Value;
                string testStatusTrackerId = _config.GetSection("KeyReportApp").GetSection("TestStatusTrackerId").Value;
                string exceptionId = _config.GetSection("KeyReportApp").GetSection("ExceptionId").Value;

                var listOrigFormatId = SearchItemId2(filter, consolOrigId);
                var listAllIUCId = SearchItemId2(filter, allIUCId);
                var listTestId = SearchItemId2(filter, testStatusTrackerId);
                var listExceptionId = SearchItemId2(filter, exceptionId);

                if (listOrigFormatId != null && listOrigFormatId.Any())
                    keyReportItemId.OrigFormatItemId = listOrigFormatId[0];

                if (listAllIUCId != null && listAllIUCId.Any())
                    keyReportItemId.AllIUCItemId = listAllIUCId[0];

                if (listTestId != null && listTestId.Any())
                    keyReportItemId.TestItemId = listTestId[0];

                if (listExceptionId != null && listExceptionId.Any())
                    keyReportItemId.ExceptionItemId = listExceptionId[0];


                return Ok(keyReportItemId);
            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorGetItemId");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetItemId");
                return BadRequest();
            }

        }


        //[AllowAnonymous]
        [HttpGet("getkeyreportitem/{itemId:int}")]
        public IActionResult GetKeyReportwithItemID(int itemId)
        {
            try
            {

                var listKeyReportItem = GetKeyReportItem(itemId);

                return Ok(listKeyReportItem);
            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorGetItemId");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetItemId");
                return BadRequest();
            }

        }

        [AllowAnonymous]
        [HttpPost("getscreenshot")]
        public IActionResult GetScreenshot([FromBody] KeyReportFilter filter)
        {
            try
            {
                List<KeyReportScreenshotUpload> listScreenshot = new List<KeyReportScreenshotUpload>();

                //check if not exists
                var checkScreenshot = _soxContext.KeyReportScreenshotUpload.Where(x =>
                    x.Client.Equals(filter.ClientName) &&
                    x.ReportName.Equals(filter.KeyReportName) &&
                    x.Fy.Equals(filter.FY) &&
                    x.ControlId.Equals(filter.ControlId)
                ).ToList();

                if(checkScreenshot != null && checkScreenshot.Any())
                {
                    listScreenshot = checkScreenshot;
                }

                return Ok(listScreenshot);

            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorGetScreenshot");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetScreenshot");
                return BadRequest();
            }

        }

        [AllowAnonymous]
        [HttpPost("getreportfile")]
        public IActionResult GetKeyReportFile([FromBody] KeyReportFilter filter)
        {
            try
            {
                List<KeyReportFile> listReportFile = new List<KeyReportFile>();

                //check if not exists
                var checkReportFile = _soxContext.KeyReportFile.Where(x =>
                    x.Client.Equals(filter.ClientName) &&
                    x.ReportName.Equals(filter.KeyReportName) &&
                    x.Fy.Equals(filter.FY) &&
                    x.ControlId.Equals(filter.ControlId)
                ).ToList();

                if (checkReportFile != null && checkReportFile.Any())
                {
                    listReportFile = checkReportFile;
                }

                return Ok(listReportFile);

            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorGetKeyReportFile");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetKeyReportFile");
                return BadRequest();
            }

        }


        [AllowAnonymous]
        [HttpPost("savepodio/leadsheet")]
        public IActionResult SavePodioKeyReportLeadsheet()
        {
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("KeyReportApp").GetSection("LeadsheetAppId").Value;
                PodioKey.ClientSecret = _config.GetSection("KeyReportApp").GetSection("LeadsheetAppToken").Value;

                return Ok();
            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorSavePodioKeyReportLeadsheet");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SavePodioKeyReportLeadsheet");
                return BadRequest();
            }
        }
        
        
        [HttpPost("savepodio/origformat")]
        public async Task<IActionResult> SavePodioKeyReportOrigFormat([FromBody] List<KeyReportUserInput> listUserInput)
        {
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("ConsolOrigFormatId").Value;
                PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("ConsolOrigFormatToken").Value;

                if (PodioAppKey.AppId != string.Empty && PodioAppKey.AppToken != string.Empty)
                {
                    var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                    ServicePointManager.Expect100Continue = true;
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                    await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                    if (podio.IsAuthenticated())
                    {
                       
                        var checkItemId = listUserInput.Where(x => x.ItemId != 0).FirstOrDefault();

                        Item questionnaireItem;
                        if (checkItemId == null)
                            questionnaireItem = new Item();
                        else
                            questionnaireItem = new Item { ItemId = checkItemId.ItemId };


                        questionnaireItem = await PodioItems(listUserInput, questionnaireItem, podio);


                        if (checkItemId == null)
                        {
                            
                            var itemCreateId = await podio.ItemService.AddNewItem(Int32.Parse(PodioAppKey.AppId), questionnaireItem);
                            listUserInput = listUserInput.Select(c => { c.ItemId = itemCreateId; return c; }).ToList();
                        }

                        else
                        {
                            var itemUpdateId = await podio.ItemService.UpdateItem(questionnaireItem, null, null, false, true);
                        }

                        return Ok(listUserInput);
                    }
                }

                return BadRequest();

            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex);
                FileLog.Write(ex.ToString(), "ErrorSavePodioKeyReportOrigFormat");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SavePodioKeyReportOrigFormat");
                return BadRequest();
            }
        }

        [AllowAnonymous]
        [HttpPost("savepodio/alliuc")]
        public async Task<IActionResult> SavePodioKeyReportIUCKRQuestionnaire([FromBody] List<KeyReportUserInput> listUserInput)
        {
            try
            {
                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("AllIUCId").Value;
                PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("AllIUCToken").Value;

                if (PodioAppKey.AppId != string.Empty && PodioAppKey.AppToken != string.Empty)
                {
                    var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                    ServicePointManager.Expect100Continue = true;
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                    await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                    if (podio.IsAuthenticated())
                    {
                        var checkItemId = listUserInput.Where(x => x.ItemId != 0).FirstOrDefault();

                        Item questionnaireItem;
                        if (checkItemId == null)
                            questionnaireItem = new Item();
                        else
                            questionnaireItem = new Item { ItemId = checkItemId.ItemId };


                        questionnaireItem = await PodioItems(listUserInput, questionnaireItem, podio);


                        if (checkItemId == null)
                        {
                            var itemCreateId = await podio.ItemService.AddNewItem(Int32 .Parse(PodioAppKey.AppId), questionnaireItem);
                            listUserInput = listUserInput.Select(c => { c.ItemId = itemCreateId; return c; }).ToList();
                        }

                        else
                        {
                            var itemUpdateId = await podio.ItemService.UpdateItem(questionnaireItem, null, null, false, true);
                        }

                        return Ok(listUserInput);
                    }
                }

                return BadRequest();
            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorSavePodioKeyReportIUCKRQuestionnaire");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SavePodioKeyReportIUCKRQuestionnaire");
                return BadRequest();
            }
        }

        //[AllowAnonymous]
        [HttpPost("savepodio/teststatus")]
        public async Task<IActionResult> SavePodioKeyReportIUCReportList([FromBody] List<KeyReportUserInput> listUserInput)
        {
            try
            {
                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("TestStatusTrackerId").Value;
                PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("TestStatusTrackerToken").Value;

                if (PodioAppKey.AppId != string.Empty && PodioAppKey.AppToken != string.Empty)
                {
                    var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                    ServicePointManager.Expect100Continue = true;
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                    await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                    if (podio.IsAuthenticated())
                    {
                        var checkItemId = listUserInput.Where(x => x.ItemId != 0).FirstOrDefault();

                        Item questionnaireItem;
                        if (checkItemId == null)
                            questionnaireItem = new Item();
                        else
                            questionnaireItem = new Item { ItemId = checkItemId.ItemId };


                        questionnaireItem = await PodioItems(listUserInput, questionnaireItem, podio);


                        if (checkItemId == null)
                        {
                            var itemCreateId = await podio.ItemService.AddNewItem(Int32.Parse(PodioAppKey.AppId), questionnaireItem);
                            listUserInput = listUserInput.Select(c => { c.ItemId = itemCreateId; return c; }).ToList();
                        }

                        else
                        {
                            var itemUpdateId = await podio.ItemService.UpdateItem(questionnaireItem, null, null, false, true);
                        }

                        return Ok(listUserInput);
                    }
                }

                return BadRequest();
            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorSavePodioKeyReportIUCReportList");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SavePodioKeyReportIUCReportList");
                return BadRequest();
            }
        }
        
        //[AllowAnonymous]
        [HttpPost("savepodio/exception")]
        public async Task<IActionResult> SavePodioKeyReportException([FromBody] List<KeyReportUserInput> listUserInput)
        {
            try
            {
                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("ExceptionId").Value;
                PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("ExceptionToken").Value;

                if (PodioAppKey.AppId != string.Empty && PodioAppKey.AppToken != string.Empty)
                {
                    var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                    ServicePointManager.Expect100Continue = true;
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                    await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                    if (podio.IsAuthenticated())
                    {
                        var checkItemId = listUserInput.Where(x => x.ItemId != 0).FirstOrDefault();

                        Item questionnaireItem;
                        if (checkItemId == null)
                            questionnaireItem = new Item();
                        else
                            questionnaireItem = new Item { ItemId = checkItemId.ItemId };


                        questionnaireItem = await PodioItems(listUserInput, questionnaireItem, podio);


                        if (checkItemId == null)
                        {
                            var itemCreateId = await podio.ItemService.AddNewItem(Int32.Parse(PodioAppKey.AppId), questionnaireItem);
                            listUserInput = listUserInput.Select(c => { c.ItemId = itemCreateId; return c; }).ToList();
                        }

                        else
                        {
                            var itemUpdateId = await podio.ItemService.UpdateItem(questionnaireItem, null, null, false, true);
                        }

                        return Ok(listUserInput);
                    }
                }

                return BadRequest();
            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorSavePodioKeyReportException");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SavePodioKeyReportException");
                return BadRequest();
            }
        }


        //[AllowAnonymous]
        [HttpPost("save/leadsheet")]
        public IActionResult SaveKeyReportLeadsheet()
        {
            try
            {
                return Ok();
            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorSaveKeyReportLeadsheet");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SaveKeyReportLeadsheet");
                return BadRequest();
            }
        }
        
        //[AllowAnonymous] //mark
        [HttpPost("save/keyreport")] //UPDATE HERE IF KEYREPORT EXIST 
        public async Task<IActionResult> SaveKeyReportOrigFormat([FromBody] KeyReportIds listUserInput)
        {
            try
            {
                if (listUserInput.ListReport != null && listUserInput.ListReport.Count > 0)
                {
                   
                        foreach (var item in listUserInput.ListReport)
                        {

                            using (var context = _soxContext.Database.BeginTransaction())
                            {
                                //check if item exists 
                                //true - update
                               
                                var checkItemExist = _soxContext.KeyReportUserInput.Where(x => x.FieldId.Value.Equals(item.FieldId) && x.ItemId.Equals(item.ItemId)).FirstOrDefault();
                                if (checkItemExist != null)
                                {
                                //Update SoxTracker And KeyReport
                                var PullSoxTracker = _soxContext.SoxTracker
                                        .Where(x =>
                                        x.FY.Equals(item.TagFY) &&
                                        x.ClientName.Equals(item.TagClientName) &&
                                        x.KeyReportName.Equals(item.TagReportName) &&
                                        x.ControlId.Equals(item.TagControlId) &&
                                        x.Status.Equals("Active")).FirstOrDefault();

                                var PullRcm = _soxContext.Rcm.Where(x =>
                                     x.FY.Equals(PullSoxTracker.FY) &&
                                     x.ClientName.Equals(PullSoxTracker.ClientName) &&
                                     x.ControlId.Equals(PullSoxTracker.ControlId) &&
                                     x.Status.Equals("Active")).FirstOrDefault();


                                    if (listUserInput.ConsolidatedId == item.ItemId) // if update is from consolidated
                                    {
                                        switch (item.StrQuestion.ToLower())
                                        {

                                            // update AllUIC
                                            #region

                                            case string s when s.Contains("9. iuc type"):

                                                var nineiuctype = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.UicId) && x.StrQuestion.ToLower().Contains("10. iuc type")).FirstOrDefault();

                                                if (nineiuctype != null)
                                                {
                                                    nineiuctype.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(nineiuctype);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                break;

                                            case string s when s.Contains("10. system/ source"):

                                                var tensystemsource = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.UicId) && x.StrQuestion.ToLower().Contains("11. system / source")).FirstOrDefault();

                                                if (tensystemsource != null)
                                                {
                                                    tensystemsource.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(tensystemsource);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                break;

                                            case string s when s.Contains("11. report customized"):

                                                var reportcustomized = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.UicId) && x.StrQuestion.ToLower().Contains("12. report customized")).FirstOrDefault();

                                                if (reportcustomized != null)
                                                {
                                                    reportcustomized.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(reportcustomized);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                break;


                                            case string s when s.Contains("12. controls relying on iuc"):

                                                var controlsrelying = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.UicId) && x.StrQuestion.ToLower().Contains("13. controls relying on key report")).FirstOrDefault();

                                                if (controlsrelying != null)
                                                {
                                                    controlsrelying.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(controlsrelying);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                break;

                                            case string s when s.Contains("14. preparer"):

                                                var preparer = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.UicId) && x.StrQuestion.ToLower().Contains("14. preparer")).FirstOrDefault();

                                                if (preparer != null)
                                                {
                                                    preparer.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(preparer);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                break;

                                            case string s when s.Contains("15. tester"):

                                                var tester = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.UicId) && x.StrQuestion.ToLower().Contains("15. tester")).FirstOrDefault();

                                                if (tester != null)
                                                {
                                                    tester.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(tester);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                break;

                                            case string s when s.Contains("16. 1st reviewer"):
                                                var reviewer = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.UicId) && x.StrQuestion.ToLower().Contains("16. reviewer")).FirstOrDefault();

                                                var fistreviewer = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.UicId) && x.StrQuestion.ToLower().Contains("17. 1st reviewer")).FirstOrDefault();

                                                if (reviewer != null)
                                                {
                                                    reviewer.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(reviewer);
                                                    await _soxContext.SaveChangesAsync();
                                                }

                                                if (fistreviewer != null)
                                                {
                                                    fistreviewer.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(fistreviewer);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                break;

                                            case string ss when ss.Contains("17. 2nd reviewer"):
                                                var secondreviewer = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.UicId) && x.StrQuestion.ToLower().Contains("18. 2nd reviewer")).FirstOrDefault();
                                                if (secondreviewer != null)
                                                {
                                                    secondreviewer.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(secondreviewer);
                                                    await _soxContext.SaveChangesAsync();
                                                }

                                                break;

                                            case string s when s.Contains("18. a2q2 notes"):

                                                var notes = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.UicId) && x.StrQuestion.ToLower().Contains("45. notes")).FirstOrDefault();

                                                if (notes != null)
                                                {
                                                    notes.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(notes);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                break;

                                                #endregion

                                        }
                                        switch (item.StrQuestion.ToLower())
                                        {
                                            // update testtracker
                                            #region


                                            case string s when s.Contains("14. preparer"):

                                                var preparer2 = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.TestId) && x.StrQuestion.ToLower().Contains("10. preparer")).FirstOrDefault();

                                                if (preparer2 != null)
                                                {
                                                    preparer2.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(preparer2);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                break;


                                            case string s when s.Contains("15. tester"):

                                                var tester2 = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.TestId) && x.StrQuestion.ToLower().Contains("18. tester")).FirstOrDefault();

                                                if (tester2 != null)
                                                {
                                                    tester2.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(tester2);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                break;

                                            case string s when s.Contains("13. process owner"):

                                                var processowner = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.TestId) && x.StrQuestion.ToLower().Contains("11. process owner")).FirstOrDefault();

                                                if (processowner != null)
                                                {
                                                    processowner.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(processowner);
                                                    await _soxContext.SaveChangesAsync();
                                                }

                                                break;

                                            case string s when s.Contains("16. 1st reviewer"):


                                                var fistreviewer22 = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.TestId) && x.StrQuestion.ToLower().Contains("19. 1st reviewer")).FirstOrDefault();


                                                if (fistreviewer22 != null)
                                                {
                                                    fistreviewer22.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(fistreviewer22);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                break;

                                            case string s when s.Contains("17. 2nd reviewer"):

                                                var secondreviewer2 = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.TestId) && x.StrQuestion.ToLower().Contains("20. 2nd reviewer")).FirstOrDefault();

                                                if (secondreviewer2 != null)
                                                {
                                                    secondreviewer2.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(secondreviewer2);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                break;

                                            case string s when s.Contains("35. it report owner"):

                                                var itowner = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.TestId) && x.StrQuestion.ToLower().Contains("13. key report it owner")).FirstOrDefault();

                                                if (itowner != null)
                                                {
                                                    itowner.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(itowner);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                break;

                                                #endregion

                                        }
                                        switch (item.StrQuestion.ToLower())
                                        {
                                            //update rcm relationships 
                                            #region 
                                            case string s when s.Contains("4. control activity"):

                                                    var ControlActivity = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.UicId) && x.StrQuestion.ToLower().Contains("4. control activity")).FirstOrDefault();
                                                
                                                    var ControlActivity2 = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.TestId) && x.StrQuestion.ToLower().Contains("6. control activity")).FirstOrDefault();

                                                    var ControlActivity3 = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.ExceptionId) && x.StrQuestion.ToLower().Contains("5. control activity")).FirstOrDefault();

                                                    if (ControlActivity != null)
                                                    {
                                                        ControlActivity.StrAnswer = item.StrAnswer;
                                                        _soxContext.Update(ControlActivity);
                                                        await _soxContext.SaveChangesAsync();
                                                    }
                                                    if (ControlActivity2 != null)
                                                    {
                                                        ControlActivity2.StrAnswer = item.StrAnswer;
                                                        _soxContext.Update(ControlActivity2);
                                                        await _soxContext.SaveChangesAsync();
                                                    }

                                                    if (ControlActivity3 != null)
                                                    {
                                                        ControlActivity3.StrAnswer = item.StrAnswer;
                                                        _soxContext.Update(ControlActivity3);
                                                        await _soxContext.SaveChangesAsync();
                                                    }

                                                    PullRcm.ControlActivityFy19 = item.StrAnswer;
                                                    _soxContext.Update(PullRcm);
                                                    await _soxContext.SaveChangesAsync();

                                                    break;


                                                case string ss when ss.Contains("8. key/non-key report"):
                                                case string s when s.Contains("5. key/non-key control"):


                                                    var KeyNonKey = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.UicId) && x.StrQuestion.ToLower().Contains("5. key/non-key control")).FirstOrDefault();
                                                    var KeyNonKey2 = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.UicId) && x.StrQuestion.ToLower().Contains("8. key/non-key report")).FirstOrDefault();


                                                    if (KeyNonKey != null || KeyNonKey2 != null)
                                                    {
                                                        KeyNonKey.StrAnswer = item.StrAnswer;
                                                        _soxContext.Update(KeyNonKey);

                                                        KeyNonKey2.StrAnswer = item.StrAnswer;
                                                        _soxContext.Update(KeyNonKey2);
                                                        await _soxContext.SaveChangesAsync();
                                                    }

                                                    var KeyNonKey3 = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.TestId) && x.StrQuestion.ToLower().Contains("7. key/non-key control")).FirstOrDefault();
                                                    var KeyNonKey4 = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.TestId) && x.StrQuestion.ToLower().Contains("8. key/non-key report")).FirstOrDefault();


                                                    if (KeyNonKey3 != null || KeyNonKey4 != null)
                                                    {
                                                        KeyNonKey3.StrAnswer = item.StrAnswer;
                                                        _soxContext.Update(KeyNonKey3);

                                                        KeyNonKey4.StrAnswer = item.StrAnswer;
                                                        _soxContext.Update(KeyNonKey4);
                                                        await _soxContext.SaveChangesAsync();
                                                    }

                                                    PullRcm.Key = item.StrAnswer;
                                                    _soxContext.Update(PullRcm);
                                                    await _soxContext.SaveChangesAsync();

                                                    break;

                                                case string s when s.Contains("7. source process"):

                                                    var SourceProcess = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.UicId) && x.StrQuestion.ToLower().Contains("7. source process")).FirstOrDefault();

                                                    if (SourceProcess != null)
                                                    {
                                                        SourceProcess.StrAnswer = item.StrAnswer;
                                                        _soxContext.Update(SourceProcess);
                                                        await _soxContext.SaveChangesAsync();
                                                    }

                                                    var SourceProcess2 = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.TestId) && x.StrQuestion.ToLower().Contains("4. source process")).FirstOrDefault();

                                                    if (SourceProcess2 != null)
                                                    {
                                                        SourceProcess2.StrAnswer = item.StrAnswer;
                                                        _soxContext.Update(SourceProcess2);
                                                        await _soxContext.SaveChangesAsync();
                                                    }

                                                    PullRcm.Process = item.StrAnswer;
                                                    _soxContext.Update(PullRcm);
                                                    await _soxContext.SaveChangesAsync();

                                                    break;
                                                #endregion
                                        }

                                        item.Id = checkItemExist.Id;
                                        _soxContext.Entry(checkItemExist).CurrentValues.SetValues(item);
                                        await _soxContext.SaveChangesAsync();

                                    }

                                    if (listUserInput.UicId == item.ItemId) //if update is from uic
                                    {

                                        switch (item.StrQuestion.ToLower())
                                        {
                                            //Update Consolidated
                                            #region
                                            case string s when s.Contains("10. iuc type"):

                                                var nineiuctype = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.ConsolidatedId) && x.StrQuestion.ToLower().Contains("9. iuc type")).FirstOrDefault();

                                                if (nineiuctype != null)
                                                {
                                                    nineiuctype.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(nineiuctype);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                break;

                                            case string s when s.Contains("11. system/ source"):

                                                var tensystemsource = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.ConsolidatedId) && x.StrQuestion.ToLower().Contains("10. system / source")).FirstOrDefault();

                                                if (tensystemsource != null)
                                                {
                                                    tensystemsource.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(tensystemsource);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                break;

                                            case string s when s.Contains("12. report customized"):

                                                var reportcustomized = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.ConsolidatedId) && x.StrQuestion.ToLower().Contains("11. report customized")).FirstOrDefault();

                                                if (reportcustomized != null)
                                                {
                                                    reportcustomized.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(reportcustomized);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                break;


                                            case string s when s.Contains("13. controls relying on iuc"):

                                                var controlsrelying = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.ConsolidatedId) && x.StrQuestion.ToLower().Contains("12. controls relying on key report")).FirstOrDefault();

                                                if (controlsrelying != null)
                                                {
                                                    controlsrelying.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(controlsrelying);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                break;

                                            case string s when s.Contains("14. preparer"):

                                                var preparer = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.ConsolidatedId) && x.StrQuestion.ToLower().Contains("14. preparer")).FirstOrDefault();

                                                if (preparer != null)
                                                {
                                                    preparer.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(preparer);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                break;

                                            case string s when s.Contains("15. tester"):

                                                var tester = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.ConsolidatedId) && x.StrQuestion.ToLower().Contains("15. tester")).FirstOrDefault();

                                                if (tester != null)
                                                {
                                                    tester.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(tester);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                break;

                                            case string s when s.Contains("17. 1st reviewer"):
                                                //var reviewer = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.UicId) && x.StrQuestion.ToLower().Contains("16. reviewer")).FirstOrDefault();
                                                //if (reviewer != null)
                                                //{
                                                //  reviewer.StrAnswer = item.StrAnswer;
                                                //_soxContext.Update(reviewer);
                                                //await _soxContext.SaveChangesAsync();
                                                //} walang reviewer si consolidarted

                                                var fistreviewer = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.ConsolidatedId) && x.StrQuestion.ToLower().Contains("16. 1st reviewer")).FirstOrDefault();


                                                if (fistreviewer != null)
                                                {
                                                    fistreviewer.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(fistreviewer);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                break;

                                            case string ss when ss.Contains("18. 2nd reviewer"):
                                                var secondreviewer = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.ConsolidatedId) && x.StrQuestion.ToLower().Contains("17. 2nd reviewer")).FirstOrDefault();
                                                if (secondreviewer != null)
                                                {
                                                    secondreviewer.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(secondreviewer);
                                                    await _soxContext.SaveChangesAsync();
                                                }

                                                break;

                                            case string s when s.Contains("45. a2q2 notes"):

                                                var notes = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.ConsolidatedId) && x.StrQuestion.ToLower().Contains("18. notes")).FirstOrDefault();

                                                if (notes != null)
                                                {
                                                    notes.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(notes);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                break;
                                                #endregion

                                        }
                                        switch (item.StrQuestion.ToLower())
                                        {
                                            // update testTracker
                                            #region
                                            case string s when s.Contains("unique key report"):
                                                var uniquekey = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.TestId) && x.StrQuestion.ToLower().Contains("unique key report")).FirstOrDefault();

                                                if (uniquekey != null)
                                                {
                                                    uniquekey.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(uniquekey);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                break;

                                            case string s when s.Contains("14. preparer"):

                                                var preparer2 = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.TestId) && x.StrQuestion.ToLower().Contains("10. preparer")).FirstOrDefault();

                                                if (preparer2 != null)
                                                {
                                                    preparer2.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(preparer2);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                break;



                                            case string s when s.Contains("15. tester"):

                                                var tester2 = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.TestId) && x.StrQuestion.ToLower().Contains("18. tester")).FirstOrDefault();

                                                if (tester2 != null)
                                                {
                                                    tester2.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(tester2);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                break;



                                            case string s when s.Contains("39. process owner"):

                                                var processowner = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.TestId) && x.StrQuestion.ToLower().Contains("11. process owner")).FirstOrDefault();

                                                if (processowner != null)
                                                {
                                                    processowner.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(processowner);
                                                    await _soxContext.SaveChangesAsync();
                                                }

                                                break;

                                            case string s when s.Contains("17. 1st reviewer"):


                                                var fistreviewer22 = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.TestId) && x.StrQuestion.ToLower().Contains("19. 1st reviewer")).FirstOrDefault();


                                                if (fistreviewer22 != null)
                                                {
                                                    fistreviewer22.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(fistreviewer22);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                break;

                                            case string s when s.Contains("18. 2nd reviewer"):

                                                var secondreviewer2 = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.TestId) && x.StrQuestion.ToLower().Contains("20. 2nd reviewer")).FirstOrDefault();

                                                if (secondreviewer2 != null)
                                                {
                                                    secondreviewer2.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(secondreviewer2);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                break;



                                            case string s when s.Contains("35. it report owner"):

                                                var itowner = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.TestId) && x.StrQuestion.ToLower().Contains("13. key report it owner")).FirstOrDefault();

                                                if (itowner != null)
                                                {
                                                    itowner.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(itowner);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                break;

                                                #endregion

                                        }
                                        switch (item.StrQuestion.ToLower())
                                        {
                                            //update rcm relationships 
                                            #region 
                                            case string s when s.Contains("4. control activity"):

                                                var ControlActivity = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.ConsolidatedId) && x.StrQuestion.ToLower().Contains("4. control activity")).FirstOrDefault();

                                                var ControlActivity2 = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.TestId) && x.StrQuestion.ToLower().Contains("6. control activity")).FirstOrDefault();

                                                var ControlActivity3 = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.ExceptionId) && x.StrQuestion.ToLower().Contains("5. control activity")).FirstOrDefault();

                                                if (ControlActivity != null)
                                                {
                                                    ControlActivity.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(ControlActivity);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                if (ControlActivity2 != null)
                                                {
                                                    ControlActivity2.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(ControlActivity2);
                                                    await _soxContext.SaveChangesAsync();
                                                }

                                                if (ControlActivity3 != null)
                                                {
                                                    ControlActivity3.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(ControlActivity3);
                                                    await _soxContext.SaveChangesAsync();
                                                }

                                                PullRcm.ControlActivityFy19 = item.StrAnswer;
                                                _soxContext.Update(PullRcm);
                                                await _soxContext.SaveChangesAsync();

                                                break;


                                            case string ss when ss.Contains("8. key/non-key report"):
                                            case string s when s.Contains("5. key/non-key control"):


                                                var KeyNonKey = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.ConsolidatedId) && x.StrQuestion.ToLower().Contains("5. key/non-key control")).FirstOrDefault();
                                                var KeyNonKey2 = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.ConsolidatedId) && x.StrQuestion.ToLower().Contains("8. key/non-key report")).FirstOrDefault();


                                                if (KeyNonKey != null || KeyNonKey2 != null)
                                                {
                                                    KeyNonKey.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(KeyNonKey);

                                                    KeyNonKey2.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(KeyNonKey2);
                                                    await _soxContext.SaveChangesAsync();
                                                }

                                                var KeyNonKey3 = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.TestId) && x.StrQuestion.ToLower().Contains("7. key/non-key control")).FirstOrDefault();
                                                var KeyNonKey4 = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.TestId) && x.StrQuestion.ToLower().Contains("8. key/non-key report")).FirstOrDefault();


                                                if (KeyNonKey3 != null || KeyNonKey4 != null)
                                                {
                                                    KeyNonKey3.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(KeyNonKey3);

                                                    KeyNonKey4.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(KeyNonKey4);
                                                    await _soxContext.SaveChangesAsync();
                                                }

                                                PullRcm.Key = item.StrAnswer;
                                                _soxContext.Update(PullRcm);
                                                await _soxContext.SaveChangesAsync();

                                                break;

                                            case string s when s.Contains("7. source process"):

                                                var SourceProcess = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.UicId) && x.StrQuestion.ToLower().Contains("7. source process")).FirstOrDefault();

                                                if (SourceProcess != null)
                                                {
                                                    SourceProcess.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(SourceProcess);
                                                    await _soxContext.SaveChangesAsync();
                                                }

                                                var SourceProcess2 = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.TestId) && x.StrQuestion.ToLower().Contains("4. source process")).FirstOrDefault();

                                                if (SourceProcess2 != null)
                                                {
                                                    SourceProcess2.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(SourceProcess2);
                                                    await _soxContext.SaveChangesAsync();
                                                }

                                                PullRcm.Process = item.StrAnswer;
                                                _soxContext.Update(PullRcm);
                                                await _soxContext.SaveChangesAsync();

                                                break;
                                                #endregion
                                        }

                                        item.Id = checkItemExist.Id;
                                        _soxContext.Entry(checkItemExist).CurrentValues.SetValues(item);
                                        await _soxContext.SaveChangesAsync();
                                    }

                                    if (listUserInput.TestId == item.ItemId) //if update is from test tracker
                                    {

                                        switch (item.StrQuestion.ToLower()) // update Consolidated
                                        {

                                            // update Consolidated
                                            #region

                                            case string s when s.Contains("10. preparer"):

                                                var preparer = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.ConsolidatedId) && x.StrQuestion.ToLower().Contains("14. preparer")).FirstOrDefault();

                                                if (preparer != null)
                                                {
                                                    preparer.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(preparer);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                break;

                                            case string s when s.Contains("11. process Owner"):

                                                var tester = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.ConsolidatedId) && x.StrQuestion.ToLower().Contains("16. 1st reviewer")).FirstOrDefault();

                                                if (tester != null)
                                                {
                                                    tester.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(tester);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                break;

                                            case string s when s.Contains("19. 1st reviewer"):


                                                var fistreviewer22 = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.ConsolidatedId) && x.StrQuestion.ToLower().Contains("16. 1st reviewer")).FirstOrDefault();


                                                if (fistreviewer22 != null)
                                                {
                                                    fistreviewer22.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(fistreviewer22);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                break;

                                            case string s when s.Contains("20. 2nd reviewer"):

                                                var secondreviewer2 = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.ConsolidatedId) && x.StrQuestion.ToLower().Contains("17. 2nd reviewer")).FirstOrDefault();

                                                if (secondreviewer2 != null)
                                                {
                                                    secondreviewer2.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(secondreviewer2);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                break;
                                                #endregion

                                        }
                                        switch (item.StrQuestion.ToLower())
                                        {
                                            //update uic
                                            #region

                                            case string s when s.Contains("unique key report"):
                                                var uniquekey = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.UicId) && x.StrQuestion.ToLower().Contains("unique key report")).FirstOrDefault();

                                                if (uniquekey != null)
                                                {
                                                    uniquekey.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(uniquekey);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                break;

                                            case string s when s.Contains("10. preparer"):

                                                var preparer2 = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.UicId) && x.StrQuestion.ToLower().Contains("14. preparer")).FirstOrDefault();

                                                if (preparer2 != null)
                                                {
                                                    preparer2.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(preparer2);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                break;

                                            case string s when s.Contains("11. process owner"):

                                                var tester2 = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.UicId) && x.StrQuestion.ToLower().Contains("39. process owner")).FirstOrDefault();

                                                if (tester2 != null)
                                                {
                                                    tester2.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(tester2);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                break;

                                            case string s when s.Contains("13. key report it owner"):

                                                var processowner = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.UicId) && x.StrQuestion.ToLower().Contains("35. it report owner")).FirstOrDefault();

                                                if (processowner != null)
                                                {
                                                    processowner.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(processowner);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                break;


                                            case string s when s.Contains("19. 1st reviewer"):


                                                var fistreviewer222 = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.UicId) && x.StrQuestion.ToLower().Contains("17. 1st reviewer")).FirstOrDefault();


                                                if (fistreviewer222 != null)
                                                {
                                                    fistreviewer222.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(fistreviewer222);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                break;



                                            case string s when s.Contains("20. 2nd reviewer"):

                                                var fistreviewer12 = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.UicId) && x.StrQuestion.ToLower().Contains("18. 2nd reviewer")).FirstOrDefault();


                                                if (fistreviewer12 != null)
                                                {
                                                    fistreviewer12.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(fistreviewer12);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                break;

                                                #endregion

                                        }
                                        switch (item.StrQuestion.ToLower())
                                        {

                                            //update rcm relationships 
                                            #region 
                                            case string s when s.Contains("6. control activity"):

                                                        var ControlActivity = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.ConsolidatedId) && x.StrQuestion.ToLower().Contains("4. control activity")).FirstOrDefault();

                                                        var ControlActivity2 = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.UicId) && x.StrQuestion.ToLower().Contains("4. control activity")).FirstOrDefault();

                                                        var ControlActivity3 = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.ExceptionId) && x.StrQuestion.ToLower().Contains("5. control activity")).FirstOrDefault();

                                                        if (ControlActivity != null)
                                                        {
                                                            ControlActivity.StrAnswer = item.StrAnswer;
                                                            _soxContext.Update(ControlActivity);
                                                            await _soxContext.SaveChangesAsync();
                                                        }
                                                        if (ControlActivity2 != null)
                                                        {
                                                            ControlActivity2.StrAnswer = item.StrAnswer;
                                                            _soxContext.Update(ControlActivity2);
                                                            await _soxContext.SaveChangesAsync();
                                                        }

                                                        if (ControlActivity3 != null)
                                                        {
                                                            ControlActivity3.StrAnswer = item.StrAnswer;
                                                            _soxContext.Update(ControlActivity3);
                                                            await _soxContext.SaveChangesAsync();
                                                        }

                                                        PullRcm.ControlActivityFy19 = item.StrAnswer;
                                                        _soxContext.Update(PullRcm);
                                                        await _soxContext.SaveChangesAsync();

                                                        break;


                                                    case string ss when ss.Contains("8. key/non-key report"):
                                                    case string s when s.Contains("7. key/non-key control"):


                                                        var KeyNonKey = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.ConsolidatedId) && x.StrQuestion.ToLower().Contains("5. key/non-key control")).FirstOrDefault();
                                                        var KeyNonKey2 = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.ConsolidatedId) && x.StrQuestion.ToLower().Contains("8. key/non-key report")).FirstOrDefault();


                                                        if (KeyNonKey != null || KeyNonKey2 != null)
                                                        {
                                                            KeyNonKey.StrAnswer = item.StrAnswer;
                                                            _soxContext.Update(KeyNonKey);

                                                            KeyNonKey2.StrAnswer = item.StrAnswer;
                                                            _soxContext.Update(KeyNonKey2);
                                                            await _soxContext.SaveChangesAsync();
                                                        }

                                                        var KeyNonKey3 = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.UicId) && x.StrQuestion.ToLower().Contains("5. key/non-key control")).FirstOrDefault();
                                                        var KeyNonKey4 = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.UicId) && x.StrQuestion.ToLower().Contains("8. key/non-key report")).FirstOrDefault();


                                                        if (KeyNonKey3 != null || KeyNonKey4 != null)
                                                        {
                                                            KeyNonKey3.StrAnswer = item.StrAnswer;
                                                            _soxContext.Update(KeyNonKey3);

                                                            KeyNonKey4.StrAnswer = item.StrAnswer;
                                                            _soxContext.Update(KeyNonKey4);
                                                            await _soxContext.SaveChangesAsync();
                                                        }

                                                        PullRcm.Key = item.StrAnswer;
                                                        _soxContext.Update(PullRcm);
                                                        await _soxContext.SaveChangesAsync();

                                                        break;

                                                    case string s when s.Contains("4. source process"):

                                                        var SourceProcess = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.UicId) && x.StrQuestion.ToLower().Contains("7. source process")).FirstOrDefault();

                                                        if (SourceProcess != null)
                                                        {
                                                            SourceProcess.StrAnswer = item.StrAnswer;
                                                            _soxContext.Update(SourceProcess);
                                                            await _soxContext.SaveChangesAsync();
                                                        }

                                                        var SourceProcess2 = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.ConsolidatedId) && x.StrQuestion.ToLower().Contains("7. source process")).FirstOrDefault();

                                                        if (SourceProcess2 != null)
                                                        {
                                                            SourceProcess2.StrAnswer = item.StrAnswer;
                                                            _soxContext.Update(SourceProcess2);
                                                            await _soxContext.SaveChangesAsync();
                                                        }

                                                        PullRcm.Process = item.StrAnswer;
                                                        _soxContext.Update(PullRcm);
                                                        await _soxContext.SaveChangesAsync();

                                                        break;
                                                        #endregion

                                        }


                                        item.Id = checkItemExist.Id;
                                        _soxContext.Entry(checkItemExist).CurrentValues.SetValues(item);
                                        await _soxContext.SaveChangesAsync();
                                    }

                                    if (listUserInput.ExceptionId == item.ItemId)
                                    {

                                        switch (item.StrQuestion.ToLower())
                                        {

                                            case string s when s.Contains("5. control activity"):

                                                var ControlActivity = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.UicId) && x.StrQuestion.ToLower().Contains("4. control activity")).FirstOrDefault();

                                                var ControlActivity2 = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.TestId) && x.StrQuestion.ToLower().Contains("6. control activity")).FirstOrDefault();

                                                var ControlActivity3 = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(listUserInput.ConsolidatedId) && x.StrQuestion.ToLower().Contains("4. control activity")).FirstOrDefault();

                                                if (ControlActivity != null)
                                                {
                                                    ControlActivity.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(ControlActivity);
                                                    await _soxContext.SaveChangesAsync();
                                                }
                                                if (ControlActivity2 != null)
                                                {
                                                    ControlActivity2.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(ControlActivity2);
                                                    await _soxContext.SaveChangesAsync();
                                                }

                                                if (ControlActivity3 != null)
                                                {
                                                    ControlActivity3.StrAnswer = item.StrAnswer;
                                                    _soxContext.Update(ControlActivity3);
                                                    await _soxContext.SaveChangesAsync();
                                                }

                                                PullRcm.ControlActivityFy19 = item.StrAnswer;
                                                _soxContext.Update(PullRcm);
                                                await _soxContext.SaveChangesAsync();

                                            break;

                                        }
                                        item.Id = checkItemExist.Id;
                                        _soxContext.Entry(checkItemExist).CurrentValues.SetValues(item);
                                        await _soxContext.SaveChangesAsync();
                                    }


                            }
                                await context.CommitAsync();
                                
                            }
                          
                        }
                    
                }
                
                return Ok();
            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorSaveKeyReportOrigFormat");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SaveKeyReportOrigFormat");
                return BadRequest();
            }
        }

        //[AllowAnonymous]
        [HttpPost("save/keyreport_new")]  //ADD HERE IF KEYREPORT IS NEW
        public async Task<IActionResult> SaveIfNew([FromBody] KeyReportIds listUserInput)
        {
            try
            {
                
                if (listUserInput.ListReport != null && listUserInput.ListReport.Count > 0)
                {

                    foreach (var item in listUserInput.ListReport)
                    {
                        
                        using (var context = _soxContext.Database.BeginTransaction())
                        {
                            //check if item exists 
                            //true - update
                            //false - insert

                            var checkItemExist = _soxContext.KeyReportUserInput.Where(x => x.FieldId.Value.Equals(item.FieldId) && x.ItemId.Equals(item.ItemId)).FirstOrDefault();
                            
                            if (checkItemExist != null)
                            {
                                if(checkItemExist.StrAnswer != item.StrAnswer) { 
                                    checkItemExist.StrAnswer = item.StrAnswer;
                                    await _soxContext.SaveChangesAsync();
                                }
                                // no need to update to avoid overide
                            }
                            else
                            {
                                _soxContext.Add(item);
                                await _soxContext.SaveChangesAsync();
                            }
                            await context.CommitAsync();

                        }

                    }

                }

                return Ok();

            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorSaveKeyReportOrigFormat");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SaveKeyReportOrigFormat");
                return BadRequest();
            }

        }

        [AllowAnonymous]
        [HttpPost("save/screenshot")] 
        public async Task<IActionResult> SaveScreenshot([FromBody] List<KeyReportScreenshotUpload> listScreenshot)
        {
            try
            {

                if (listScreenshot != null && listScreenshot.Any())
                {

                    var unique = listScreenshot.Select(x => new { x.Client, x.ReportName, x.Fy, x.ControlId }).FirstOrDefault();

                    foreach (var item in listScreenshot)
                    {

                        using (var context = _soxContext.Database.BeginTransaction())
                        {

                            //check if not exists
                            var checkItemExist = _soxContext.KeyReportScreenshotUpload.Where(x => 
                                x.NewFilename.Equals(item.NewFilename) && 
                                x.Client.Equals(item.Client) &&
                                x.ReportName.Equals(item.ReportName) &&
                                x.Fy.Equals(item.Fy) &&
                                x.ControlId.Equals(item.ControlId)
                            ).FirstOrDefault();

                            if (checkItemExist == null)
                            {
                                _soxContext.Add(item);
                                await _soxContext.SaveChangesAsync();
                                await context.CommitAsync();
                            }
                           
                        }

                    }

                    //check for items that are remove
                    var allScreenshot = _soxContext.KeyReportScreenshotUpload.Where(x =>
                                x.Client.Equals(unique.Client) &&
                                x.ReportName.Equals(unique.ReportName) &&
                                x.Fy.Equals(unique.Fy) &&
                                x.ControlId.Equals(unique.ControlId)).ToList();

                    var inDBButNotInListScreenshot = allScreenshot.Except(listScreenshot);
                    if(inDBButNotInListScreenshot != null && inDBButNotInListScreenshot.Any())
                    {
                        foreach (var item in inDBButNotInListScreenshot)
                        {
                            using (var context = _soxContext.Database.BeginTransaction())
                            {
                                    _soxContext.Remove(item);
                                    await _soxContext.SaveChangesAsync();
                                    await context.CommitAsync();
                              

                            }
                        }
                        
                    }

                }

                return Ok();

            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorSaveScreenshot");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SaveScreenshot");
                return BadRequest();
            }

        }

        [AllowAnonymous]
        [HttpPost("save/reportfile")]
        public async Task<IActionResult> SaveReportFile([FromBody] List<KeyReportFile> listReportFile)
        {
            try
            {

                if (listReportFile != null && listReportFile.Any())
                {

                    var unique = listReportFile.Select(x => new { x.Client, x.ReportName, x.Fy, x.ControlId }).FirstOrDefault();

                    foreach (var item in listReportFile)
                    {

                        using (var context = _soxContext.Database.BeginTransaction())
                        {

                            //check if not exists
                            var checkItemExist = _soxContext.KeyReportFile.Where(x =>
                                x.NewFilename.Equals(item.NewFilename) &&
                                x.Client.Equals(item.Client) &&
                                x.ReportName.Equals(item.ReportName) &&
                                x.Fy.Equals(item.Fy) &&
                                x.ControlId.Equals(item.ControlId)
                            ).FirstOrDefault();

                            if (checkItemExist == null)
                            {
                                _soxContext.Add(item);
                                await _soxContext.SaveChangesAsync();
                                await context.CommitAsync();
                            }

                        }

                    }

                    //check for items that are remove
                    var allScreenshot = _soxContext.KeyReportFile.Where(x =>
                                x.Client.Equals(unique.Client) &&
                                x.ReportName.Equals(unique.ReportName) &&
                                x.Fy.Equals(unique.Fy) &&
                                x.ControlId.Equals(unique.ControlId)).ToList();

                    var inDBButNotInListScreenshot = allScreenshot.Except(listReportFile);
                    if (inDBButNotInListScreenshot != null && inDBButNotInListScreenshot.Any())
                    {
                        foreach (var item in inDBButNotInListScreenshot)
                        {
                            using (var context = _soxContext.Database.BeginTransaction())
                            {
                                _soxContext.Remove(item);
                                await _soxContext.SaveChangesAsync();
                                await context.CommitAsync();


                            }
                        }

                    }

                }

                return Ok();

            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorSaveReportFile");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SaveReportFile");
                return BadRequest();
            }

        }


        //[AllowAnonymous]
        [HttpPost("keyreport1")]
        public IActionResult GetKeyReport1Control([FromBody] KeyReportFilter filter)
        {

            try
            {

                //we get workpaper or questionnaire per 

                string consolOrigId = _config.GetSection("KeyReportApp").GetSection("ConsolOrigFormatId").Value;
                //string allIUCId = _config.GetSection("KeyReportApp").GetSection("AllIUCId").Value;
                //string testStatusTrackerId = _config.GetSection("KeyReportApp").GetSection("TestStatusTrackerId").Value;
                //string exceptionId = _config.GetSection("KeyReportApp").GetSection("ExceptionId").Value;

                var checkKeyReportQuestionnaire = _soxContext.KeyReportQuestion
                    .Where(x => x.AppId.Equals(consolOrigId))  //25886802
                    .Include(x => x.Options)
                    .AsNoTracking();

                if (checkKeyReportQuestionnaire != null)
                {
                    return Ok(checkKeyReportQuestionnaire.ToList());
                }
                else
                    return NoContent();

             


            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetKeyReport1Control {ex}", "ErrorGetKeyReport1Controll");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetKeyReport1Controll");
                if (_environment.IsDevelopment())
                {
                    return BadRequest(ex.ToString());
                }
                else
                {
                    return BadRequest();
                }
            }




        }

        //[AllowAnonymous]
        [HttpPost("keyreport2")]
        public IActionResult GetKeyReport2Control([FromBody] KeyReportFilter filter)
        {

            try
            {

                //we get workpaper or questionnaire per 

                //string consolOrigId = _config.GetSection("KeyReportApp").GetSection("ConsolOrigFormatId").Value;
                string allIUCId = _config.GetSection("KeyReportApp").GetSection("AllIUCId").Value;
                //string testStatusTrackerId = _config.GetSection("KeyReportApp").GetSection("TestStatusTrackerId").Value;
                //string exceptionId = _config.GetSection("KeyReportApp").GetSection("ExceptionId").Value;

                var checkKeyReportQuestionnaire = _soxContext.KeyReportQuestion
                    .Where(x => x.AppId.Equals(allIUCId))
                    .Include(x => x.Options).AsNoTracking();

                if (checkKeyReportQuestionnaire != null)
                {
                    return Ok(checkKeyReportQuestionnaire.ToList());
                }
                else
                    return NoContent();

                    


            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetKeyReport3Control {ex}", "ErrorGetKeyReport3Control");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetKeyReport3Control");
                if (_environment.IsDevelopment())
                {
                    return BadRequest(ex.ToString());
                }
                else
                {
                    return BadRequest();
                }
            }




        }


        //[AllowAnonymous]
        [HttpPost("keyreport3")]
        public IActionResult GetKeyReport3Control([FromBody] KeyReportFilter filter)
        {

            try
            {

                //we get workpaper or questionnaire per 

                //string consolOrigId = _config.GetSection("KeyReportApp").GetSection("ConsolOrigFormatId").Value;
                //string allIUCId = _config.GetSection("KeyReportApp").GetSection("AllIUCId").Value;
                string testStatusTrackerId = _config.GetSection("KeyReportApp").GetSection("TestStatusTrackerId").Value;
                //string exceptionId = _config.GetSection("KeyReportApp").GetSection("ExceptionId").Value;

                var checkKeyReportQuestionnaire = _soxContext.KeyReportQuestion
                    .Where(x => x.AppId.Equals(testStatusTrackerId))
                    .Include(x => x.Options).AsNoTracking();

                if (checkKeyReportQuestionnaire != null)
                {
                    return Ok(checkKeyReportQuestionnaire.ToList());
                }
                else
                    return NoContent();

            
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetKeyReport3Control {ex}", "ErrorGetKeyReport3Control");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetKeyReport3Control");
                if (_environment.IsDevelopment())
                {
                    return BadRequest(ex.ToString());
                }
                else
                {
                    return BadRequest();
                }
            }




        }


        //[AllowAnonymous]
        [HttpPost("keyreport4")]
        public IActionResult GetKeyReport4Control([FromBody] KeyReportFilter filter)
        {

            try
            {

                //we get workpaper or questionnaire per 

                //string consolOrigId = _config.GetSection("KeyReportApp").GetSection("ConsolOrigFormatId").Value;
                //string allIUCId = _config.GetSection("KeyReportApp").GetSection("AllIUCId").Value;
                //string testStatusTrackerId = _config.GetSection("KeyReportApp").GetSection("TestStatusTrackerId").Value;
                string exceptionId = _config.GetSection("KeyReportApp").GetSection("ExceptionId").Value;

                var checkKeyReportQuestionnaire = _soxContext.KeyReportQuestion
                    .Where(x => x.AppId.Equals(exceptionId))
                    .Include(x => x.Options).AsNoTracking();

                if (checkKeyReportQuestionnaire != null)
                {
                    return Ok(checkKeyReportQuestionnaire.ToList());
                }
                else
                    return NoContent();

            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetKeyReport4Control {ex}", "ErrorGetKeyReport4Control");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetKeyReport4Control");
                if (_environment.IsDevelopment())
                {
                    return BadRequest(ex.ToString());
                }
                else
                {
                    return BadRequest();
                }
            }




        }



        private async Task<Item> PodioItems(List<KeyReportUserInput> listUserInput, Item questionnaireItem, Podio podio)
        {
            foreach (var item in listUserInput)
            {

                if (item.FieldId != null)
                {
                    switch (item.Type)
                    {
                        case "text":
                            if (item.StrAnswer != null && item.StrAnswer != string.Empty)
                            {
                                var textItem = questionnaireItem.Field<TextItemField>(item.FieldId.Value);
                                textItem.Value = item.StrAnswer;
                            }
                            else
                            {
                                var textItem = questionnaireItem.Field<TextItemField>(item.FieldId.Value);
                                textItem.Value = " ";
                            }
                            break;
                        case "category":
                            if (item.StrAnswer != null && item.StrAnswer != string.Empty)
                            {
                                var categoryItem = questionnaireItem.Field<CategoryItemField>(item.FieldId.Value);
                                categoryItem.OptionText = item.StrAnswer;
                            }
                            break;
                        case "date":
                            if (item.StrAnswer != null && item.StrAnswer != string.Empty)
                            {


                                DateTime dtValue, dtValue2;
                                if (DateTime.TryParse(item.StrAnswer, out dtValue))
                                {
                                    var dateField = questionnaireItem.Field<DateItemField>(item.FieldId.Value);
                                    dateField.Start = dtValue;
                                    if (item.StrAnswer2 != null && item.StrAnswer2 != string.Empty)
                                    {
                                        if (DateTime.TryParse(item.StrAnswer2, out dtValue2))
                                            dateField.End = dtValue2;
                                        else
                                            dateField.End = null;
                                    }
                                }

                            }
                            break;
                        case "app":
                            if (item.StrAnswer != null && item.StrAnswer != string.Empty)
                            {

                                switch (item.StrQuestion.ToLower())
                                {
                                    case string s when s.Contains("2. client name"):
                                        var searchClient = _soxContext.ClientSs.Where(x => x.Name.Equals(item.StrAnswer)).AsNoTracking().FirstOrDefault();
                                        if (searchClient != null)
                                        {
                                            List<int> listItem = new List<int>();
                                            listItem.Add(searchClient.PodioItemId);
                                            var appReference = questionnaireItem.Field<AppItemField>(item.FieldId.Value);
                                            appReference.ItemIds = listItem;                                            
                                        }
                                        break;

                                    case string s when s.Contains("key control using iuc"):
                                        var searchControlId = _soxContext.RcmControlId.Where(x => x.ControlId.Equals(item.StrAnswer)).AsNoTracking().FirstOrDefault();
                                        if (searchControlId != null)
                                        {
                                            List<int> listItem = new List<int>();
                                            listItem.Add(searchControlId.PodioItemId);
                                            var appReference = questionnaireItem.Field<AppItemField>(item.FieldId.Value);
                                            appReference.ItemIds = listItem;
                                        }
                                        break;

                                    case string s when s.Contains("control activity using iuc"):
                                        var searchControlAct = _soxContext.KeyReportControlActivity.Where(x => x.ControlActivity.Equals(item.StrAnswer)).AsNoTracking().FirstOrDefault();
                                        if (searchControlAct != null)
                                        {
                                            List<int> listItem = new List<int>();
                                            listItem.Add(searchControlAct.PodioItemId);
                                            var appReference = questionnaireItem.Field<AppItemField>(item.FieldId.Value);
                                            appReference.ItemIds = listItem;
                                        }
                                        break;

                                    case string s when s.Contains("key/non-key control"):
                                        var searchKey = _soxContext.KeyReportKeyControl.Where(x => x.Key.Equals(item.StrAnswer)).AsNoTracking().FirstOrDefault();
                                        if (searchKey != null)
                                        {
                                            List<int> listItem = new List<int>();
                                            listItem.Add(searchKey.PodioItemId);
                                            var appReference = questionnaireItem.Field<AppItemField>(item.FieldId.Value);
                                            appReference.ItemIds = listItem;
                                        }
                                        break;

                                    case string s when s.Contains("name of iuc"):
                                        var searchRepName = _soxContext.KeyReportName.Where(x => x.Name.Equals(item.StrAnswer)).AsNoTracking().FirstOrDefault();
                                        if (searchRepName != null)
                                        {
                                            List<int> listItem = new List<int>();
                                            listItem.Add(searchRepName.PodioItemId);
                                            var appReference = questionnaireItem.Field<AppItemField>(item.FieldId.Value);
                                            appReference.ItemIds = listItem;
                                        }
                                        break;

                                    case string s when s.Contains("report customized"):
                                        var searchRepCustom = _soxContext.KeyReportReportCustomized.Where(x => x.ReportCustomized.Equals(item.StrAnswer)).AsNoTracking().FirstOrDefault();
                                        if (searchRepCustom != null)
                                        {
                                            List<int> listItem = new List<int>();
                                            listItem.Add(searchRepCustom.PodioItemId);
                                            var appReference = questionnaireItem.Field<AppItemField>(item.FieldId.Value);
                                            appReference.ItemIds = listItem;
                                        }
                                        break;

                                    case string s when s.Contains("iuc/key report type"):
                                        var searchIucType = _soxContext.KeyReportIUCType.Where(x => x.IUCType.Equals(item.StrAnswer)).AsNoTracking().FirstOrDefault();
                                        if (searchIucType != null)
                                        {
                                            List<int> listItem = new List<int>();
                                            listItem.Add(searchIucType.PodioItemId);
                                            var appReference = questionnaireItem.Field<AppItemField>(item.FieldId.Value);
                                            appReference.ItemIds = listItem;
                                        }
                                        break;

                                    case string s when s.Contains("controls relying on iuc"):
                                        var searchControlRelIuc = _soxContext.KeyReportControlsRelyingIUC.Where(x => x.ControlsRelyingIUC.Equals(item.StrAnswer)).AsNoTracking().FirstOrDefault();
                                        if (searchControlRelIuc != null)
                                        {
                                            List<int> listItem = new List<int>();
                                            listItem.Add(searchControlRelIuc.PodioItemId);
                                            var appReference = questionnaireItem.Field<AppItemField>(item.FieldId.Value);
                                            appReference.ItemIds = listItem;
                                        }
                                        break;

                                    case string s when s.Contains("preparer"):
                                        var searchPreparer = _soxContext.KeyReportPreparer.Where(x => x.Preparer.Equals(item.StrAnswer)).AsNoTracking().FirstOrDefault();
                                        if (searchPreparer != null)
                                        {
                                            List<int> listItem = new List<int>();
                                            listItem.Add(searchPreparer.PodioItemId);
                                            var appReference = questionnaireItem.Field<AppItemField>(item.FieldId.Value);
                                            appReference.ItemIds = listItem;
                                        }
                                        break;

                                    case string s when s.Contains("unique key report"):
                                        var searchUniqueRep = _soxContext.KeyReportUniqueKeyReport.Where(x => x.UniqueKeyReport.Equals(item.StrAnswer)).AsNoTracking().FirstOrDefault();
                                        if (searchUniqueRep != null)
                                        {
                                            List<int> listItem = new List<int>();
                                            listItem.Add(searchUniqueRep.PodioItemId);
                                            var appReference = questionnaireItem.Field<AppItemField>(item.FieldId.Value);
                                            appReference.ItemIds = listItem;
                                        }
                                        break;

                                    case string s when s.Contains("28. notes"):
                                        var searchRepNotes = _soxContext.KeyReportNotes.Where(x => x.ReportNotes.Equals(item.StrAnswer)).AsNoTracking().FirstOrDefault();
                                        if (searchRepNotes != null)
                                        {
                                            List<int> listItem = new List<int>();
                                            listItem.Add(searchRepNotes.PodioItemId);
                                            var appReference = questionnaireItem.Field<AppItemField>(item.FieldId.Value);
                                            appReference.ItemIds = listItem;
                                        }
                                        break;

                                    case string s when s.Contains("key report number"):
                                        var searchKeyRepNum = _soxContext.KeyReportNumber.Where(x => x.ReportNumber.Equals(item.StrAnswer)).AsNoTracking().FirstOrDefault();
                                        if (searchKeyRepNum != null)
                                        {
                                            List<int> listItem = new List<int>();
                                            listItem.Add(searchKeyRepNum.PodioItemId);
                                            var appReference = questionnaireItem.Field<AppItemField>(item.FieldId.Value);
                                            appReference.ItemIds = listItem;
                                        }
                                        break;

                                    case string s when s.Contains("tester"):
                                        var searchRepTester = _soxContext.KeyReportTester.Where(x => x.Tester.Equals(item.StrAnswer)).AsNoTracking().FirstOrDefault();
                                        if (searchRepTester != null)
                                        {
                                            List<int> listItem = new List<int>();
                                            listItem.Add(searchRepTester.PodioItemId);
                                            var appReference = questionnaireItem.Field<AppItemField>(item.FieldId.Value);
                                            appReference.ItemIds = listItem;
                                        }
                                        break;

                                    case string s when s.Contains("reviewer") || s.Contains("2nd reviewer"):
                                        var searchReviewer = _soxContext.KeyReportReviewer.Where(x => x.Reviewer.Equals(item.StrAnswer)).AsNoTracking().FirstOrDefault();
                                        if (searchReviewer != null)
                                        {
                                            List<int> listItem = new List<int>();
                                            listItem.Add(searchReviewer.PodioItemId);
                                            var appReference = questionnaireItem.Field<AppItemField>(item.FieldId.Value);
                                            appReference.ItemIds = listItem;
                                        }
                                        break;

                                    case string s when s.Contains("sub-process"):
                                        var searchRcmProcess = _soxContext.RcmSubProcess.Where(x => x.SubProcess.Equals(item.StrAnswer)).AsNoTracking().FirstOrDefault();
                                        if (searchRcmProcess != null)
                                        {
                                            List<int> listItem = new List<int>();
                                            listItem.Add(searchRcmProcess.PodioItemId);
                                            var appReference = questionnaireItem.Field<AppItemField>(item.FieldId.Value);
                                            appReference.ItemIds = listItem;
                                        }
                                        break;

                                    case string s when s.Contains("process owner"):
                                        var searchRcmControlOwner = _soxContext.RcmControlOwner.Where(x => x.ControlOwner.Equals(item.StrAnswer)).AsNoTracking().FirstOrDefault();
                                        if (searchRcmControlOwner != null)
                                        {
                                            List<int> listItem = new List<int>();
                                            listItem.Add(searchRcmControlOwner.PodioItemId);
                                            var appReference = questionnaireItem.Field<AppItemField>(item.FieldId.Value);
                                            appReference.ItemIds = listItem;
                                        }
                                        break;

                                }
                            }
                            break;
                        case "image":
                            if (item.StrAnswer != null && item.StrAnswer != string.Empty)
                            {
                                //// Upload file
                                //string startupPath = Directory.GetCurrentDirectory();
                                //string filename = item.StrAnswer;
                                //string path = Path.Combine(startupPath, "include", "upload", "image", filename);
                                //var uploadedFile = await podio.FileService.UploadFile(path, filename);

                                //// Set FileIds
                                //if (uploadedFile != null)
                                //{
                                //    //Debug.WriteLine($"uploadedFile: {uploadedFile.Result.FileId}");
                                //    ImageItemField imageField = questionnaireItem.Field<ImageItemField>(item.FieldId.Value);
                                //    imageField.FileIds = new List<int> { uploadedFile.FileId };
                                //}

                            }
                            break;
                        case "duration":
                            if (item.StrAnswer != null && item.StrAnswer != string.Empty)
                            {
                                TimeSpan ts;
                                if (TimeSpan.TryParse(item.StrAnswer, CultureInfo.CurrentCulture, out ts))
                                {
                                    DurationItemField durationField = questionnaireItem.Field<DurationItemField>(item.FieldId.Value);
                                    durationField.Value = ts;
                                }
                            }
                            break;
                    }
                }

            }
            return questionnaireItem;
        }


        //end KeyReport Tracker

        private Dictionary<string, string> GetMimeTypes()
        {
            return new Dictionary<string, string>
            {
                {".txt", "text/plain"},
                {".pdf", "application/pdf"},
                {".doc", "application/vnd.ms-word"},
                {".docx", "application/vnd.ms-word"},
                {".xls", "application/vnd.ms-excel"},
                {".xlsx", "application/vnd.ms-excel"},
                {".xlsm", "application/vnd.ms-excel.addin.macroEnabled.12"},
                {".png", "image/png"},
                {".jpg", "image/jpeg"},
            };
        }

        private List<int> SearchItemId(KeyReportFilter filter, string appId)
        {
            List<int> listItemIdFY = new List<int>();
            List<int> listItemIdClient = new List<int>();
            List<int> listItemIdReportName = new List<int>();
            List<int> listItemIdControlId = new List<int>();

            //item id for FY
            var itemIdFY = _soxContext.KeyReportUserInput.Where(x =>
                    x.StrAnswer.ToLower().Equals(filter.FY)
                    && x.AppId.Equals(appId)
                )
                .AsNoTracking()
                .Select(x => x.ItemId)
                .Distinct()
                .ToList();
            if (itemIdFY != null)
            {
                listItemIdFY = itemIdFY;
            }

            //item id for client
            var itemIdClient = _soxContext.KeyReportUserInput.Where(x =>
                    x.StrAnswer.ToLower().Equals(filter.ClientName)
                    && x.AppId.Equals(appId)
                )
                .AsNoTracking()
                .Select(x => x.ItemId)
                .Distinct()
                .ToList();
            if (itemIdClient != null)
            {
                listItemIdClient = itemIdClient;
            }

            var itemIdReportName = _soxContext.KeyReportUserInput.Where(x =>
                    x.StrAnswer.ToLower().Equals(filter.KeyReportName)
                    && x.AppId.Equals(appId)
                )
                .AsNoTracking()
                .Select(x => x.ItemId)
                .Distinct()
                .ToList();
            if (itemIdReportName != null)
            {
                listItemIdReportName = itemIdReportName;
            }

            var itemIdControlId = _soxContext.KeyReportUserInput.Where(x =>
                    x.StrAnswer.ToLower().Equals(filter.ControlId)
                    && x.AppId.Equals(appId)
                )
                .AsNoTracking()
                .Select(x => x.ItemId)
                .Distinct()
                .ToList();
            if (itemIdControlId != null)
            {
                listItemIdControlId = itemIdControlId;
            }

            //var itemId = listItemIdFY.Intersect(listItemIdClient).ToList();

            //HashSet<int> hashSet = new HashSet<int>(listItemIdFY);
            //hashSet.IntersectWith(listItemIdClient);
            //hashSet.IntersectWith(listItemIdReportName);
            //hashSet.IntersectWith(listItemIdControlId);
            //List<int> intersection = hashSet.ToList();

            /*  disabled for testing mark
            List<int> intersection = new List<int>();
            foreach (int i in IntersectAllIfEmpty(listItemIdFY, listItemIdClient, listItemIdReportName, listItemIdControlId))
            {
                Debug.WriteLine(i);
                intersection.Add(i);
            }
            */



            return itemIdReportName;//"intersection";
        }

        private List<int> SearchItemId2(KeyReportFilter filter, string appId)
        {
            List<int> listItemId = new List<int>();


            var itemIdReportName = _soxContext.KeyReportUserInput.Where(x =>
                    x.AppId.Equals(appId)
                    && x.TagReportName.ToLower().Equals(filter.KeyReportName)
                    && x.TagControlId.ToLower().Equals(filter.ControlId)
                    && x.TagClientName.ToLower().Equals(filter.ClientName)
                    && x.TagFY.ToLower().Equals(filter.FY)
                    && x.TagStatus.ToLower() != "inactive"
                )
                .AsNoTracking()
                .Select(x => x.ItemId)
                .Distinct()
                .ToList();
            if (itemIdReportName != null)
            {
                foreach (var item1 in itemIdReportName)
                {
                    if (!listItemId.Contains(item1))
                        listItemId.Add(item1);
                }
            }

            return listItemId;
        }

        private List<KeyReportUserInput> GetKeyReportItem(int itemId) 
        {
            List<KeyReportUserInput> listKeyReportItem = new List<KeyReportUserInput>();
            var itemIdFY = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(itemId)).OrderBy(x => x.Position).AsNoTracking().ToList();
            if (itemIdFY != null)
            {
                listKeyReportItem = itemIdFY;
            }

            return listKeyReportItem;
        }


        /* disabled for testing
        public static IEnumerable<T> IntersectAllIfEmpty<T>(params IEnumerable<T>[] lists)
        {
            IEnumerable<T> results = null;

            lists = lists.Where(l => l.Any()).ToArray();

            if (lists.Length > 0)
            {
                results = lists[0];

                for (int i = 1; i < lists.Length; i++)
                    results = results.Intersect(lists[i]);
            }
            else
            {
                results = new T[0];
            }

            return results;
        }
        */
        private void AddExcelImage(ExcelWorksheet ws, 
            string path, 
            string clientName, 
            string reportName, 
            List<string> listImage,
            int startRow)
        {
            int row = startRow;
           
            foreach (var image in listImage.OrderBy(x => x))
            {
                //var dimension = GetImageDimension(clientName, reportName, image);
                //string imagepath = Path.Combine(path, $"{clientname}", $"{reportname}", image);
                string imagepath = Path.Combine(path, image);
                FileInfo file = new FileInfo(imagepath);
                if (file != null && file.Exists)
                {
                    Bitmap img = new Bitmap(imagepath);
                    int height = img.Height;
                    int width = img.Width;
                    Guid guid = Guid.NewGuid();
                    //var screenshot = ws.Drawings.AddPicture(image, file);
                    var screenshot = ws.Drawings.AddPicture(guid.ToString(), file);
                    screenshot.SetSize(width, height);
                    screenshot.SetPosition(row, 0, 1, 0);
                    row += 56; //1920 x 1080
                }
            }
           
        }

        private (int, int) GetImageDimension(string clientName, string reportName, string fileName)
        {
            string startupPath = Directory.GetCurrentDirectory();
            string path = Path.Combine(startupPath, "include", "sharefile", "download", $"{clientName}", $"{reportName}", fileName);
            int height = 0;
            int width = 0;
            FileInfo file = new FileInfo(path);
            if (file != null && file.Exists)
            {
                Bitmap img = new Bitmap(path);
                height = img.Height;
                width = img.Width;
            }
            return (height, width);
        }

        private static string GetColumnName(int index)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            var value = "";

            if (index >= letters.Length)
                value += letters[index / letters.Length - 1];

            value += letters[index % letters.Length];

            return value;
        }

        private void LeadsheetOutputFileTab3(ExcelWorksheet ws2, KeyReportScreenshot filter, List<KeyReportUserInput> keyReportItem, string sbControlId)
        {
            ExcelService xlsService = new ExcelService();
            //Tab3 (KeyReportName) - 
            ws2.View.ZoomScale = 80;
            //disable grid
            ws2.View.ShowGridLines = false;

            //set column width
            ws2.Column(1).Width = 16.43;
            ws2.Column(2).Width = 20;
            ws2.Column(3).Width = 20;
            ws2.Column(4).Width = 20;
            ws2.Column(5).Width = 20;

            ws2.Cells[1, 1].Value = "Key Report:";
            ws2.Cells[2, 1].Value = "Control #:";
            ws2.Cells[3, 1].Value = "Provided By:";
            ws2.Cells[4, 1].Value = "Tested by:";
            ws2.Cells[5, 1].Value = "Description:";

            ws2.Cells[1, 2].Value = filter.Filter.KeyReportName; //"Key Report:";
            ws2.Cells[2, 2].Value = sbControlId; //"Control #:";
            ws2.Cells[3, 2].Value = //"Provided By:";
            ws2.Cells[4, 2].Value = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("14. preparer")).Select(x => x.StrAnswer).FirstOrDefault(); //"Tested by:";
            ws2.Cells[5, 2].Value = //"Description:";
            ws2.Cells[5, 2].Value = $"(Excel / Pdf / Text) Output of the {filter.Filter.KeyReportName}";

            xlsService.ExcelSetBackgroundColorDarkBlue(ws2, 1, 1, 5, 1);
            xlsService.ExcelSetFontColorWhite(ws2, 1, 1, 5, 1);

            xlsService.ExcelSetBorder(ws2, 1, 1, 5, 5);
            xlsService.ExcelSetArialSize10(ws2, 1, 1, 10, 14);

            ws2.Cells["B" + 1 + ":E" + 1].Merge = true;
            ws2.Cells["B" + 2 + ":E" + 2].Merge = true;
            ws2.Cells["B" + 3 + ":E" + 3].Merge = true;
            ws2.Cells["B" + 4 + ":E" + 4].Merge = true;
            ws2.Cells["B" + 5 + ":E" + 5].Merge = true;
            xlsService.ExcelSetFontBold(ws2, 1, 1, 5, 1);
        }

        private void LeadsheetOutputFileTab4(ExcelWorksheet ws3, KeyReportScreenshot filter, List<KeyReportUserInput> keyReportItem, string sbControlId, string screenshotPath)
        {
            ExcelService xlsService = new ExcelService();
            //Tab4 (3. Sample Testing_Completeness) - 
            
            ws3.View.ZoomScale = 80;
            //disable grid
            ws3.View.ShowGridLines = false;
            //set column width
            ws3.Column(1).Width = 16.43;
            ws3.Column(2).Width = 20;
            ws3.Column(3).Width = 20;
            ws3.Column(4).Width = 20;
            ws3.Column(5).Width = 20;

            ws3.Cells[1, 1].Value = "Key Report:";
            ws3.Cells[2, 1].Value = "Control #:";
            ws3.Cells[3, 1].Value = "Provided By:";
            ws3.Cells[4, 1].Value = "Tested by:";
            ws3.Cells[5, 1].Value = "Description:";

            ws3.Cells[1, 2].Value = filter.Filter.KeyReportName; //"Key Report:";
            ws3.Cells[2, 2].Value = sbControlId; //"Control #:";
            ws3.Cells[3, 2].Value = //"Provided By:";
            ws3.Cells[4, 2].Value = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("14. preparer")).Select(x => x.StrAnswer).FirstOrDefault(); //"Tested by:";
            ws3.Cells[5, 2].Value = //"Description:";
            ws3.Cells[5, 2].Value = "Comparison Between (System) Screenshot and (Excel/Pdf/Text) Output";


            ws3.Cells[11, 1].Value = "From screenshots";

            xlsService.ExcelSetBackgroundColorDarkBlue(ws3, 1, 1, 5, 1);
            xlsService.ExcelSetFontColorWhite(ws3, 1, 1, 5, 1);
            xlsService.ExcelSetFontColorRed(ws3, 11, 1, 11, 1);


            //xlsService.ExcelSetFontBold(ws3, 1, 1, 5, 1);

            xlsService.ExcelSetArialSize10(ws3, 1, 1, 11, 26);


            xlsService.ExcelSetBorder(ws3, 1, 1, 5, 5);
            ws3.Cells[11, 26].Value = "From report";
            ws3.Cells["B" + 1 + ":E" + 1].Merge = true;
            ws3.Cells["B" + 2 + ":E" + 2].Merge = true;
            ws3.Cells["B" + 3 + ":E" + 3].Merge = true;
            ws3.Cells["B" + 4 + ":E" + 4].Merge = true;
            ws3.Cells["B" + 5 + ":E" + 5].Merge = true;

            xlsService.ExcelSetFontBold(ws3, 1, 1, 5, 1);
            xlsService.ExcelSetFontBold(ws3, 11, 1, 11, 26);
            xlsService.ExcelSetFontColorRed(ws3, 11, 14, 11, 26);

            //end of Tab4
        }

        private void LeadsheetOutputFileTab5(ExcelWorksheet ws4, KeyReportScreenshot filter, List<KeyReportUserInput> keyReportItem, string sbControlId, string screenshotPath)
        {
            //Tab5 (4. Sample Testing_Accuracy) - 
            ExcelService xlsService = new ExcelService();
            ws4.View.ZoomScale = 80;
            ws4.View.ShowGridLines = false;

            //set column width
            ws4.Column(1).Width = 16.43;
            ws4.Column(2).Width = 20;
            ws4.Column(3).Width = 20;
            ws4.Column(4).Width = 20;
            ws4.Column(5).Width = 20;

            ws4.Cells[1, 1].Value = "Key Report:";
            ws4.Cells[2, 1].Value = "Control #:";
            ws4.Cells[3, 1].Value = "Provided By:";
            ws4.Cells[4, 1].Value = "Tested by:";
            ws4.Cells[5, 1].Value = "Description:";

            ws4.Cells[1, 2].Value = filter.Filter.KeyReportName; //"Key Report:";
            ws4.Cells[2, 2].Value = sbControlId; //"Control #:";
            ws4.Cells[3, 2].Value = //"Provided By:";
            ws4.Cells[4, 2].Value = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("14. preparer")).Select(x => x.StrAnswer).FirstOrDefault(); //"Tested by:";
            ws4.Cells[5, 2].Value = //"Description:";
            ws4.Cells[5, 2].Value = "Comparison Between (System) Screenshot and (Excel/Pdf/Text) Output";

            ws4.Cells[11, 1].Value = "From screenshots";


            xlsService.ExcelSetBackgroundColorDarkBlue(ws4, 1, 1, 5, 1);
            xlsService.ExcelSetFontColorWhite(ws4, 1, 1, 5, 1);

            xlsService.ExcelSetFontColorRed(ws4, 11, 1, 11, 1);

            xlsService.ExcelSetArialSize10(ws4, 1, 1, 11, 26);

            xlsService.ExcelSetBorder(ws4, 1, 1, 5, 5);
            ws4.Cells[11, 26].Value = "From report";
            ws4.Cells["B" + 1 + ":E" + 1].Merge = true;
            ws4.Cells["B" + 2 + ":E" + 2].Merge = true;
            ws4.Cells["B" + 3 + ":E" + 3].Merge = true;
            ws4.Cells["B" + 4 + ":E" + 4].Merge = true;
            ws4.Cells["B" + 5 + ":E" + 5].Merge = true;
            xlsService.ExcelSetFontBold(ws4, 1, 1, 5, 1);
            xlsService.ExcelSetFontBold(ws4, 11, 1, 11, 26);
            xlsService.ExcelSetFontColorRed(ws4, 11, 14, 11, 26);

           
        }

        [AllowAnonymous]
        [HttpPost("methods")]
        public IActionResult fetch_methods([FromBody] String type) 
        {

            try
            {
                var checkMethods = _soxContext.CAMethodLibrary.Where(x =>
                    x.MethodType.Equals(type)).OrderBy(x => x.MethodName);
                if (checkMethods != null)
                {
                    return Ok(checkMethods.ToList());
                }
                else
                    return NoContent();
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetMethods {ex}", "ErrorGetMethods");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetMethods");
                if (_environment.IsDevelopment())
                {
                    return BadRequest(ex.ToString());
                }
                else
                {
                    return BadRequest();
                }
            }
        }

        [HttpPost("fetch/parameters")]
        public IActionResult fetch_parameters([FromBody] KeyReportQuestionsFilter filter)
        {
            try
            {
                List<ParametersLibrary> parametersLibrary = new List<ParametersLibrary>();
                var checkMethods = _soxContext.CAMethodLibrary.Where(x =>
                    x.Id.Equals(filter.method)).OrderBy(x => x.Id);
                if (checkMethods != null)
                {
                    List<CAMethodLibrary> methodList = new List<CAMethodLibrary>();
                    methodList = checkMethods.ToList();
                    //methodList = checkMethods.ToList();
                    foreach (var item in methodList)
                    {
                        var checkQuestions = _soxContext.ParametersLibrary.Where(x =>
                            x.Method.Equals(item.MethodName) && 
                            x.KeyReportName.Contains(filter.reportName) &&
                            x.ClientName.Equals(filter.clientName)).OrderBy(x => x.Id);
                        if (checkQuestions != null)
                        {
                            foreach (var quest in checkQuestions)
                            {
                                parametersLibrary.Add(quest);
                            }

                        }
                       
                    }
                    return Ok(parametersLibrary);
                    //return Ok(checkMethods.ToList());
                    //query for paramters questions
                }
                else
                    return NoContent();
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetParametersQuestion {ex}", "ErrorGetParametersQuestion");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetParametersQuestion");
                if (_environment.IsDevelopment())
                {
                    return BadRequest(ex.ToString());
                }
                else
                {
                    return BadRequest();
                }
            }
        }
        
        [HttpPost("fetch/reports")]
        public IActionResult fetch_reports([FromBody] KeyReportQuestionsFilter filter)
        {
            try
            {
                List<ReportsLibrary> reportsLibrary = new List<ReportsLibrary>();
                var checkMethods = _soxContext.CAMethodLibrary.Where(x =>
                    x.Id.Equals(filter.method)).OrderBy(x => x.Id);
                if (checkMethods != null)
                {
                    List<CAMethodLibrary> methodList = new List<CAMethodLibrary>();
                    methodList = checkMethods.ToList();
                    //methodList = checkMethods.ToList();
                    foreach (var item in methodList)
                    {
                        var checkQuestions = _soxContext.ReportsLibrary.Where(x =>
                            x.Method.Equals(item.MethodName) &&
                            x.KeyReportName.Contains(filter.reportName) &&
                            x.ClientName.Equals(filter.clientName)).OrderBy(x => x.Id);
                        if (checkQuestions != null)
                        {
                            foreach (var quest in checkQuestions)
                            {
                                reportsLibrary.Add(quest);
                            }

                        }

                    }
                    return Ok(reportsLibrary);
                }
                else
                    return NoContent();
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetReportsQuestion {ex}", "ErrorGetReportsQuestion");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetReportsQuestion");
                if (_environment.IsDevelopment())
                {
                    return BadRequest(ex.ToString());
                }
                else
                {
                    return BadRequest();
                }
            }
        }
        
        [HttpPost("fetch/completeness")]
        public IActionResult fetch_completeness([FromBody] KeyReportQuestionsFilter filter)
        {

            try
            {
                List<CompletenessLibrary> completenessLibrary = new List<CompletenessLibrary>();
                var checkMethods = _soxContext.CAMethodLibrary.Where(x =>
                    x.Id.Equals(filter.method)).OrderBy(x => x.Id);
                if (checkMethods != null)
                {
                    List<CAMethodLibrary> methodList = new List<CAMethodLibrary>();
                    methodList = checkMethods.ToList();
                    //methodList = checkMethods.ToList();
                    foreach (var item in methodList)
                    {
                        var checkQuestions = _soxContext.CompletenessLibrary.Where(x =>
                            x.Method.Equals(item.MethodName) &&
                            x.KeyReportName.Contains(filter.reportName) &&
                            x.ClientName.Equals(filter.clientName)).OrderBy(x => x.Id);
                        if (checkQuestions != null)
                        {
                            foreach (var quest in checkQuestions)
                            {
                                completenessLibrary.Add(quest);
                            }

                        }

                    }
                    return Ok(completenessLibrary);
                }
                else
                    return NoContent();
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetCompletenessQuestion {ex}", "ErrorGetCompletenessQuestion");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetCompletenessQuestion");
                if (_environment.IsDevelopment())
                {
                    return BadRequest(ex.ToString());
                }
                else
                {
                    return BadRequest();
                }
            }
        }

        [HttpPost("fetch/accuracy")]
        public IActionResult fetch_accuracy ([FromBody] KeyReportQuestionsFilter filter)
        {
            try
            {
                List<AccuracyLibrary> accuracyLibrary = new List<AccuracyLibrary>();
                var checkMethods = _soxContext.CAMethodLibrary.Where(x =>
                    x.Id.Equals(filter.method)).OrderBy(x => x.Id);
                if (checkMethods != null)
                {
                    List<CAMethodLibrary> methodList = new List<CAMethodLibrary>();
                    methodList = checkMethods.ToList();
                    //methodList = checkMethods.ToList();
                    foreach (var item in methodList)
                    {
                        var checkQuestions = _soxContext.AccuracyLibrary.Where(x =>
                            x.Method.Equals(item.MethodName) &&
                            x.KeyReportName.Contains(filter.reportName) &&
                            x.ClientName.Equals(filter.clientName)).OrderBy(x => x.Id);
                        if (checkQuestions != null)
                        {
                            foreach (var quest in checkQuestions)
                            {
                                accuracyLibrary.Add(quest);
                            }
                        }
                    }
                    return Ok(accuracyLibrary);
                }
                else
                    return NoContent();
            }
            catch (Exception ex)
            {
                FileLog.Write($"Error GetAccuracyQuestion {ex}", "ErrorGetAccuracyQuestion");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetCompletenessQuestion");
                if (_environment.IsDevelopment())
                {
                    return BadRequest(ex.ToString());
                }
                else
                {
                    return BadRequest();
                }
            }
        }

        [AllowAnonymous]
        [HttpPost("leadsheet/fetch")]
        public IActionResult GenerateKeyReportViewAccess([FromBody] KeyReportScreenshot filter)
        {
            try
            {
                String[] answers = new String[31];
                List<int> listItemIdAll = new List<int>();
                List<int> listItemIdUnique = new List<int>();
                string appId = _config.GetSection("KeyReportApp").GetSection("AllIUCId").Value;
                string appIdTestStat = _config.GetSection("KeyReportApp").GetSection("TestStatusTrackerId").Value;

                var inputItemId = _soxContext.KeyReportUserInput.Where(x =>
                                       x.TagFY.ToLower().Equals(filter.Filter.FY.ToLower())
                                       && x.TagClientName.ToLower().Equals(filter.Filter.ClientName.ToLower())
                                       && x.TagReportName.ToLower().Equals(filter.Filter.KeyReportName.ToLower())
                                       && x.AppId.Equals(appId)
                                       && x.TagStatus.ToLower() != "inactive"
                                   )
                                   .AsNoTracking()
                                   .Select(x => x.ItemId)
                                   .Distinct()
                                   .ToList();

                foreach (var itemId in inputItemId)
                {
                    if (!listItemIdAll.Contains(itemId) && itemId != 0)
                    {
                        listItemIdAll.Add(itemId);
                    }
                }
                if (listItemIdAll != null && listItemIdAll.Count > 0)
                {
                    //Get unique itemid
                    foreach (var item in listItemIdAll)
                    {
                        var inputItemId4 = _soxContext.KeyReportUserInput.Where(x =>
                                                x.ItemId.Equals(item)
                                                && x.StrQuestion.ToLower().Contains("unique key report")
                                                && x.StrAnswer.ToLower().Equals("yes")
                                                && x.TagStatus.ToLower() != "inactive"
                                            )
                                            .AsNoTracking()
                                            .Select(x => new { x.ItemId, x.TagControlId })
                                            .Distinct()
                                            .ToList();

                        foreach (var itemUnique in inputItemId4)
                        {
                            if (!listItemIdUnique.Contains(itemUnique.ItemId) && itemUnique.ItemId != 0)
                            {
                                listItemIdUnique.Add(itemUnique.ItemId);
                                filter.Filter.ControlId = itemUnique.TagControlId;
                            }
                        }
                    }

                    var testingStatusItemId = _soxContext.KeyReportUserInput.Where(x =>
                                                    x.TagClientName.Equals(filter.Filter.ClientName)
                                                    && x.TagFY.Equals(filter.Filter.FY)
                                                    && x.TagReportName.Equals(filter.Filter.KeyReportName)
                                                    && x.TagControlId.Equals(filter.Filter.ControlId)
                                                    && x.TagStatus.ToLower() != "inactive"
                                                    && x.AppId.Equals(appIdTestStat)
                                                    && x.StrQuestion.ToLower().Equals("unique key report")
                                                    && x.StrAnswer.ToLower().Equals("yes")
                                                )
                                                .AsNoTracking()
                                                .Select(x => x.ItemId)
                                                .Distinct()
                                                .FirstOrDefault();

                    List<KeyReportUserInput> testingStatus = new List<KeyReportUserInput>();
                    if (testingStatusItemId != 0)
                    {
                        testingStatus = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(testingStatusItemId))
                                                .AsNoTracking()
                                                .Distinct()
                                                .ToList();
                    }

                    //Processing excel output
                    if (listItemIdUnique != null && listItemIdUnique.Count > 0)
                    {
                        var sampleProc_parameter = "";
                        var sampleProc_report = "";
                        var sampleProc_completeness = "";
                        var sampleProc_accuracy = "";
                        foreach (var itemId in listItemIdUnique)
                        {
                            //Get All IUC
                            var keyReportItem = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(itemId)).ToList();

                            StringBuilder sbPurpose = new StringBuilder();
                            List<string> listPurpose = new List<string>();
                            int count = 0;

                            //consolidate value from key and non-key and store in list
                            if (listItemIdAll != null && listItemIdAll.Count > 0)
                            {
                                //here we loop through all the key and non-key value
                                foreach (var item in listItemIdAll)
                                {
                                    var leadsheetItem = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(item)).ToList();
                                    if (leadsheetItem != null && leadsheetItem.Count > 0)
                                    {
                                        //Debug.WriteLine($"ItemId : {item}");
                                        var leadsheetPurpose = leadsheetItem.Where(x => x.StrQuestion.ToLower().Contains("what is the purpose of the report")).Select(x => x.StrAnswer).FirstOrDefault();
                                        if (leadsheetPurpose != null)
                                        {
                                            //Debug.WriteLine($"Purpose : {purpose}");
                                            if (!listPurpose.Contains(leadsheetPurpose))
                                                listPurpose.Add(leadsheetPurpose);
                                        }
                                    }
                                }
                            }

                            //process list of consolidated data from key and non key
                            if (listPurpose.Count > 0)
                            {
                                count = 0;
                                foreach (var item in listPurpose)
                                {
                                    sbPurpose.Append($"{GetColumnName(count)}: {item}");
                                    //sbPurpose.Append(Environment.NewLine);
                                    count++;
                                }
                            }



                            var controlId = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("key control using iuc")).Select(x => x.StrAnswer).FirstOrDefault();

                            var controlFrequency = _soxContext.Rcm.Where(x =>
                                                            x.FY.ToLower().Equals(filter.Filter.FY.ToLower())
                                                            && x.ClientName.ToLower().Equals(filter.Filter.ClientName.ToLower())
                                                            && x.ControlId.ToLower().Equals(controlId.ToLower()))
                                                        .Select(x => x.ControlFrequency)
                                                        .AsNoTracking()
                                                        .FirstOrDefault();


                            answers[0] = filter.Filter.KeyReportName;
                            answers[1] = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("was this report previously evaluated for reliability")).Select(x => x.StrAnswer).FirstOrDefault(); //"Was this report previously evaluated for reliability? (If so, link to those procedures)";
                            answers[2] = sbPurpose.ToString(); //"What is the purpose of the report?";
                            answers[3] = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("what are the key data fields used from this report")).Select(x => x.StrAnswer).FirstOrDefault();
                            answers[4] = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("iuc/key report type")).Select(x => x.StrAnswer).FirstOrDefault();
                            answers[5] = string.Empty; //"Define a condition that would represent an exception on this report";
                            answers[6] = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("key report system name")).Select(x => x.StrAnswer).FirstOrDefault(); //"What is the application/system name? (source of the data used to generate the report)";
                            answers[7] = string.Empty; //keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("key report type")).Select(x => x.StrAnswer).FirstOrDefault(); //"Is this report: custom, canned, custom query?";
                            answers[8] = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("when was the report last modified")).Select(x => x.StrAnswer).FirstOrDefault(); //"When was the last time this report was modified?";
                            answers[9] = string.Empty; //keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("are parameters")).Select(x => x.StrAnswer).FirstOrDefault(); //"What are the key report parameters?";
                            answers[10] = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("are parameters input into the report each time it is run")).Select(x => x.StrAnswer).FirstOrDefault(); //"Are parameters input into the report each time it is run?";
                            answers[11] = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("how does the report user verify the report is complete")).Select(x => x.StrAnswer).FirstOrDefault();
                            answers[12] = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("how does the report user verify the report is accurate")).Select(x => x.StrAnswer).FirstOrDefault();
                            answers[13] = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("how does the report user verify the report data has integrity")).Select(x => x.StrAnswer).FirstOrDefault();
                            var overrideSampleProc = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("override sample procedures")).Select(x => x.StrAnswer).FirstOrDefault();                                                                                                                                                                                         //ws.Cells[17, 3].Value = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("what are the procedures performed to assess the accuracy, completeness, and validity of the source data")).Select(x => x.StrAnswer).FirstOrDefault(); //"Procedures performed to assess the accuracy, completeness, and validity of the source data.";
                            var completenessMethods = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("key report completeness methods")).Select(x => x.StrAnswer).FirstOrDefault(); //"Procedures performed to assess the accuracy, completeness, and validity of the source data.";
                            var completenessAnswer = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("key report completeness answer")).Select(x => x.StrAnswer).FirstOrDefault();
                            var accuracyMethods = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("key report accuracy methods")).Select(x => x.StrAnswer).FirstOrDefault();
                            var accuracyAnswer = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("key report accuracy answer")).Select(x => x.StrAnswer).FirstOrDefault();
                            var parameterMethods = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("key report parameter methods")).Select(x => x.StrAnswer).FirstOrDefault();
                            var parameterAnswer = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("key report parameter answer")).Select(x => x.StrAnswer).FirstOrDefault();
                            var reportMethods = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("key report report methods")).Select(x => x.StrAnswer).FirstOrDefault();
                            var reportAnswer = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("key report report answer")).Select(x => x.StrAnswer).FirstOrDefault();

                            //string[] a1, a2, a3, a4, a5, b1, b2, b3, b4, b5, c1, c2, c3, c4, c5, d1, d2, d3, d4, d5, d6;

                            string[] parameterMethodArray = !string.IsNullOrEmpty(parameterMethods) ? parameterMethods.Split(";") : null;
                            string[] reportMethodArray = !string.IsNullOrEmpty(reportMethods) ? reportMethods.Split(";") : null;
                            string[] completenessMethodArray = !string.IsNullOrEmpty(completenessMethods) ? completenessMethods.Split(";") : null;
                            string[] accuracyMethodArray = !string.IsNullOrEmpty(accuracyMethods) ? accuracyMethods.Split(";") : null;
                            string[] parameterAnswerArray = !string.IsNullOrEmpty(parameterAnswer) ? parameterAnswer.Split("////") : null;
                            string[] reportAnswerArray = !string.IsNullOrEmpty(reportAnswer) ? reportAnswer.Split("////") : null;
                            string[] completenessAnswerArray = !string.IsNullOrEmpty(completenessAnswer) ? completenessAnswer.Split("////") : null;
                            string[] accuracyAnswerArray = !string.IsNullOrEmpty(accuracyAnswer) ? accuracyAnswer.Split("////") : null;
                            int counter = 0;
                            if (!string.IsNullOrEmpty(overrideSampleProc))
                            {

                                answers[14] = overrideSampleProc;
                            }
                            else
                            {
                                counter = 0;
                                if (parameterMethodArray != null)
                                {
                                    foreach (var method in parameterMethodArray)
                                    {
                                        if (method != null && method != string.Empty)
                                        {

                                            int result = Int32.Parse(method);
                                            var tempMethod = _soxContext.CAMethodLibrary.Where(x => x.Id.Equals(result)).FirstOrDefault();
                                            if (tempMethod != null)
                                            {
                                                var parameter_lib = _soxContext.ParametersLibrary.Where(x =>
                                                                        x.KeyReportName.Equals(filter.Filter.KeyReportName)
                                                                        && x.Method.Equals(tempMethod.MethodName)).FirstOrDefault();
                                                if (parameter_lib != null)
                                                {
                                                    if (counter == 0 && sampleProc_parameter == "")
                                                    {
                                                        sampleProc_parameter += "A. ";
                                                    }
                                                    if (parameterAnswerArray != null && parameterAnswerArray.Any())
                                                    {
                                                        sampleProc_parameter += parameter_lib.Parameter + "<br/>";
                                                        sampleProc_parameter = sampleProc_parameter.Replace("<p>", string.Empty)
                                                                                                    .Replace("</p>", string.Empty)
                                                                                                    .Replace("</strong", string.Empty)
                                                                                                    .Replace("<strong class=\"text-bold\">", string.Empty);


                                                        if (parameterAnswerArray[counter] != null)
                                                        {
                                                            string[] tempParamAnswers = parameterAnswerArray[counter].Split(";");
                                                            sampleProc_parameter = sampleProc_parameter.Replace("(A1)", tempParamAnswers[0])
                                                                                                    .Replace("(A2)", tempParamAnswers[1])
                                                                                                    .Replace("(A3)", tempParamAnswers[2])
                                                                                                    .Replace("(A4)", tempParamAnswers[3])
                                                                                                    .Replace("(A5)", tempParamAnswers[4])
                                                                                                    .Replace("(A6)", tempParamAnswers[5])
                                                                                                    .Replace("(A7)", tempParamAnswers[6])
                                                                                                    .Replace("(A8)", tempParamAnswers[7])
                                                                                                    .Replace("(A9)", tempParamAnswers[8])
                                                                                                    .Replace("(A10)", tempParamAnswers[9]);

                                                        }
                                                    }

                                                }
                                            }
                                            counter++;
                                        }

                                    }
                                }
                                counter = 0;
                                if (reportMethodArray != null)
                                {
                                    foreach (var method in reportMethodArray)
                                    {
                                        if (method != null && method != string.Empty && method != "")
                                        {
                                            int result = Int32.Parse(method);
                                            var tempMethod = _soxContext.CAMethodLibrary.Where(x => x.Id.Equals(result)).FirstOrDefault();
                                            if (tempMethod != null)
                                            {
                                                var report_lib = _soxContext.ReportsLibrary.Where(x =>
                                                                        x.KeyReportName.Equals(filter.Filter.KeyReportName)
                                                                        && x.Method.Equals(tempMethod.MethodName)).FirstOrDefault();
                                                if (report_lib != null)
                                                {
                                                    if (counter == 0 && sampleProc_report == "")
                                                    {
                                                        sampleProc_report += "B. ";
                                                    }
                                                    sampleProc_report += report_lib.Report + "<br/>";
                                                    sampleProc_report = sampleProc_report.Replace("<p>", string.Empty)
                                                                                            .Replace("</p>", string.Empty)
                                                                                            .Replace("</strong", string.Empty)
                                                                                            .Replace("<strong class=\"text-bold\">", string.Empty);

                                                    if (reportAnswerArray[counter] != null && reportAnswerArray[counter] != "")
                                                    {
                                                        string[] tempReportAnswers = reportAnswerArray[counter].Split(";");
                                                        sampleProc_report = sampleProc_report.Replace("(B1)", tempReportAnswers[0])
                                                                                            .Replace("(B2)", tempReportAnswers[1])
                                                                                            .Replace("(B3)", tempReportAnswers[2])
                                                                                            .Replace("(B4)", tempReportAnswers[3])
                                                                                            .Replace("(B5)", tempReportAnswers[4])
                                                                                            .Replace("(B6)", tempReportAnswers[5])
                                                                                            .Replace("(B7)", tempReportAnswers[6])
                                                                                            .Replace("(B8)", tempReportAnswers[7])
                                                                                            .Replace("(B9)", tempReportAnswers[8])
                                                                                            .Replace("(B10)", tempReportAnswers[9]);
                                                    }
                                                }
                                            }
                                            counter++;
                                        }
                                    }
                                }
                                counter = 0;
                                if (completenessMethodArray != null)
                                {

                                    foreach (var method in completenessMethodArray)
                                    {
                                        if (method != null && method != string.Empty)
                                        {
                                            int result = Int32.Parse(method);
                                            var tempMethod = _soxContext.CAMethodLibrary.Where(x => x.Id.Equals(result)).FirstOrDefault();
                                            if (tempMethod != null)
                                            {
                                                //sampleProc += tempMethod.MethodName + "\n";
                                                var completeness = _soxContext.CompletenessLibrary.Where(x =>
                                                                        x.KeyReportName.Equals(filter.Filter.KeyReportName)
                                                                        && x.Method.Equals(tempMethod.MethodName)).FirstOrDefault();
                                                if (completeness != null)
                                                {
                                                    if (counter == 0 && sampleProc_completeness == "")
                                                    {
                                                        sampleProc_completeness += "C. ";
                                                    }
                                                    sampleProc_completeness += completeness.Completeness + "<br/>";
                                                    sampleProc_completeness = sampleProc_completeness.Replace("<p>", string.Empty)
                                                                                                        .Replace("</p>", string.Empty)
                                                                                                        .Replace("</strong", string.Empty)
                                                                                                        .Replace("<strong class=\"text-bold\">", string.Empty);

                                                    if (completenessAnswerArray[counter] != null && completenessAnswerArray[counter] != "")
                                                    {
                                                        string[] tempCompletenessAnswers = completenessAnswerArray[counter].Split(";");
                                                        sampleProc_completeness = sampleProc_completeness.Replace("(C1)", tempCompletenessAnswers[0])
                                                                                                            .Replace("(C2)", tempCompletenessAnswers[1])
                                                                                                            .Replace("(C3)", tempCompletenessAnswers[2])
                                                                                                            .Replace("(C4)", tempCompletenessAnswers[3])
                                                                                                            .Replace("(C5)", tempCompletenessAnswers[4])
                                                                                                            .Replace("(C6)", tempCompletenessAnswers[5])
                                                                                                            .Replace("(C7)", tempCompletenessAnswers[6])
                                                                                                            .Replace("(C8)", tempCompletenessAnswers[7])
                                                                                                            .Replace("(C9)", tempCompletenessAnswers[8])
                                                                                                            .Replace("(C10)", tempCompletenessAnswers[9]);
                                                    }
                                                }
                                            }
                                            counter++;

                                        }
                                    }
                                }
                                counter = 0;
                                if (accuracyMethodArray != null)
                                {
                                    foreach (var method in accuracyMethodArray)
                                    {
                                        if (method != null && method != string.Empty)
                                        {
                                            int result = Int32.Parse(method);
                                            var tempMethod = _soxContext.CAMethodLibrary.Where(x => x.Id.Equals(result)).FirstOrDefault();
                                            if (tempMethod != null)
                                            {
                                                var accuracy_lib = _soxContext.AccuracyLibrary.Where(x =>
                                                                        x.KeyReportName.Equals(filter.Filter.KeyReportName)
                                                                        && x.Method.Equals(tempMethod.MethodName)).FirstOrDefault();
                                                if (accuracy_lib != null)
                                                {
                                                    if (counter == 0 && sampleProc_accuracy == "")
                                                    {
                                                        sampleProc_accuracy += "D. ";
                                                    }
                                                    sampleProc_accuracy += accuracy_lib.Accuracy + "<br/>";
                                                    sampleProc_accuracy = sampleProc_accuracy.Replace("<p>", string.Empty)
                                                                                                .Replace("</p>", string.Empty)
                                                                                                .Replace("</strong", string.Empty)
                                                                                                .Replace("<strong class=\"text-bold\">", string.Empty);
                                                    if (accuracyAnswerArray[counter] != null && accuracyAnswerArray[counter] != "")
                                                    {
                                                        string[] tempAccuracyAnswers = accuracyAnswerArray[counter].Split(";");
                                                        sampleProc_accuracy = sampleProc_accuracy.Replace("(D1)", tempAccuracyAnswers[0])
                                                                                                    .Replace("(D2)", tempAccuracyAnswers[1])
                                                                                                    .Replace("(D3)", tempAccuracyAnswers[2])
                                                                                                    .Replace("(D4)", tempAccuracyAnswers[3])
                                                                                                    .Replace("(D5)", tempAccuracyAnswers[4])
                                                                                                    .Replace("(D6)", tempAccuracyAnswers[5])
                                                                                                    .Replace("(D7)", tempAccuracyAnswers[6])
                                                                                                    .Replace("(D8)", tempAccuracyAnswers[7])
                                                                                                    .Replace("(D9)", tempAccuracyAnswers[8])
                                                                                                    .Replace("(D10)", tempAccuracyAnswers[9]);
                                                    }
                                                }
                                                counter++;
                                            }
                                        }
                                    }
                                }
                                answers[14] = sampleProc_parameter + "<br/>" + sampleProc_report + "<br/>" + sampleProc_completeness + "<br/>" + sampleProc_accuracy;
                                if (string.IsNullOrEmpty(sampleProc_parameter) && string.IsNullOrEmpty(sampleProc_report) && string.IsNullOrEmpty(sampleProc_completeness) && string.IsNullOrEmpty(sampleProc_accuracy))
                                {
                                    answers[14] = string.Empty;
                                }
                            }
                            //answers[14] = sampleProc_report + "<br/>" + sampleProc_completeness;
                            answers[15] = string.Empty;
                            answers[16] = keyReportItem.Where(x => x.StrQuestion.ToLower().Contains("meeting date")).Select(x => x.StrAnswer).FirstOrDefault(); //"When did the tester observe the report being run?";
                            answers[17] = string.Empty; //controlFrequency; //Frequency of the report/query generation
                            answers[18] = string.Empty; //"Are there other UAT, Change Management testing, ITGC that support the reliability of this Key Report?";
                            answers[19] = string.Empty;
                            answers[20] = string.Empty; //"Testing Information ";
                            answers[21] = string.Empty;
                            answers[22] = string.Empty;
                            answers[23] = string.Empty;
                            answers[24] = testingStatus.Where(x => x.StrQuestion.ToLower().Contains("10. Reviewer")).Select(x => x.StrAnswer).FirstOrDefault(); //"Who performed the review?";
                            answers[25] = string.Empty;
                            answers[26] = string.Empty;
                            answers[27] = string.Empty; //"Conclusion";
                            answers[28] = string.Empty;
                            answers[29] = string.Empty;
                            answers[30] = keyReportItem.Where(x => x.StrQuestion.ToLower().Equals("28. notes")).Select(x => x.StrAnswer).FirstOrDefault();//"Notes";
                        }
                        //List<string> list_answers = answers.OfType<string>().ToList(); // this isn't going to be fast.
                        return Ok(answers);
                    }
                }

                return NoContent();
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GenerateKeyReportViewAccess {ex}", "ErrorGenerateKeyReportViewAccess");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GenerateKeyReportViewAccess");
                if (_environment.IsDevelopment())
                {
                    return BadRequest(ex.ToString());
                }
                else
                {
                    return BadRequest();
                }
            }
        }

        [HttpPost("tracker/generate-records-orig-format")]
        public IActionResult GenerateOriginalFormat([FromBody] KeyReportFilter filter)
        {
            try
            {
                List<int> itemId = new List<int>();
                List<ViewAccessOrigFormat> listOrigFormat = new List<ViewAccessOrigFormat>();
                string appIdConsolOrig = _config.GetSection("KeyReportApp").GetSection("ConsolOrigFormatId").Value;
                
                //get client name
                var checkClient = _soxContext.KeyReportUserInput
                    .Where(x => x.TagClientName.ToLower().Equals(filter.ClientName.ToLower()))
                    .AsNoTracking()
                    .Select(x => x.ItemId)
                    .Distinct()
                    .ToList();
                if (checkClient != null)
                {
                    foreach (var Id in checkClient)
                    {
                        var checkFY = _soxContext.KeyReportUserInput
                            .Where(x =>
                                x.TagFY.ToLower().Equals(filter.FY.ToLower())
                                && x.ItemId.Equals(Id)
                                && x.TagStatus.ToLower() != "inactive"
                                && x.AppId.Equals(appIdConsolOrig)
                            )
                            .AsNoTracking()
                            .Select(x => x.ItemId)
                            .Distinct()
                            .ToList();
                        if (checkFY != null)
                        {
                            foreach (var item in checkFY)
                            {
                                if (!itemId.Contains(item))
                                    itemId.Add(item);
                            }
                        }
                    }

                }


                foreach (var (Id, counter) in itemId.Select((value, i) => (value, i)))
                {
                    var checkConsolFormat = _soxContext.KeyReportUserInput.Where(x =>
                                                x.ItemId.Equals(Id)
                                                && x.AppId.Equals(appIdConsolOrig)
                                                && (
                                                    x.StrQuestion.ToLower().Contains("key control id")
                                                    || x.StrQuestion.ToLower().Contains("control activity")
                                                    || x.StrQuestion.ToLower().Contains("key/non-key control")
                                                    || x.StrQuestion.ToLower().Contains("name of key report/iuc")
                                                    || x.StrQuestion.ToLower().Contains("source process")
                                                    || x.StrQuestion.ToLower().Contains("key/non-key report")
                                                    || x.StrQuestion.ToLower().Contains("iuc type")
                                                    || x.StrQuestion.ToLower().Contains("system / source")
                                                    || x.StrQuestion.ToLower().Contains("report customized")
                                                    || x.StrQuestion.ToLower().Contains("controls relying on key report")
                                                    || x.StrQuestion.ToLower().Contains("preparer")
                                                    || x.StrQuestion.ToLower().Contains("reviewer")
                                                    || x.StrQuestion.ToLower().Contains("report notes")
                                                    )
                                                )
                                            .AsNoTracking()
                                            .ToList();

                    if (checkConsolFormat != null && checkConsolFormat.Count > 0)
                    {
                        listOrigFormat.Add(new ViewAccessOrigFormat
                        {
                            no = counter + 1,
                            keyControl = checkConsolFormat.Where(x => x.StrQuestion.ToLower().Contains("key control id")).Select(x => x.StrAnswer).FirstOrDefault(),
                            controlActivity = checkConsolFormat.Where(x => x.StrQuestion.ToLower().Contains("control activity")).Select(x => x.StrAnswer).FirstOrDefault(),
                            keyNonKeyControl = checkConsolFormat.Where(x => x.StrQuestion.ToLower().Contains("key/non-key control")).Select(x => x.StrAnswer).FirstOrDefault(),
                            nameIUC = checkConsolFormat.Where(x => x.StrQuestion.ToLower().Contains("name of key report/iuc")).Select(x => x.StrAnswer).FirstOrDefault(),
                            sourceProcess = checkConsolFormat.Where(x => x.StrQuestion.ToLower().Contains("source process")).Select(x => x.StrAnswer).FirstOrDefault(),
                            keyReport = checkConsolFormat.Where(x => x.StrQuestion.ToLower().Contains("key/non-key report")).Select(x => x.StrAnswer).FirstOrDefault(),
                            iucType = checkConsolFormat.Where(x => x.StrQuestion.ToLower().Contains("iuc type")).Select(x => x.StrAnswer).FirstOrDefault(),
                            systemSource = checkConsolFormat.Where(x => x.StrQuestion.ToLower().Contains("system / source")).Select(x => x.StrAnswer).FirstOrDefault(),
                            reportCustomized = checkConsolFormat.Where(x => x.StrQuestion.ToLower().Contains("report customized")).Select(x => x.StrAnswer).FirstOrDefault(),
                            controlRelyingIUC = checkConsolFormat.Where(x => x.StrQuestion.ToLower().Contains("controls relying on key report")).Select(x => x.StrAnswer).FirstOrDefault(),
                            preparer = checkConsolFormat.Where(x => x.StrQuestion.ToLower().Contains("preparer")).Select(x => x.StrAnswer).FirstOrDefault(),
                            reviewer = checkConsolFormat.Where(x => x.StrQuestion.ToLower().Contains("reviewer")).Select(x => x.StrAnswer).FirstOrDefault(),
                            notes = checkConsolFormat.Where(x => x.StrQuestion.ToLower().Contains("report notes")).Select(x => x.StrAnswer).FirstOrDefault()
                        });
                    }

                }
                return Ok(listOrigFormat);
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GenerateOriginalFormat {ex}", "ErrorGenerateOriginalFormat");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GenerateOriginalFormat");
                if (_environment.IsDevelopment())
                {
                    return BadRequest(ex.ToString());
                }
                else
                {
                    return BadRequest();
                }
            }


            //return Ok(excelFilename);
        }
        
        [HttpPost("tracker/generate-records-all-iuc")]
        public IActionResult GenerateAllIuc([FromBody] KeyReportFilter filter)
        {
            string excelFilename = string.Empty;

            try
            {
                List<int> itemId = new List<int>();
                List<ViewAccessAllIuc> listAllIuc = new List<ViewAccessAllIuc>();
                string appIdAllIUC = _config.GetSection("KeyReportApp").GetSection("AllIUCId").Value;

                //get client name
                var checkClient = _soxContext.KeyReportUserInput
                    .Where(x => x.TagClientName.ToLower().Equals(filter.ClientName.ToLower()))
                    .AsNoTracking()
                    .Select(x => x.ItemId)
                    .Distinct()
                    .ToList();
                if (checkClient != null)
                {
                    foreach (var Id in checkClient)
                    {
                        var checkFY = _soxContext.KeyReportUserInput
                            .Where(x =>
                                x.TagFY.ToLower().Equals(filter.FY.ToLower())
                                && x.ItemId.Equals(Id)
                                && x.TagStatus.ToLower() != "inactive"
                                && x.AppId.Equals(appIdAllIUC)
                            )
                            .AsNoTracking()
                            .Select(x => x.ItemId)
                            .Distinct()
                            .ToList();
                        if (checkFY != null)
                        {
                            foreach (var item in checkFY)
                            {
                                if (!itemId.Contains(item))
                                    itemId.Add(item);
                            }

                        }
                    }

                }

                foreach (var(Id, counter) in itemId.Select((value, i) => (value, i)))
                {
                    var checkAllIUC = _soxContext.KeyReportUserInput.Where(x =>
                            x.ItemId.Equals(Id)
                            && x.AppId.Equals(appIdAllIUC))
                        .AsNoTracking()
                        .ToList();

                    if (checkAllIUC != null && checkAllIUC.Count > 0)
                    {
                        listAllIuc.Add(new ViewAccessAllIuc
                        {
                            no = counter + 1,
                            keyControl = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("key control id")).Select(x => x.StrAnswer).FirstOrDefault(),
                            controlActivity = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("control activity")).Select(x => x.StrAnswer).FirstOrDefault(),
                            keyNonKeyControl = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("key/non-key control")).Select(x => x.StrAnswer).FirstOrDefault(),
                            nameIUC = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("name of key report/iuc")).Select(x => x.StrAnswer).FirstOrDefault(),
                            sourceProcess = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("source process")).Select(x => x.StrAnswer).FirstOrDefault(),
                            keyReport = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("key/non-key report")).Select(x => x.StrAnswer).FirstOrDefault(),
                            uniqueKeyReport = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("unique key report")).Select(x => x.StrAnswer).FirstOrDefault(),
                            iucType = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("iuc type")).Select(x => x.StrAnswer).FirstOrDefault(),
                            systemSource = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("system / source")).Select(x => x.StrAnswer).FirstOrDefault(),
                            reportCustomized = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("report customized")).Select(x => x.StrAnswer).FirstOrDefault(),
                            controlRelyingIUC = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("controls relying on key report")).Select(x => x.StrAnswer).FirstOrDefault(),
                            preparer = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("preparer")).Select(x => x.StrAnswer).FirstOrDefault(),
                            reviewer = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("16. reviewer")).Select(x => x.StrAnswer).FirstOrDefault(),
                            addedToKeyReportTracker = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("added to key report tracker")).Select(x => x.StrAnswer).FirstOrDefault(),
                            reportNotes = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("report notes")).Select(x => x.StrAnswer).FirstOrDefault(),
                            questions = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("36. questions")).Select(x => x.StrAnswer).FirstOrDefault(),
                            fastlyNotesAndQuestions = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("notes and questions")).Select(x => x.StrAnswer).FirstOrDefault(),
                            meetingDate = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("meeting date")).Select(x => x.StrAnswer).FirstOrDefault(),
                            process = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("40. process")).Select(x => x.StrAnswer).FirstOrDefault(),
                            descriptionKeyReport = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("description of key report (key data fields used, purpose of report)")).Select(x => x.StrAnswer).FirstOrDefault(),
                            keyReportType = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("key report type")).Select(x => x.StrAnswer).FirstOrDefault(),
                            howReportGenerated = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("how is the report generated?")).Select(x => x.StrAnswer).FirstOrDefault(),
                            howReportUsed = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("how is the report used to support the control(s)")).Select(x => x.StrAnswer).FirstOrDefault(),
                            stepsPerformedAccuracy = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("what steps are performed to validate the accuracy of the report")).Select(x => x.StrAnswer).FirstOrDefault(),
                            stepsPerformedCompleteness = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("what steps are performed to validate the completeness of the report")).Select(x => x.StrAnswer).FirstOrDefault(),
                            stepsPerformedValidateSource = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("what steps are performed to validate the source data used by the report")).Select(x => x.StrAnswer).FirstOrDefault(),
                            areParameters = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("are parameters (e.g. date ranges) input each time this report is run")).Select(x => x.StrAnswer).FirstOrDefault(),
                            whoIsAuthorized = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("who is authorized to make/request changes to this report")).Select(x => x.StrAnswer).FirstOrDefault(),
                            effectiveDate = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("effective date")).Select(x => x.StrAnswer).FirstOrDefault(),
                            whoHasAccessToEdit = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("who has access to edit/modify this report")).Select(x => x.StrAnswer).FirstOrDefault(),
                            whoHasAccessToRun = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("who has access to run the report? restricted report access")).Select(x => x.StrAnswer).FirstOrDefault(),
                            reportLastModified = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("when was this report last modified")).Select(x => x.StrAnswer).FirstOrDefault(),
                            howWasItTestedWhenLastModified = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("how was it tested when last modified")).Select(x => x.StrAnswer).FirstOrDefault(),
                            itReportOwner = checkAllIUC.Where(x => x.StrQuestion.ToLower().Contains("it report owner")).Select(x => x.StrAnswer).FirstOrDefault()
                        });
                }

                }
                return Ok(listAllIuc);
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GenerateAllIuc {ex}", "ErrorGenerateAllIuc");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GenerateAllIuc");
                if (_environment.IsDevelopment())
                {
                    return BadRequest(ex.ToString());
                }
                else
                {
                    return BadRequest();
                }
            }
        }
        
        [HttpPost("tracker/generate-records-test-status-tracker")]    
        public IActionResult GenerateTestStatus([FromBody] KeyReportFilter filter)
        {
            try
            {
                List<int> itemId = new List<int>();
                List<ViewAccessTestStatus> listTestStatus = new List<ViewAccessTestStatus>();
                string appIdTestStatus = _config.GetSection("KeyReportApp").GetSection("TestStatusTrackerId").Value;

                //get client name
                var checkClient = _soxContext.KeyReportUserInput
                    .Where(x => x.TagClientName.ToLower().Equals(filter.ClientName.ToLower()))
                    .AsNoTracking()
                    .Select(x => x.ItemId)
                    .Distinct()
                    .ToList();
                if (checkClient != null)
                {
                    foreach (var Id in checkClient)
                    {
                        var checkFY = _soxContext.KeyReportUserInput
                            .Where(x =>
                                x.TagFY.ToLower().Equals(filter.FY.ToLower())
                                && x.ItemId.Equals(Id)
                                && x.TagStatus.ToLower() != "inactive"
                                && x.AppId.Equals(appIdTestStatus)
                            )
                            .AsNoTracking()
                            .Select(x => x.ItemId)
                            .Distinct()
                            .ToList();
                        if (checkFY != null)
                        {
                            foreach (var item in checkFY)
                            {
                                if (!itemId.Contains(item))
                                    itemId.Add(item);
                            }

                        }
                    }

                }

                foreach (var (Id, counter) in itemId.Select((value, i) => (value, i)))
                {
                    var checkTestStatus = _soxContext.KeyReportUserInput.Where(x =>
                            x.ItemId.Equals(Id)
                            && x.AppId.Equals(appIdTestStatus))
                        .AsNoTracking()
                        .ToList();

                    if (checkTestStatus != null && checkTestStatus.Count > 0)
                    {
                        listTestStatus.Add(new ViewAccessTestStatus
                        {
                            no = counter + 1,
                            nameIuc = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("name of key report/iuc")).Select(x => x.StrAnswer).FirstOrDefault(),
                            process = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("source process")).Select(x => x.StrAnswer).FirstOrDefault(),
                            keyControl = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("key control id")).Select(x => x.StrAnswer).FirstOrDefault(),
                            controlActvity = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("control activity")).Select(x => x.StrAnswer).FirstOrDefault(),
                            keyNonKeyControl = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("key/non-key control")).Select(x => x.StrAnswer).FirstOrDefault(),
                            keyReport = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("key/non-key report")).Select(x => x.StrAnswer).FirstOrDefault(),
                            uniqueKeyReport = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("unique key report")).Select(x => x.StrAnswer).FirstOrDefault(),
                            preparer = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("preparer")).Select(x => x.StrAnswer).FirstOrDefault(),
                            processOwner = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("process owner")).Select(x => x.StrAnswer).FirstOrDefault(),
                            keyReportOwner = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("key report owner")).Select(x => x.StrAnswer).FirstOrDefault(),
                            keyReportITOwner = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("key report it owner")).Select(x => x.StrAnswer).FirstOrDefault(),
                            setupLeadsheetForTesting = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("set up report leadsheet for testing")).Select(x => x.StrAnswer).FirstOrDefault(),
                            scheduleProcessOwnerMeeting = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("schedule process owner meeting")).Select(x => x.StrAnswer).FirstOrDefault(),
                            reportReceived = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("reports received")).Select(x => x.StrAnswer).FirstOrDefault(),
                            pbcStatus = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("pbc status")).Select(x => x.StrAnswer).FirstOrDefault(),
                            tester = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("tester")).Select(x => x.StrAnswer).FirstOrDefault(),
                            firstReviewer = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("1st reviewer")).Select(x => x.StrAnswer).FirstOrDefault(),
                            secondReviewer = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("2nd reviewer")).Select(x => x.StrAnswer).FirstOrDefault(),
                            testingStatus = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("testing status")).Select(x => x.StrAnswer).FirstOrDefault(),
                            a2q2DueDateTesting = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("a2q2 due date (testing)")).Select(x => x.StrAnswer).FirstOrDefault(),
                            sentToClient = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("sending to client status")).Select(x => x.StrAnswer).FirstOrDefault(),
                            clientReviewStatus = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("client review status")).Select(x => x.StrAnswer).FirstOrDefault(),
                            sentToDeloitte = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("sending to auditor status")).Select(x => x.StrAnswer).FirstOrDefault(),
                            a2q2Notes = checkTestStatus.Where(x => x.StrQuestion.ToLower().Contains("a2q2 notes")).Select(x => x.StrAnswer).FirstOrDefault(),
                        });
                    }
                }
                return Ok(listTestStatus);
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GenerateTestStatus {ex}", "ErrorGenerateTestStatus");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GenerateTestStatus");
                if (_environment.IsDevelopment())
                {
                    return BadRequest(ex.ToString());
                }
                else
                {
                    return BadRequest();
                }
            }
        }
        
        [HttpPost("tracker/generate-records-exceptions")]
        public IActionResult GenerateListExceptions([FromBody] KeyReportFilter filter)
        {
            try
            {
                List<int> itemId = new List<int>();
                List<ViewAccessListOfExceptions> listOfExceptions = new List<ViewAccessListOfExceptions>();
                string appIdException = _config.GetSection("KeyReportApp").GetSection("ExceptionId").Value;

                //get client name
                var checkClient = _soxContext.KeyReportUserInput
                    .Where(x => x.TagClientName.ToLower().Equals(filter.ClientName.ToLower()))
                    .AsNoTracking()
                    .Select(x => x.ItemId)
                    .Distinct()
                    .ToList();
                if (checkClient != null)
                {
                    foreach (var Id in checkClient)
                    {
                        var checkFY = _soxContext.KeyReportUserInput
                            .Where(x =>
                                x.TagFY.ToLower().Equals(filter.FY.ToLower())
                                && x.ItemId.Equals(Id)
                                && x.TagStatus.ToLower() != "inactive"
                                && x.AppId.Equals(appIdException)
                            )
                            .AsNoTracking()
                            .Select(x => x.ItemId)
                            .Distinct()
                            .ToList();
                        if (checkFY != null)
                        {
                            foreach (var item in checkFY)
                            {
                                if (!itemId.Contains(item))
                                    itemId.Add(item);
                            }

                        }
                    }

                }

                foreach (var (Id, counter) in itemId.Select((value, i) => (value, i)))
                {
                    var checkListException = _soxContext.KeyReportUserInput.Where(x =>
                            x.ItemId.Equals(Id)
                            && x.AppId.Equals(appIdException))
                        .AsNoTracking()
                        .ToList();

                    if (checkListException != null && checkListException.Count > 0)
                    {
                        listOfExceptions.Add(new ViewAccessListOfExceptions
                        {
                            no = counter + 1,
                            nameIuc = checkListException.Where(x => x.StrQuestion.ToLower().Contains("name of key report/iuc")).Select(x => x.StrAnswer).FirstOrDefault(),
                            keyControl = checkListException.Where(x => x.StrQuestion.ToLower().Contains("key control id")).Select(x => x.StrAnswer).FirstOrDefault(),
                            controlActvity = checkListException.Where(x => x.StrQuestion.ToLower().Contains("control activity")).Select(x => x.StrAnswer).FirstOrDefault(),
                            exceptionNoted = checkListException.Where(x => x.StrQuestion.ToLower().Contains("exceptions noted")).Select(x => x.StrAnswer).FirstOrDefault(),
                            reasonForException = checkListException.Where(x => x.StrQuestion.ToLower().Contains("reasons for exceptions")).Select(x => x.StrAnswer).FirstOrDefault(),
                            remediation = checkListException.Where(x => x.StrQuestion.ToLower().Contains("remediation")).Select(x => x.StrAnswer).FirstOrDefault()
                        });
                    }

                }
                return Ok(listOfExceptions);
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GenerateListExceptions {ex}", "ErrorGenerateListExceptions");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GenerateListExceptions");
                if (_environment.IsDevelopment())
                {
                    return BadRequest(ex.ToString());
                }
                else
                {
                    return BadRequest();
                }
            }
        }
        //public ExcelWorksheet CopyWorksheet(ExcelWorkbook workbook, string existingWorksheetName, string newWorksheetName)
        //{
        //    ExcelWorksheet worksheet = workbook.Worksheets.Copy(existingWorksheetName, newWorksheetName);
        //    return worksheet;
        //}

    }
}
