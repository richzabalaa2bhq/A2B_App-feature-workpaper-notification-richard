using A2B_App.Server.Data;
using A2B_App.Server.Services;
//using A2B_App.Server.Log;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
//using FileLog = A2B_App.Server.Log.FileLog;
using A2B_App.Shared.Sox;
using System.Linq;
using Microsoft.EntityFrameworkCore;
using System.Threading.Tasks;
using A2B_App.Shared.Podio;
using System.Net;
using PodioAPI;
using PodioAPI.Models;
using PodioAPI.Utils.ItemFields;
using Newtonsoft.Json;
using System.Net.Http;
using System.Text;
using Newtonsoft.Json.Linq;

namespace A2B_App.Server.Controllers
{
    [Authorize]
    [Route("api/[controller]")]
    [ApiController]
    public class SoxTrackerController : ControllerBase
    {

        private readonly IConfiguration _config;
        private readonly ILogger<SoxTrackerController> _logger;
        private readonly IWebHostEnvironment _environment;
        private readonly SoxContext _soxContext;
        //private readonly TimeContext _timeContext;

        public SoxTrackerController(
            IConfiguration config,
            ILogger<SoxTrackerController> logger,
            IWebHostEnvironment environment,
            SoxContext soxContext
        )
        {
            _config = config;
            _logger = logger;
            _environment = environment;
            _soxContext = soxContext;
        }

        [AllowAnonymous]
        [HttpGet("fy")]
        public IActionResult GetListSoxTrackerFYAsync()
        {
            List<string> _listFY = new List<string>();
            try
            {
                using (var context = _soxContext.Database.BeginTransaction())
                {
                    //Get all active FY
                    var listFy = _soxContext.SoxTracker
                        .Select(x => x.FY);

                    //check if value is not null
                    if (listFy != null)
                    {
                        _listFY = listFy.Distinct().OrderBy(x => x).ToList();
                    }
                }
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetListSoxTrackerFYAsync {ex}", "ErrorGetListSoxTrackerFYAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetListSoxTrackerFYAsync");

            }

            if (_listFY != null)
            {
                _listFY.Sort(); //sort list 
                return Ok(_listFY.ToArray());
            }
            else
            {
                return NoContent();
            }

        }

        [AllowAnonymous]
        [HttpPost("tracker")]
        public IActionResult GetTrackerAsync([FromBody] KeyReportFilter filter)
        {
            SoxTracker soxTracker = new SoxTracker();
            try
            {

                //Get all active FY
                var checkSoxTracker = _soxContext.SoxTracker
                    .Where(x =>
                        x.FY.ToLower().Equals(filter.FY.ToLower())
                        && x.ClientName.ToLower().Equals(filter.ClientName.ToLower())
                        && x.ControlId.ToLower().Equals(filter.ControlId.ToLower())
                        && x.KeyReport.ToLower().Equals("yes")
                        && x.KeyReportName.ToLower().Equals(filter.KeyReportName.ToLower())
                        && x.Status.ToLower().Equals("active")
                    )
                    .AsNoTracking()
                    .FirstOrDefault();

                //check if value is not null
                if (checkSoxTracker != null)
                {
                    return Ok(checkSoxTracker);
                }

                return NoContent();


            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "GetTrackerAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetTrackerAsync");
                return BadRequest();
            }

        }


        [AllowAnonymous]
        [HttpPost("clientbyyear")]
        public IActionResult GetListClientAsync([FromBody] RcmQuestionnaireFilter filter)
        {
            List<RcmOutputFile> _listClient = new List<RcmOutputFile>();
            try
            {
                string rcmSharefile = _config.GetSection("SharefileApi").GetSection("SoxTrackerFolder").GetSection("Link").Value;
                using (var context = _soxContext.Database.BeginTransaction())
                {
                    var listRcm = _soxContext.SoxTracker
                        .Where(x =>
                            x.FY.Equals(filter.FY) &&
                            x.ClientName != string.Empty &&
                            x.ClientName != null
                        )
                        .Select(x => x.ClientName);

                    if (listRcm != null)
                    {
                        var listClient = listRcm.Distinct().OrderBy(x => x).ToList();
                        foreach (var item in listClient)
                        {
                            _listClient.Add(new RcmOutputFile
                            {
                                ClientName = item,
                                LoadingStatus = string.Empty,
                                SharefileLink = rcmSharefile,
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
        [HttpPost("control")]
        public IActionResult GetSoxTrackerControl([FromBody] RcmQuestionnaireFilter filter)
        {
            SoxTracker _soxTracker = new SoxTracker();
            try
            {

                //var checkClientName = _soxContext.Rcm.Where(x => x.ClientName.Equals(filter.Client)).Select(x => x.ClientNameText).FirstOrDefault();
                //if(checkClientName != null)
                //{
                //}
                var checkSoxTracker = _soxContext.SoxTracker
                    .Where(x =>
                        x.FY.Equals(filter.FY) &&
                        x.ClientName.Equals(filter.Client) &&
                        x.Process.Equals(filter.Process) &&
                        x.Subprocess.Equals(filter.SubProcess) &&
                        x.ControlId.Equals(filter.ControlId)
                    )
                    .FirstOrDefault();

                if (checkSoxTracker != null)
                {
                    _soxTracker = checkSoxTracker;
                }



            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetListClientAsync {ex}", "ErrorGetListClientAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetListClientAsync");
            }

            if (_soxTracker.Id != 0)
            {
                return Ok(_soxTracker);
            }
            else
            {
                return NoContent();
            }

        }


        [AllowAnonymous]
        [HttpGet("questionnaire")]
        public IActionResult GetSoxTrackerQuestionnaire()
        {
            SoxTrackerQuestionnaire _soxTrackerQuestionnaire = new SoxTrackerQuestionnaire();
            try
            {

                //Get SoxTrackerQuestionnaire
                var checkSoxTrackerQuestionnaire = _soxContext.SoxTrackerQuestionnaire
                    //.Include(x => x.ListSoxTrackerAppCategory)
                    //.Include(x => x.ListSoxTrackerAppRelationship)
                    .FirstOrDefault();

                if (checkSoxTrackerQuestionnaire != null)
                {
                    //Get SoxTrackerQuestionnaire Category
                    var checkCategoryOptions = _soxContext.SoxTrackerAppCategory.Where(x => x.SoxTrackerQuestionnaire.Equals(checkSoxTrackerQuestionnaire));

                    //Get SoxTrackerQuestionnaire App Relationship
                    var checkAppRelationship = _soxContext.SoxTrackerAppRelationship.Where(x => x.SoxTrackerQuestionnaire.Equals(checkSoxTrackerQuestionnaire));

                    if (checkCategoryOptions != null)
                    {
                        _soxTrackerQuestionnaire.ListSoxTrackerAppCategory = checkCategoryOptions.ToList();
                    }
                    if (checkAppRelationship != null)
                    {
                        _soxTrackerQuestionnaire.ListSoxTrackerAppRelationship = checkAppRelationship.ToList();
                    }

                    _soxTrackerQuestionnaire = checkSoxTrackerQuestionnaire;
                }



            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetSoxTrackerQuestionnaire {ex}", "ErrorGetSoxTrackerQuestionnaire");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetSoxTrackerQuestionnaire");
            }

            if (_soxTrackerQuestionnaire.Id != 0)
            {
                return Ok(_soxTrackerQuestionnaire);
            }
            else
            {
                return NoContent();
            }

        }


        [AllowAnonymous]
        [HttpPost("generate/{clientName}/{Fy}")]
        public IActionResult GenerateTrackerFile(string clientName, string Fy)
        {
            //List<string> excelFilename = new List<string>();
            string excelFilename = string.Empty;

            try
            {

                //Get Sox Tracker
                List<SoxTracker> _tracker = new List<SoxTracker>();
                var checkTracker = _soxContext.SoxTracker
                    .Where(x => x.ClientName.Equals(clientName) && x.FY.Equals(Fy))
                    .OrderBy(x => x.Process)
                    .ThenBy(x => x.ControlId);
                if (checkTracker != null && checkTracker.Count() > 0)
                {
                    _tracker = checkTracker.ToList(); ;
                    if (_tracker != null && _tracker.Count > 0)
                    {
                        //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        using (ExcelPackage xls = new ExcelPackage())
                        {
                            ExcelService xlsService = new ExcelService();
                            FormatService txtFormat = new FormatService();
                            //set sheet name
                            var ws = xls.Workbook.Worksheets.Add("Tracker");

                            ws.View.ZoomScale = 90;
                            //set column width
                            ws.Column(1).Width = 15;
                            ws.Column(2).Width = 15;
                            ws.Column(3).Width = 25;
                            ws.Column(4).Width = 15;
                            ws.Column(5).Width = 15;
                            ws.Column(6).Width = 35;
                            ws.Column(7).Width = 10;
                            ws.Column(8).Width = 15;
                            ws.Column(9).Width = 10;
                            ws.Column(10).Width = 14;
                            ws.Column(11).Width = 10;
                            ws.Column(12).Width = 15;
                            ws.Column(13).Width = 22;
                            ws.Column(14).Width = 22;
                            ws.Column(15).Width = 22;
                            ws.Column(16).Width = 15;
                            ws.Column(17).Width = 15;
                            ws.Column(18).Width = 15;
                            ws.Column(19).Width = 15;
                            ws.Column(20).Width = 15;
                            ws.Column(21).Width = 15;
                            ws.Column(22).Width = 15;
                            ws.Column(23).Width = 15;
                            ws.Column(24).Width = 15;
                            ws.Column(25).Width = 15;
                            ws.Column(26).Width = 15;
                            ws.Column(27).Width = 15;
                            ws.Column(28).Width = 15;
                            ws.Column(29).Width = 15;
                            ws.Column(30).Width = 15;
                            ws.Column(31).Width = 15;
                            ws.Column(32).Width = 15;
                            ws.Column(33).Width = 15;
                            //disable grid
                            ws.View.ShowGridLines = false;

                            //set row
                            int row = 1;

                            //set title
                            ws.Cells[row, 1].Value = clientName;
                            xlsService.ExcelSetArialSize10(ws, row, 1, row, 1);
                            xlsService.ExcelSetFontSize(ws, row, 1, row, 1, 12);
                            row++;
                            ws.Cells[row, 1].Value = "Sox Tracker";
                            xlsService.ExcelSetArialSize10(ws, row, 1, row, 1);
                            xlsService.ExcelSetFontSize(ws, row, 1, row, 1, 12);
                            row++;
                            ws.Cells[row, 1].Value = Fy;
                            xlsService.ExcelSetArialSize10(ws, row, 1, row, 1);
                            xlsService.ExcelSetFontBold(ws, 1, 1, row, 1); //(workspace, from row, from column, to row, to column)
                            xlsService.ExcelSetFontSize(ws, row, 1, row, 1, 12);
                            row += 3;

                            //set table header
                            ws.Row(row).Height = 21;
                            ws.Cells[row, 1].Value = "Consolidated RCM";
                            ws.Cells[row, 7].Value = "Master PBC ";
                            ws.Cells["A" + row + ":F" + row].Merge = true;
                            ws.Cells["G" + row + ":O" + row].Merge = true;
                            ws.Cells["P" + row + ":AE" + row].Merge = true;
                            xlsService.ExcelSetHorizontalAlignCenter(ws, row, 1, row, 31);
                            xlsService.ExcelSetVerticalAlignCenter(ws, row, 1, row, 31);
                            xlsService.ExcelSetBorder(ws, row, 1, row, 31);
                            xlsService.ExcelSetBackgroundColorGray(ws, row, 1, row, 6);
                            xlsService.ExcelSetBackgroundColorRed(ws, row, 7, row, 15);
                            xlsService.ExcelSetBackgroundColorGreen(ws, row, 16, row, 31);
                            xlsService.ExcelSetBackgroundColorLightBlue(ws, row, 32, row, 32);
                            xlsService.ExcelSetBackgroundColorOrange(ws, row, 33, row, 33);
                            xlsService.ExcelSetFontColorWhite(ws, row, 7, row, 15);
                            xlsService.ExcelSetArialSize10(ws, row, 1, row, 33);
                            xlsService.ExcelSetFontBold(ws, row, 1, row, 33);
                            row++;

                            ws.Cells[row, 1].Value = "Process";
                            ws.Cells[row, 2].Value = "Sub-process";
                            ws.Cells[row, 3].Value = "Control ID";
                            ws.Cells[row, 4].Value = "Control Activity";
                            ws.Cells[row, 5].Value = "Control Owner";
                            ws.Cells[row, 6].Value = "Test Procedures";
                            ws.Cells[row, 7].Value = "PBC";
                            ws.Cells[row, 8].Value = "PBC Owner";
                            ws.Cells[row, 9].Value = "Population File Request";
                            ws.Cells[row, 10].Value = "Sample selection/sub-selection";
                            ws.Cells[row, 11].Value = "R3 sample testing required";
                            ws.Cells[row, 12].Value = "WT PBC Status";
                            ws.Cells[row, 13].Value = "Round 1 PBC Status";
                            ws.Cells[row, 14].Value = "Round 2 PBC Status";
                            ws.Cells[row, 15].Value = "Round 3 PBC Status";
                            ws.Cells[row, 16].Value = "WT Tester";
                            ws.Cells[row, 17].Value = "WT 1st Level Reviewer";
                            ws.Cells[row, 18].Value = "WT 2nd Level Reviewer";
                            ws.Cells[row, 19].Value = "WT Status";
                            ws.Cells[row, 20].Value = "Round 1 Tester";
                            ws.Cells[row, 21].Value = "R1 1st Level Reviewer";
                            ws.Cells[row, 22].Value = "R1 2nd Level Reviewer";
                            ws.Cells[row, 23].Value = "Round 1 Status";
                            ws.Cells[row, 24].Value = "Round 2 Tester";
                            ws.Cells[row, 25].Value = "R2 1st Level Reviewer";
                            ws.Cells[row, 26].Value = "R2 2nd Level Reviewer";
                            ws.Cells[row, 27].Value = "Round 2 Status";
                            ws.Cells[row, 28].Value = "Round 3 Tester";
                            ws.Cells[row, 29].Value = "R3 1st Level Reviewer";
                            ws.Cells[row, 30].Value = "R3 2nd Level Reviewer";
                            ws.Cells[row, 31].Value = "Round 3 Status";
                            ws.Cells[row, 32].Value = "Notes for client status";
                            ws.Cells[row, 33].Value = "Status Notes";
                            xlsService.ExcelSetHorizontalAlignCenter(ws, row, 1, row, 33);
                            xlsService.ExcelSetVerticalAlignCenter(ws, row, 1, row, 33);
                            xlsService.ExcelSetBorder(ws, row, 1, row, 31);
                            xlsService.ExcelSetBackgroundColorGray(ws, row, 1, row, 6);
                            xlsService.ExcelSetBackgroundColorRed(ws, row, 7, row, 15);
                            xlsService.ExcelSetBackgroundColorGreen(ws, row, 16, row, 31);
                            xlsService.ExcelSetBackgroundColorLightBlue(ws, row, 32, row, 32);
                            xlsService.ExcelSetBackgroundColorOrange(ws, row, 33, row, 33);
                            xlsService.ExcelSetBackgroundColorGreen(ws, row, 12, row, 13);
                            xlsService.ExcelSetBackgroundColorOrange(ws, row, 14, row, 14);
                            xlsService.ExcelSetBackgroundColorOrange(ws, row, 24, row, 27);
                            xlsService.ExcelSetBackgroundColorRed(ws, row, 28, row, 31);
                            xlsService.ExcelSetFontColorWhite(ws, row, 28, row, 31);
                            xlsService.ExcelSetFontColorWhite(ws, row, 7, row, 11);
                            xlsService.ExcelSetFontColorWhite(ws, row, 15, row, 15);
                            xlsService.ExcelSetArialSize10(ws, row, 1, row, 33);
                            xlsService.ExcelSetFontBold(ws, row, 1, row, 33);
                            xlsService.ExcelSetBorder(ws, row - 1, 32, row, 32);
                            xlsService.ExcelSetBorder(ws, row - 1, 33, row, 33);
                            xlsService.ExcelWrapText(ws, row, 1, row, 33);

                            //set auto filter
                            ws.Cells[row, 1, row, 33].AutoFilter = true;
                            row++;

                            ////set vertical header
                            //ws.Cells[row, 6].Style.TextRotation = 90;
                            //ws.Cells[row, 7].Style.TextRotation = 90;
                            //ws.Cells[row, 8].Style.TextRotation = 90;
                            //ws.Cells[row, 9].Style.TextRotation = 90;
                            //ws.Cells[row, 10].Style.TextRotation = 90;



                            //row++;

                            //loop through all Sox Tracker found
                            foreach (var item in _tracker)
                            {

                                var rcmData = _soxContext.Rcm
                                    .Where(x =>
                                        x.ClientName.Equals(clientName)
                                        && x.FY.Equals(Fy)
                                        && x.ControlId.Equals(item.ControlId))
                                    .FirstOrDefault();

                                ws.Cells[row, 1].Value = item.Process != null ? item.Process : string.Empty;
                                ws.Cells[row, 2].Value = item.Subprocess != null ? item.Subprocess : string.Empty;
                                ws.Cells[row, 3].Value = item.ControlId != null ? item.ControlId : string.Empty;
                                if (rcmData != null)
                                {
                                    //ws.Cells[row, 4].Value = rcmData.ControlActivityFy19 != null ? rcmData.ControlActivityFy19 : string.Empty;06/20/21
                                    ws.Cells[row, 4].Value = rcmData.ControlActivityFy19 != null ? txtFormat.FormatwithNewLine(rcmData.ControlActivityFy19,true) : string.Empty;
                                    ws.Cells[row, 5].Value = rcmData.ControlOwner != null ? rcmData.ControlOwner : string.Empty;
                                    //ws.Cells[row, 6].Value = rcmData.TestProc != null ? rcmData.TestProc : string.Empty;06/20/21
                                    ws.Cells[row, 6].Value = rcmData.TestProc! != null ? txtFormat.FormatwithNewLine(rcmData.TestProc, true) : string.Empty;
                                    ws.Cells[row, 7].Value = rcmData.PbcList != null ? txtFormat.FormatwithNewLine(rcmData.PbcList, true) : string.Empty;
                                    
                                }
                                //ws.Cells[row, 7].Value = item.PBC != null ? item.PBC : string.Empty; 06/20/21
                                ws.Cells[row, 8].Value = item.PBCOwner != null ? item.PBCOwner : string.Empty;
                                ws.Cells[row, 9].Value = item.PopulationFileRequest != null ? item.PopulationFileRequest : string.Empty;
                                ws.Cells[row, 10].Value = item.SampleSelection != null ? item.SampleSelection : string.Empty;
                                ws.Cells[row, 11].Value = item.R3Sample != null ? txtFormat.FormatwithNewLine(item.R3Sample,true) : string.Empty; //"R3 sample testing required";06/20/21
                                ws.Cells[row, 12].Value = item.WTPBC != null ? item.WTPBC : string.Empty; //"WT PBC Status";
                                ws.Cells[row, 13].Value = item.R1PBC != null ? item.R1PBC : string.Empty; //"Round 1 PBC Status";
                                ws.Cells[row, 14].Value = item.R2PBC != null ? item.R2PBC : string.Empty; //"Round 2 PBC Status";
                                ws.Cells[row, 15].Value = item.R3PBC != null ? item.R3PBC : string.Empty; //"Round 3 PBC Status";
                                ws.Cells[row, 16].Value = item.WTTester != null ? item.WTTester : string.Empty; //"WT Tester";
                                ws.Cells[row, 17].Value = item.WT1LReviewer != null ? item.WT1LReviewer : string.Empty; //"WT 1st Level Reviewer";
                                ws.Cells[row, 18].Value = item.WT2LReviewer != null ? item.WT2LReviewer : string.Empty; //"WT 2nd Level Reviewer";
                                ws.Cells[row, 19].Value = item.WTTestingStatus != null ? item.WTTestingStatus : string.Empty; //"WT Status";
                                ws.Cells[row, 20].Value = item.R1Tester != null ? item.R1Tester : string.Empty; //"Round 1 Tester";
                                ws.Cells[row, 21].Value = item.R11LReviewer != null ? item.R11LReviewer : string.Empty; //"R1 1st Level Reviewer";
                                ws.Cells[row, 22].Value = item.R12LReviewer != null ? item.R12LReviewer : string.Empty; //"R1 2nd Level Reviewer";
                                ws.Cells[row, 23].Value = item.R1TestingStatus != null ? item.R1TestingStatus : string.Empty; //"Round 1 Status";
                                ws.Cells[row, 24].Value = item.R2Tester != null ? item.R2Tester : string.Empty; //"Round 2 Tester";
                                ws.Cells[row, 25].Value = item.R21LReviewer != null ? item.R21LReviewer : string.Empty; //"R2 1st Level Reviewer";
                                ws.Cells[row, 26].Value = item.R22LReviewer != null ? item.R22LReviewer : string.Empty; //"R2 2nd Level Reviewer";
                                ws.Cells[row, 27].Value = item.R2TestingStatus != null ? item.R2TestingStatus : string.Empty; //"Round 2 Status";
                                ws.Cells[row, 28].Value = item.R3Tester != null ? item.R3Tester : string.Empty; //"Round 3 Tester";
                                ws.Cells[row, 29].Value = item.R31LReviewer != null ? item.R31LReviewer : string.Empty; //"R3 1st Level Reviewer";
                                ws.Cells[row, 30].Value = item.R32LReviewer != null ? item.R32LReviewer : string.Empty; //"R3 2nd Level Reviewer";
                                ws.Cells[row, 31].Value = item.R3TestingStatus != null ? item.R3TestingStatus : string.Empty; //"Round 3 Status";
                                                                                                                              //ws.Cells[row, 32].Value = "Notes for client status";
                                                                                                                              //ws.Cells[row, 33].Value = "Status Notes";

                                xlsService.ExcelSetVerticalAlignCenter(ws, row, 1, row, 33);
                                xlsService.ExcelSetHorizontalAlignLeft(ws, row, 1, row, 33);
                                xlsService.ExcelSetArialSize10(ws, row, 1, row, 33);
                                xlsService.ExcelSetBorder(ws, row, 1, row, 33);
                                xlsService.ExcelWrapText(ws, row, 1, row, 33);
                                row++;
                            }


                            //save file
                            string startupPath = Environment.CurrentDirectory;
                            //string strSourceDownload = startupPath + "\\include\\sampleselection\\download\\";
                            string strSourceDownload = Path.Combine(startupPath, "include", "upload", "soxtracker");

                            if (!Directory.Exists(strSourceDownload))
                            {
                                Directory.CreateDirectory(strSourceDownload);
                            }
                            var ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                            string filename = $"Tracker-{clientName}-{ts}.xlsx";
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
                    else
                    {
                        return NoContent();
                    }
                }
                else
                {
                    return NoContent();
                }



            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GenerateSoxTrackerFile {ex}", "ErrorGenerateSoxTrackerFile");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GenerateSoxTrackerFile");
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
        [HttpPost("podiocreate")]
        public async Task<IActionResult> CreatePodioSoxTrackerAsync([FromBody] SoxTracker soxTracker)
        {
            bool status = false;

            try
            {
                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("SoxTrackerApp").GetSection("AppId").Value;
                PodioAppKey.AppToken = _config.GetSection("SoxTrackerApp").GetSection("AppToken").Value;


                if (PodioAppKey.AppId != string.Empty && PodioAppKey.AppToken != string.Empty)
                {
                    var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                    ServicePointManager.Expect100Continue = true;
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                    await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                    if (podio.IsAuthenticated() && soxTracker != null)
                    {

                        Item soxTrackerItem = new Item();

                        if (soxTracker.FY != string.Empty && soxTracker.FY != null)
                        {
                            int q1Field = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q1").Value);
                            var textQ1 = soxTrackerItem.Field<TextItemField>(q1Field);
                            textQ1.Value = soxTracker.FY;
                        }

                        if (soxTracker.ClientName != string.Empty && soxTracker.ClientName != null)
                        {
                            var checkClientItemId = _soxContext.ClientSs.Where(x => x.Name.Equals(soxTracker.ClientName)).Select(x => x.ItemId).FirstOrDefault();
                            if (checkClientItemId != null)
                            {
                                int q2Field = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q2").Value);
                                var clientRef = soxTrackerItem.Field<AppItemField>(q2Field);
                                List<int> listClient = new List<int>();
                                listClient.Add(checkClientItemId.Value);
                                clientRef.ItemIds = listClient;
                            }
                        }

                        if (soxTracker.Process != string.Empty && soxTracker.Process != null)
                        {

                            var checkRcmProcess = _soxContext.RcmProcess.Where(x => x.Process.Equals(soxTracker.Process)).FirstOrDefault();
                            if (checkRcmProcess != null)
                            {
                                int q3Field = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q3").Value);
                                var processRef = soxTrackerItem.Field<AppItemField>(q3Field);
                                List<int> listProcess = new List<int>();
                                listProcess.Add(checkRcmProcess.PodioItemId);
                                processRef.ItemIds = listProcess;
                            }

                        }

                        if (soxTracker.Subprocess != string.Empty && soxTracker.Subprocess != null)
                        {
                            var checkRcmSubProcess = _soxContext.RcmSubProcess.Where(x => x.SubProcess.Equals(soxTracker.Subprocess)).FirstOrDefault();
                            if (checkRcmSubProcess != null)
                            {
                                int q4Field = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q4").Value);
                                var subProcessRef = soxTrackerItem.Field<AppItemField>(q4Field);
                                List<int> listSubProcess = new List<int>();
                                listSubProcess.Add(checkRcmSubProcess.PodioItemId);
                                subProcessRef.ItemIds = listSubProcess;
                            }
                        }

                        if (soxTracker.ControlId != string.Empty && soxTracker.ControlId != null)
                        {
                            var checkRcmControlId = _soxContext.RcmControlId.Where(x => x.ControlId.Equals(soxTracker.ControlId)).FirstOrDefault();
                            if (checkRcmControlId != null)
                            {
                                int q5Field = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q5").Value);
                                var controlIdRef = soxTrackerItem.Field<AppItemField>(q5Field);
                                List<int> listControlId = new List<int>();
                                listControlId.Add(checkRcmControlId.PodioItemId);
                                controlIdRef.ItemIds = listControlId;
                            }
                        }


                        //populate remaining fields
                        PopulatePodioFields(soxTrackerItem, soxTracker, false);

                        //set podio status to active
                        int statusField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Status").Value);
                        var categoryStatus = soxTrackerItem.Field<CategoryItemField>(statusField);
                        categoryStatus.OptionText = "Active";


                        var roundId = await podio.ItemService.AddNewItem(Int32.Parse(PodioAppKey.AppId), soxTrackerItem);

                        soxTracker.PodioItemId = int.Parse(roundId.ToString());
                        status = true;

                    }
                }

            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex.ToString());
                //ErrorLog.Write(ex);
                FileLog.Write(ex.ToString(), "ErrorCreatePodioSoxTrackerAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "CreatePodioSoxTrackerAsync");
                if (_environment.IsDevelopment())
                {
                    return BadRequest(ex.ToString());
                }
                else
                {
                    return BadRequest();
                }

            }
            if (status)
            {
                return Ok(soxTracker);
            }
            else
            {
                return NoContent();
            }

        }

        [AllowAnonymous]
        [HttpPost("podioupdate")]
        public async Task<IActionResult> UpdatePodioSoxTrackerAsync([FromBody] SoxTracker soxTracker)
        {
            bool status = false;
            int _KeyReportID = 0; 
            try
            {
                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("SoxTrackerApp").GetSection("AppId").Value;
                PodioAppKey.AppToken = _config.GetSection("SoxTrackerApp").GetSection("AppToken").Value;

                if (PodioAppKey.AppId != string.Empty && PodioAppKey.AppToken != string.Empty)
                {
                    var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                    ServicePointManager.Expect100Continue = true;
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                    await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                    if (podio.IsAuthenticated() && soxTracker != null)
                    {

                        var SoxDatbaseItem = _soxContext.SoxTracker
                        .Where(x =>
                            x.FY.Equals(soxTracker.FY) &&
                            x.ClientName.Equals(soxTracker.ClientName) &&
                            x.Process.Equals(soxTracker.Process) &&
                            x.Subprocess.Equals(soxTracker.Subprocess) &&
                            x.ControlId.Equals(soxTracker.ControlId) &&
                            x.Status.Equals("Active")).FirstOrDefault();

                        var RCMDatbaseItem = _soxContext.SoxTracker.Where(x =>
                         x.FY.Equals(soxTracker.FY) &&
                         x.ClientName.Equals(soxTracker.ClientName) &&
                         x.Process.Equals(soxTracker.Process) &&
                         x.Subprocess.Equals(soxTracker.Subprocess) &&
                         x.ControlId.Equals(soxTracker.ControlId) &&
                         x.Status.Equals("Active")).FirstOrDefault();

                        //Console.WriteLine(SoxDatbaseItem.PodioItemId);

                        Item soxTrackerItem = new Item();
                        soxTrackerItem.ItemId = SoxDatbaseItem.PodioItemId;

                        int statusField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Status").Value);
                        var categoryStatus = soxTrackerItem.Field<CategoryItemField>(statusField);
                        categoryStatus.OptionText = "Active";

                        #region Podio Fields
                        PopulatePodioFields(soxTrackerItem, soxTracker, true);
                        #endregion


                        System.Diagnostics.Debug.WriteLine(soxTrackerItem);

                        var roundId = await podio.ItemService.UpdateItem(soxTrackerItem); //uncomment after dev
                        if (roundId != null)
                        {
                            soxTracker.PodioItemId = roundId.Value;
                            
                        }

                        if (soxTracker.KeyReportName != string.Empty && soxTracker.KeyReportName != null)
                        {

                            if (soxTracker.KeyReport == "Yes")
                            {

                                var KeyReportDatbaseItem = _soxContext.KeyReportUserInput
                               .Where(x =>
                               x.TagFY.Equals(soxTracker.FY) &&
                               x.TagClientName.Equals(soxTracker.ClientName) &&
                               x.TagReportName.Equals(soxTracker.KeyReportName) &&
                               x.TagControlId.Equals(soxTracker.ControlId) &&
                               x.TagStatus.Equals("Active")).FirstOrDefault();

                                //Console.WriteLine(KeyReportDatbaseItem);

                                if (KeyReportDatbaseItem == null)
                                {

                                    try
                                    {
                                        PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                                        PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                                        //=================================================================KeyReport Add Keyreport======================================================================

                                        PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("KeyReportNameId").Value;
                                        PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("KeyReportNameToken").Value;
                                        if (PodioAppKey.AppId != string.Empty && PodioAppKey.AppToken != string.Empty)
                                        {
                                            ServicePointManager.Expect100Continue = true;
                                            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                                            await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                                            if (podio.IsAuthenticated() && soxTracker != null)
                                            {

                                                Item KeyReportItem = new Item();

                                                //FY
                                                if (soxTracker.KeyReportName != string.Empty && soxTracker.KeyReportName != null)
                                                {
                                                    int KeyReportNameField = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportNameFieldId").GetSection("Field1").Value);
                                                    var KeyReportNameValue = KeyReportItem.Field<TextItemField>(KeyReportNameField);
                                                    KeyReportNameValue.Value = soxTracker.KeyReportName;
                                                }
                                                else
                                                {
                                                    int KeyReportNameField = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportNameFieldId").GetSection("Field1").Value);
                                                    var KeyReportNameValue = KeyReportItem.Field<TextItemField>(KeyReportNameField);
                                                    KeyReportNameValue.Value = " ";
                                                }
                                                if (soxTracker.Status != string.Empty && soxTracker.Status != null)
                                                {
                                                    soxTracker.Status = " ";
                                                }

                                                var KeyReportNameID = await podio.ItemService.AddNewItem(Int32.Parse(PodioAppKey.AppId), KeyReportItem); //uncomment after dev
                                                //Console.WriteLine("Key report ID =>", KeyReportNameID);
                                                soxTracker.PodioItemId = KeyReportNameID;
                                                _KeyReportID = KeyReportNameID;

                                                //Insert
                                                KeyReportUserInput _KeyReportUserInput = new KeyReportUserInput();
                                                //_KeyReportUserInput.CreatedOn = 
                                                _KeyReportUserInput.StrAnswer = "";
                                                _KeyReportUserInput.StrAnswer2 = "";
                                                _KeyReportUserInput.StrQuestion = "";
                                                _KeyReportUserInput.Description = "";
                                                _KeyReportUserInput.Position = 0;
                                                _KeyReportUserInput.AppId = PodioAppKey.AppId;
                                                _KeyReportUserInput.FieldId = 0;
                                                _KeyReportUserInput.ItemId = KeyReportNameID;
                                                _KeyReportUserInput.Type = "";
                                                _KeyReportUserInput.Tag = "";
                                                _KeyReportUserInput.Link = soxTracker.PodioLink;
                                                _KeyReportUserInput.TagFY = soxTracker.FY;
                                                _KeyReportUserInput.TagClientName = soxTracker.ClientName;
                                                _KeyReportUserInput.TagReportName = soxTracker.KeyReportName;
                                                _KeyReportUserInput.TagControlId = soxTracker.ControlId;
                                                _KeyReportUserInput.TagStatus = "Active";

                                                using (var context = _soxContext.Database.BeginTransaction())
                                                {
                                                    _soxContext.Add(_KeyReportUserInput);
                                                    await _soxContext.SaveChangesAsync();
                                                    context.Commit();
                                                }

                                                //Console.WriteLine("inserted");

                                            }
                                        }
                                        //===================================================================================================Consolidated App ===========================================================================================


                                        PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("ConsolOrigFormatId").Value;
                                        PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("ConsolOrigFormatToken").Value;

                                        if (PodioAppKey.AppId != string.Empty && PodioAppKey.AppToken != string.Empty)
                                        {
                                            ServicePointManager.Expect100Continue = true;
                                            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                                            await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                                            if (podio.IsAuthenticated() && soxTracker != null)
                                            {
                                                Item KeyReportConsolidated = new Item();

                                                //========fy=================//
                                                if (soxTracker.FY != string.Empty && soxTracker.FY != null)
                                                {
                                                    int FyField = int.Parse(_config.GetSection("KeyReportApp").GetSection("Consolidated_Fields").GetSection("FY").Value);
                                                    var FyFieldValue = KeyReportConsolidated.Field<TextItemField>(FyField);
                                                    FyFieldValue.Value = soxTracker.FY;
                                                }
                                                else
                                                {
                                                    int FyField = int.Parse(_config.GetSection("KeyReportApp").GetSection("Consolidated_Fields").GetSection("FY").Value);
                                                    var FyFieldValue = KeyReportConsolidated.Field<TextItemField>(FyField);
                                                    FyFieldValue.Value = " ";
                                                }


                                                //========screenshot=================//
                                                int Num2Screenshot = int.Parse(_config.GetSection("KeyReportApp").GetSection("Consolidated_Fields").GetSection("Screenshot").Value);
                                                ImageItemField imageField = KeyReportConsolidated.Field<ImageItemField>(Num2Screenshot);
                                                imageField.FileIds = new List<int>();



                                                //========Client Name=================//
                                                if (soxTracker.ClientName != string.Empty && soxTracker.ClientName != null)
                                                {
                                                    var checkClientItemId = _soxContext.ClientSs.Where(x => x.ClientName.Equals(soxTracker.ClientName)).Select(x => x.ItemId).FirstOrDefault();
                                                    //Console.WriteLine(checkClientItemId);
                                                    if (checkClientItemId != null)
                                                    {

                                                        int ClientNameField = int.Parse(_config.GetSection("KeyReportApp").GetSection("Consolidated_Fields").GetSection("ClientName").Value);
                                                        var ClientNameValue = KeyReportConsolidated.Field<AppItemField>(ClientNameField);
                                                        List<int> listClient = new List<int>();
                                                        listClient.Add(checkClientItemId.Value);
                                                        ClientNameValue.ItemIds = listClient;
                                                    }
                                                }
                                                else
                                                {
                                                    int ClientNameField = int.Parse(_config.GetSection("KeyReportApp").GetSection("Consolidated_Fields").GetSection("ClientName").Value);
                                                    var ClientNameValue = KeyReportConsolidated.Field<AppItemField>(ClientNameField);
                                                    List<int> listId = new List<int>();
                                                    ClientNameValue.ItemIds = listId;

                                                }


                                                // RCM CONTROL ID
                                                if (soxTracker.ControlId != string.Empty && soxTracker.ControlId != null)
                                                {
                                                    var checkRcmControlId = _soxContext.RcmControlId.Where(x => x.ControlId.Equals(soxTracker.ControlId)).FirstOrDefault();
                                                    if (checkRcmControlId != null)
                                                    {
                                                        int KeyReportId = int.Parse(_config.GetSection("KeyReportApp").GetSection("Consolidated_Fields").GetSection("KeyControlId").Value);
                                                        var controlIdRef = KeyReportConsolidated.Field<AppItemField>(KeyReportId);
                                                        List<int> listControlId = new List<int>();
                                                        listControlId.Add(checkRcmControlId.PodioItemId);
                                                        controlIdRef.ItemIds = listControlId;
                                                    }
                                                    else
                                                    {
                                                        int KeyReportId = int.Parse(_config.GetSection("KeyReportApp").GetSection("Consolidated_Fields").GetSection("KeyControlId").Value);
                                                        var controlIdRef = KeyReportConsolidated.Field<AppItemField>(KeyReportId);
                                                        List<int> listControlId = new List<int>();
                                                    }
                                                }


                                                int NameofIUC = int.Parse(_config.GetSection("KeyReportApp").GetSection("Consolidated_Fields").GetSection("KeyReport_IUC").Value);
                                                var NameofIUCref = KeyReportConsolidated.Field<AppItemField>(NameofIUC);
                                                List<int> listNameofIUC = new List<int>();
                                                listNameofIUC.Add(_KeyReportID);
                                                NameofIUCref.ItemIds = listNameofIUC;

                                                if (soxTracker.Process != string.Empty && soxTracker.Process != null)
                                                {

                                                    var checkRcmProcess = _soxContext.RcmProcess.Where(x => x.Process.Equals(soxTracker.Process)).FirstOrDefault();
                                                    if (checkRcmProcess != null)
                                                    {
                                                        int Process = int.Parse(_config.GetSection("KeyReportApp").GetSection("Consolidated_Fields").GetSection("SourceProcess").Value);
                                                        var processRef = KeyReportConsolidated.Field<AppItemField>(Process);
                                                        List<int> listProcess = new List<int>();
                                                        listProcess.Add(checkRcmProcess.PodioItemId);
                                                        processRef.ItemIds = listProcess;
                                                    }
                                                }

                                                var ConsolidatedID = await podio.ItemService.AddNewItem(Int32.Parse(PodioAppKey.AppId), KeyReportConsolidated); //uncomment after dev

                                                //Console.WriteLine("ConsolidatedID =>", ConsolidatedID);
                                                soxTracker.PodioItemId = ConsolidatedID;

                                                //Insert
                                                KeyReportOrigFormat _KeyReportOrigFormat = new KeyReportOrigFormat();

                                                _KeyReportOrigFormat.ClientName = soxTracker.ClientName;
                                                _KeyReportOrigFormat.ClientCode = soxTracker.ClientCode;
                                                _KeyReportOrigFormat.ClientItemId = soxTracker.ClientItemId;
                                                _KeyReportOrigFormat.Num = 0;
                                                _KeyReportOrigFormat.KeyControlId = soxTracker.ControlId;
                                                _KeyReportOrigFormat.ControlActivity = "";
                                                _KeyReportOrigFormat.SKey = soxTracker.KeyReport;
                                                _KeyReportOrigFormat.NameOfKeyReport = soxTracker.KeyReportName;
                                                _KeyReportOrigFormat.KeyReport = soxTracker.KeyReport;
                                                _KeyReportOrigFormat.IUCType = "";
                                                _KeyReportOrigFormat.Source = soxTracker.PodioLink;
                                                _KeyReportOrigFormat.ReportCustomized = "";
                                                _KeyReportOrigFormat.ControlRelyingIUC = "";
                                                _KeyReportOrigFormat.Preparer = soxTracker.R1Tester;
                                                _KeyReportOrigFormat.Reviewer = soxTracker.WT1LReviewer;
                                                _KeyReportOrigFormat.Notes = "Active";
                                                _KeyReportOrigFormat.PodioItemId = ConsolidatedID;
                                                _KeyReportOrigFormat.PodioUniqueId = soxTracker.PodioUniqueId;
                                                _KeyReportOrigFormat.PodioRevision = soxTracker.PodioRevision;
                                                _KeyReportOrigFormat.PodioLink = soxTracker.PodioLink;
                                                _KeyReportOrigFormat.CreatedBy = soxTracker.CreatedBy;


                                                using (var context = _soxContext.Database.BeginTransaction())
                                                {
                                                    _soxContext.Add(_KeyReportOrigFormat);
                                                    await _soxContext.SaveChangesAsync();
                                                    context.Commit();
                                                }

                                                //Console.WriteLine("_KeyReportOrigFormat inserted");
                                            }
                                        }


                                        //========================================================================================AllUIC App =================================================================================
                                        PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("AllIUCId").Value;
                                        PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("AllIUCToken").Value;

                                        if (PodioAppKey.AppId != string.Empty && PodioAppKey.AppToken != string.Empty)
                                        {
                                            ServicePointManager.Expect100Continue = true;
                                            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                                            await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                                            if (podio.IsAuthenticated() && soxTracker != null)
                                            {
                                                Item KeyReportAllIUC = new Item();
                                                if (soxTracker.FY != string.Empty && soxTracker.FY != null)
                                                {
                                                    int FyField = int.Parse(_config.GetSection("KeyReportApp").GetSection("AllIUC_Fields").GetSection("FY").Value);
                                                    var FyFieldValue = KeyReportAllIUC.Field<TextItemField>(FyField);
                                                    FyFieldValue.Value = soxTracker.FY;
                                                }
                                                else
                                                {
                                                    int FyField = int.Parse(_config.GetSection("KeyReportApp").GetSection("AllIUC_Fields").GetSection("FY").Value);
                                                    var FyFieldValue = KeyReportAllIUC.Field<TextItemField>(FyField);
                                                    FyFieldValue.Value = " ";
                                                }

                                                int Num2Screenshot = int.Parse(_config.GetSection("KeyReportApp").GetSection("AllIUC_Fields").GetSection("Screenshot").Value);
                                                ImageItemField imageField = KeyReportAllIUC.Field<ImageItemField>(Num2Screenshot);
                                                imageField.FileIds = new List<int>();

                                                if (soxTracker.ClientName != string.Empty && soxTracker.ClientName != null)
                                                {
                                                    var checkClientItemId = _soxContext.ClientSs.Where(x => x.ClientName.Equals(soxTracker.ClientName)).Select(x => x.ItemId).FirstOrDefault();
                                                    //Console.WriteLine(checkClientItemId);
                                                    if (checkClientItemId != null)
                                                    {

                                                        int ClientNameField = int.Parse(_config.GetSection("KeyReportApp").GetSection("AllIUC_Fields").GetSection("ClientName").Value);
                                                        var ClientNameValue = KeyReportAllIUC.Field<AppItemField>(ClientNameField);
                                                        List<int> listClient = new List<int>();
                                                        listClient.Add(checkClientItemId.Value);
                                                        ClientNameValue.ItemIds = listClient;
                                                    }
                                                }
                                                else
                                                {
                                                    int ClientNameField = int.Parse(_config.GetSection("KeyReportApp").GetSection("AllIUC_Fields").GetSection("ClientName").Value);
                                                    var ClientNameValue = KeyReportAllIUC.Field<AppItemField>(ClientNameField);
                                                    List<int> listId = new List<int>();
                                                    ClientNameValue.ItemIds = listId;

                                                }

                                                if (soxTracker.ControlId != string.Empty && soxTracker.ControlId != null)
                                                {
                                                    var checkRcmControlId = _soxContext.RcmControlId.Where(x => x.ControlId.Equals(soxTracker.ControlId)).FirstOrDefault();
                                                    if (checkRcmControlId != null)
                                                    {
                                                        int KeyReportId = int.Parse(_config.GetSection("KeyReportApp").GetSection("AllIUC_Fields").GetSection("KeyControl").Value);
                                                        var controlIdRef = KeyReportAllIUC.Field<AppItemField>(KeyReportId);
                                                        List<int> listControlId = new List<int>();
                                                        listControlId.Add(checkRcmControlId.PodioItemId);
                                                        controlIdRef.ItemIds = listControlId;
                                                    }
                                                    else
                                                    {
                                                        int KeyReportId = int.Parse(_config.GetSection("KeyReportApp").GetSection("AllIUC_Fields").GetSection("KeyControl").Value);
                                                        var controlIdRef = KeyReportAllIUC.Field<AppItemField>(KeyReportId);
                                                        List<int> listControlId = new List<int>();
                                                    }
                                                }


                                                int NameofIUC = int.Parse(_config.GetSection("KeyReportApp").GetSection("AllIUC_Fields").GetSection("KeyReport/IUC").Value);
                                                var NameofIUCref = KeyReportAllIUC.Field<AppItemField>(NameofIUC);
                                                List<int> listNameofIUC = new List<int>();
                                                listNameofIUC.Add(_KeyReportID);
                                                NameofIUCref.ItemIds = listNameofIUC;

                                                if (soxTracker.KeyReport != string.Empty && soxTracker.KeyReport != null) {

                                                    var KeyreportYesNo = _soxContext.KeyReportUniqueKeyReport.Where(x => x.UniqueKeyReport.Equals(soxTracker.KeyReport)).FirstOrDefault();
                                                    if (KeyreportYesNo != null)
                                                    {
                                                        int UniqueKeyReport = int.Parse(_config.GetSection("KeyReportApp").GetSection("AllIUC_Fields").GetSection("KeyReport").Value);
                                                        var UniqueKeyReportref = KeyReportAllIUC.Field<AppItemField>(UniqueKeyReport);
                                                        List<int> ListUniqueKeyReport = new List<int>();
                                                        ListUniqueKeyReport.Add(KeyreportYesNo.PodioItemId);
                                                        UniqueKeyReportref.ItemIds = ListUniqueKeyReport;
                                                    }
                                                    else {
                                                        int UniqueKeyReport = int.Parse(_config.GetSection("KeyReportApp").GetSection("AllIUC_Fields").GetSection("KeyReport").Value);
                                                        var UniqueKeyReportref = KeyReportAllIUC.Field<AppItemField>(UniqueKeyReport);
                                                        List<int> ListUniqueKeyReport = new List<int>();
                                                    }
                                                }

                                                if (soxTracker.Process != string.Empty && soxTracker.Process != null)
                                                {

                                                    var checkRcmProcess = _soxContext.RcmProcess.Where(x => x.Process.Equals(soxTracker.Process)).FirstOrDefault();
                                                    if (checkRcmProcess != null)
                                                    {
                                                        int Process = int.Parse(_config.GetSection("KeyReportApp").GetSection("AllIUC_Fields").GetSection("SourceProcess").Value);
                                                        var processRef = KeyReportAllIUC.Field<AppItemField>(Process);
                                                        List<int> listProcess = new List<int>();
                                                        listProcess.Add(checkRcmProcess.PodioItemId);
                                                        processRef.ItemIds = listProcess;
                                                    }
                                                }

                                                var KeyReportAllIUCID = 0;
                                                try {

                                                    KeyReportAllIUCID = await podio.ItemService.AddNewItem(Int32.Parse(PodioAppKey.AppId), KeyReportAllIUC); //uncomment after dev
                                                                                                                                                                 //Console.WriteLine("KeyReportAllIUCID =>", KeyReportAllIUCID);
                                                    soxTracker.PodioItemId = KeyReportAllIUCID;
                                                }
                                                catch {
                                                    Item KeyReportAllIUC2 = new Item();


                                                    if (soxTracker.FY != string.Empty && soxTracker.FY != null)
                                                    {
                                                        int FyField = int.Parse(_config.GetSection("KeyReportApp").GetSection("AllIUC_Fields").GetSection("FY").Value);
                                                        var FyFieldValue = KeyReportAllIUC2.Field<TextItemField>(FyField);
                                                        FyFieldValue.Value = soxTracker.FY;
                                                    }
                                                    else
                                                    {
                                                        int FyField = int.Parse(_config.GetSection("KeyReportApp").GetSection("AllIUC_Fields").GetSection("FY").Value);
                                                        var FyFieldValue = KeyReportAllIUC2.Field<TextItemField>(FyField);
                                                        FyFieldValue.Value = " ";
                                                    }


                                                    if (soxTracker.ClientName != string.Empty && soxTracker.ClientName != null)
                                                    {
                                                        var checkClientItemId = _soxContext.ClientSs.Where(x => x.ClientName.Equals(soxTracker.ClientName)).Select(x => x.ItemId).FirstOrDefault();
                                                        //Console.WriteLine(checkClientItemId);
                                                        if (checkClientItemId != null)
                                                        {

                                                            int ClientNameField = int.Parse(_config.GetSection("KeyReportApp").GetSection("AllIUC_Fields").GetSection("ClientName").Value);
                                                            var ClientNameValue = KeyReportAllIUC2.Field<AppItemField>(ClientNameField);
                                                            List<int> listClient = new List<int>();
                                                            listClient.Add(checkClientItemId.Value);
                                                            ClientNameValue.ItemIds = listClient;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        int ClientNameField = int.Parse(_config.GetSection("KeyReportApp").GetSection("AllIUC_Fields").GetSection("ClientName").Value);
                                                        var ClientNameValue = KeyReportAllIUC2.Field<AppItemField>(ClientNameField);
                                                        List<int> listId = new List<int>();
                                                        ClientNameValue.ItemIds = listId;

                                                    }

                                                    if (soxTracker.ControlId != string.Empty && soxTracker.ControlId != null)
                                                    {
                                                        var checkRcmControlId = _soxContext.RcmControlId.Where(x => x.ControlId.Equals(soxTracker.ControlId)).FirstOrDefault();
                                                        if (checkRcmControlId != null)
                                                        {
                                                            int KeyReportId = int.Parse(_config.GetSection("KeyReportApp").GetSection("AllIUC_Fields").GetSection("KeyControl").Value);
                                                            var controlIdRef = KeyReportAllIUC2.Field<AppItemField>(KeyReportId);
                                                            List<int> listControlId = new List<int>();
                                                            listControlId.Add(checkRcmControlId.PodioItemId);
                                                            controlIdRef.ItemIds = listControlId;
                                                        }
                                                        else
                                                        {
                                                            int KeyReportId = int.Parse(_config.GetSection("KeyReportApp").GetSection("AllIUC_Fields").GetSection("KeyControl").Value);
                                                            var controlIdRef = KeyReportAllIUC2.Field<AppItemField>(KeyReportId);
                                                            List<int> listControlId = new List<int>();
                                                        }
                                                    }

                                                    
                                                    
                                                    int NameofIUC2 = int.Parse(_config.GetSection("KeyReportApp").GetSection("AllIUC_Fields").GetSection("KeyReport/IUC").Value);
                                                    var NameofIUCref2 = KeyReportAllIUC2.Field<AppItemField>(NameofIUC2);
                                                    List<int> listNameofIUC2 = new List<int>();
                                                    listNameofIUC2.Add(_KeyReportID);
                                                    NameofIUCref2.ItemIds = listNameofIUC2;



                                                    if (soxTracker.KeyReport != string.Empty && soxTracker.KeyReport != null)
                                                    {

                                                        var KeyreportYesNo = _soxContext.KeyReportUniqueKeyReport.Where(x => x.UniqueKeyReport.Equals(soxTracker.KeyReport)).FirstOrDefault();
                                                        if (KeyreportYesNo != null)
                                                        {

                                                            int UniqueKeyReport2 = int.Parse(_config.GetSection("KeyReportApp").GetSection("AllIUC_Fields").GetSection("KeyReport").Value);
                                                            var UniqueKeyReportref2 = KeyReportAllIUC2.Field<AppItemField>(UniqueKeyReport2);
                                                            List<int> ListUniqueKeyReport2 = new List<int>();
                                                            ListUniqueKeyReport2.Add(KeyreportYesNo.PodioItemId);
                                                            UniqueKeyReportref2.ItemIds = ListUniqueKeyReport2;
                                                        }
                                                        else
                                                        {
                                                            int UniqueKeyReport2 = int.Parse(_config.GetSection("KeyReportApp").GetSection("AllIUC_Fields").GetSection("KeyReport").Value);
                                                            var UniqueKeyReportref2 = KeyReportAllIUC.Field<AppItemField>(UniqueKeyReport2);
                                                            List<int> ListUniqueKeyReport2 = new List<int>();
                                                        }
                                                    }


                                                    

                                                    

                                                    KeyReportAllIUCID = await podio.ItemService.AddNewItem(Int32.Parse(PodioAppKey.AppId), KeyReportAllIUC2); //uncomment after dev
                                                                                                                                                                 //Console.WriteLine("KeyReportAllIUCID =>", KeyReportAllIUCID);
                                                    soxTracker.PodioItemId = KeyReportAllIUCID;
                                                }
                                                

                                                //Insert
                                                KeyReportAllIUC _KeyReportAllIUC = new KeyReportAllIUC();

                                                _KeyReportAllIUC.ClientName = soxTracker.ClientName;
                                                _KeyReportAllIUC.ClientCode = soxTracker.ClientCode;
                                                _KeyReportAllIUC.ClientItemId = soxTracker.ClientItemId;
                                                _KeyReportAllIUC.Num = 0;
                                                _KeyReportAllIUC.KeyControlId = soxTracker.ControlId;
                                                _KeyReportAllIUC.ControlActivity = ""; //
                                                _KeyReportAllIUC.SKey = soxTracker.KeyReport;
                                                _KeyReportAllIUC.NameOfIUC = soxTracker.KeyReportName;
                                                _KeyReportAllIUC.KeyReport = soxTracker.KeyReport;
                                                _KeyReportAllIUC.UniqueKeyReport = soxTracker.KeyReport; // should be yes referrence on podio
                                                _KeyReportAllIUC.PodioItemId = KeyReportAllIUCID;
                                                _KeyReportAllIUC.Process = soxTracker.Process;
                                                _KeyReportAllIUC.PodioUniqueId = soxTracker.PodioUniqueId;
                                                _KeyReportAllIUC.PodioRevision = soxTracker.PodioRevision;
                                                _KeyReportAllIUC.PodioLink = soxTracker.PodioLink;
                                                _KeyReportAllIUC.CreatedBy = soxTracker.CreatedBy;

                                                using (var context = _soxContext.Database.BeginTransaction())
                                                {
                                                    _soxContext.Add(_KeyReportAllIUC);
                                                    await _soxContext.SaveChangesAsync();
                                                    context.Commit();
                                                }
                                                //Console.WriteLine("_KeyReportAllIUC inserted");

                                            }
                                        }

                                        //========================================================================================Test Status Tracker App =================================================================================
                                        PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("TestStatusTrackerId").Value;
                                        PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("TestStatusTrackerToken").Value;

                                        if (PodioAppKey.AppId != string.Empty && PodioAppKey.AppToken != string.Empty)
                                        {
                                            ServicePointManager.Expect100Continue = true;
                                            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                                            await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                                            if (podio.IsAuthenticated() && soxTracker != null)
                                            {
                                                Item KeyReportTestStatus = new Item();

                                                if (soxTracker.FY != string.Empty && soxTracker.FY != null)
                                                {
                                                    int FyField = int.Parse(_config.GetSection("KeyReportApp").GetSection("TestStatusTracker_Fields").GetSection("FY").Value);
                                                    var FyFieldValue = KeyReportTestStatus.Field<TextItemField>(FyField);
                                                    FyFieldValue.Value = soxTracker.FY;
                                                }
                                                else
                                                {
                                                    int FyField = int.Parse(_config.GetSection("KeyReportApp").GetSection("TestStatusTracker_Fields").GetSection("FY").Value);
                                                    var FyFieldValue = KeyReportTestStatus.Field<TextItemField>(FyField);
                                                    FyFieldValue.Value = " ";
                                                }

                                                int Num2Screenshot = int.Parse(_config.GetSection("KeyReportApp").GetSection("TestStatusTracker_Fields").GetSection("1Screenshot").Value);
                                                ImageItemField imageField = KeyReportTestStatus.Field<ImageItemField>(Num2Screenshot);
                                                imageField.FileIds = new List<int>();

                                                if (soxTracker.ClientName != string.Empty && soxTracker.ClientName != null)
                                                {
                                                    var checkClientItemId = _soxContext.ClientSs.Where(x => x.ClientName.Equals(soxTracker.ClientName)).Select(x => x.ItemId).FirstOrDefault();
                                                    //Console.WriteLine(checkClientItemId);
                                                    if (checkClientItemId != null)
                                                    {

                                                        int ClientNameField = int.Parse(_config.GetSection("KeyReportApp").GetSection("TestStatusTracker_Fields").GetSection("ClientName").Value);
                                                        var ClientNameValue = KeyReportTestStatus.Field<AppItemField>(ClientNameField);
                                                        List<int> listClient = new List<int>();
                                                        listClient.Add(checkClientItemId.Value);
                                                        ClientNameValue.ItemIds = listClient;
                                                    }
                                                }
                                                else
                                                {
                                                    int ClientNameField = int.Parse(_config.GetSection("KeyReportApp").GetSection("TestStatusTracker_Fields").GetSection("ClientName").Value);
                                                    var ClientNameValue = KeyReportTestStatus.Field<AppItemField>(ClientNameField);
                                                    List<int> listId = new List<int>();
                                                    ClientNameValue.ItemIds = listId;

                                                }

                                                if (soxTracker.ControlId != string.Empty && soxTracker.ControlId != null)
                                                {
                                                    var checkRcmControlId = _soxContext.RcmControlId.Where(x => x.ControlId.Equals(soxTracker.ControlId)).FirstOrDefault();
                                                    if (checkRcmControlId != null)
                                                    {
                                                        int KeyReportId = int.Parse(_config.GetSection("KeyReportApp").GetSection("TestStatusTracker_Fields").GetSection("KeyControlID").Value);
                                                        var controlIdRef = KeyReportTestStatus.Field<AppItemField>(KeyReportId);
                                                        List<int> listControlId = new List<int>();
                                                        listControlId.Add(checkRcmControlId.PodioItemId);
                                                        controlIdRef.ItemIds = listControlId;
                                                    }
                                                    else
                                                    {
                                                        int KeyReportId = int.Parse(_config.GetSection("KeyReportApp").GetSection("TestStatusTracker_Fields").GetSection("KeyControlID").Value);
                                                        var controlIdRef = KeyReportTestStatus.Field<AppItemField>(KeyReportId);
                                                        List<int> listControlId = new List<int>();
                                                    }
                                                }


                                                int NameofIUC = int.Parse(_config.GetSection("KeyReportApp").GetSection("TestStatusTracker_Fields").GetSection("KeyReport/IUC").Value);
                                                var NameofIUCref = KeyReportTestStatus.Field<AppItemField>(NameofIUC);
                                                List<int> listNameofIUC = new List<int>();
                                                listNameofIUC.Add(_KeyReportID);
                                                NameofIUCref.ItemIds = listNameofIUC;

                                                if (soxTracker.Process != string.Empty && soxTracker.Process != null)
                                                {

                                                    var checkRcmProcess = _soxContext.RcmProcess.Where(x => x.Process.Equals(soxTracker.Process)).FirstOrDefault();
                                                    if (checkRcmProcess != null)
                                                    {
                                                        int Process = int.Parse(_config.GetSection("KeyReportApp").GetSection("TestStatusTracker_Fields").GetSection("SourceProcess").Value);
                                                        var processRef = KeyReportTestStatus.Field<AppItemField>(Process);
                                                        List<int> listProcess = new List<int>();
                                                        listProcess.Add(checkRcmProcess.PodioItemId);
                                                        processRef.ItemIds = listProcess;
                                                    }
                                                }

                                                var TestStatusTrackerID = await podio.ItemService.AddNewItem(Int32.Parse(PodioAppKey.AppId), KeyReportTestStatus); //uncomment after dev
                                                //Console.WriteLine("TestStatusTracker ID =>", TestStatusTrackerID);
                                                soxTracker.PodioItemId = TestStatusTrackerID;


                                                //Insert 
                                                KeyReportTestStatusTracker _KeyReportTestStatusTracker = new KeyReportTestStatusTracker();

                                                _KeyReportTestStatusTracker.ClientName = soxTracker.ClientName;
                                                _KeyReportTestStatusTracker.ClientCode = soxTracker.ClientCode;
                                                _KeyReportTestStatusTracker.ClientItemId = soxTracker.ClientItemId;
                                                _KeyReportTestStatusTracker.NameOfIUC = soxTracker.KeyReportName;
                                                _KeyReportTestStatusTracker.Process = soxTracker.Process;
                                                _KeyReportTestStatusTracker.KeyControlId = soxTracker.ControlId;
                                                _KeyReportTestStatusTracker.KeyReport = soxTracker.KeyReport;
                                                _KeyReportTestStatusTracker.UniqueKeyReport = soxTracker.KeyReport;
                                                _KeyReportTestStatusTracker.PBCStatus = soxTracker.Status;
                                                _KeyReportTestStatusTracker.TestingStatus = soxTracker.WTTestingStatus;
                                                _KeyReportTestStatusTracker.PodioItemId = TestStatusTrackerID;
                                                _KeyReportTestStatusTracker.PodioUniqueId = soxTracker.PodioUniqueId; // should be yes referrence on podio
                                                _KeyReportTestStatusTracker.PodioRevision = soxTracker.PodioRevision;
                                                _KeyReportTestStatusTracker.PodioLink = soxTracker.PodioLink;
                                                _KeyReportTestStatusTracker.CreatedBy = soxTracker.CreatedBy;

                                                using (var context = _soxContext.Database.BeginTransaction())
                                                {
                                                    _soxContext.Add(_KeyReportTestStatusTracker);
                                                    await _soxContext.SaveChangesAsync();
                                                    context.Commit();
                                                }
                                                //Console.WriteLine("_KeyReportTestStatusTracker inserted");

                                            }
                                        }

                                        //========================================================================================Exception App =================================================================================
                                        PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("ExceptionId").Value;
                                        PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("ExceptionToken").Value;

                                        if (PodioAppKey.AppId != string.Empty && PodioAppKey.AppToken != string.Empty)
                                        {
                                            ServicePointManager.Expect100Continue = true;
                                            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                                            await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                                            if (podio.IsAuthenticated() && soxTracker != null)
                                            {
                                                Item Exception = new Item
                                                {
                                                    ItemId = soxTracker.PodioItemId
                                                };
                                                if (soxTracker.FY != string.Empty && soxTracker.FY != null)
                                                {
                                                    int FyField = int.Parse(_config.GetSection("KeyReportApp").GetSection("Exceptions_Fields").GetSection("FY").Value);
                                                    var FyFieldValue = Exception.Field<TextItemField>(FyField);
                                                    FyFieldValue.Value = soxTracker.FY;
                                                }
                                                else
                                                {
                                                    int FyField = int.Parse(_config.GetSection("KeyReportApp").GetSection("Exceptions_Fields").GetSection("FY").Value);
                                                    var FyFieldValue = Exception.Field<TextItemField>(FyField);
                                                    FyFieldValue.Value = " ";
                                                }

                                                int Num2Screenshot = int.Parse(_config.GetSection("KeyReportApp").GetSection("Exceptions_Fields").GetSection("1Screenshot").Value);
                                                ImageItemField imageField = Exception.Field<ImageItemField>(Num2Screenshot);
                                                imageField.FileIds = new List<int>();

                                                if (soxTracker.ClientName != string.Empty && soxTracker.ClientName != null)
                                                {
                                                    var checkClientItemId = _soxContext.ClientSs.Where(x => x.ClientName.Equals(soxTracker.ClientName)).Select(x => x.ItemId).FirstOrDefault();
                                                    //Console.WriteLine(checkClientItemId);
                                                    if (checkClientItemId != null)
                                                    {

                                                        int ClientNameField = int.Parse(_config.GetSection("KeyReportApp").GetSection("Exceptions_Fields").GetSection("ClientName").Value);
                                                        var ClientNameValue = Exception.Field<AppItemField>(ClientNameField);
                                                        List<int> listClient = new List<int>();
                                                        listClient.Add(checkClientItemId.Value);
                                                        ClientNameValue.ItemIds = listClient;
                                                    }
                                                }
                                                else
                                                {
                                                    int ClientNameField = int.Parse(_config.GetSection("KeyReportApp").GetSection("Exceptions_Fields").GetSection("ClientName").Value);
                                                    var ClientNameValue = Exception.Field<AppItemField>(ClientNameField);
                                                    List<int> listId = new List<int>();
                                                    ClientNameValue.ItemIds = listId;

                                                }

                                                if (soxTracker.ControlId != string.Empty && soxTracker.ControlId != null)
                                                {
                                                    var checkRcmControlId = _soxContext.RcmControlId.Where(x => x.ControlId.Equals(soxTracker.ControlId)).FirstOrDefault();
                                                    if (checkRcmControlId != null)
                                                    {
                                                        int KeyReportId = int.Parse(_config.GetSection("KeyReportApp").GetSection("Exceptions_Fields").GetSection("KeyControlID").Value);
                                                        var controlIdRef = Exception.Field<AppItemField>(KeyReportId);
                                                        List<int> listControlId = new List<int>();
                                                        listControlId.Add(checkRcmControlId.PodioItemId);
                                                        controlIdRef.ItemIds = listControlId;
                                                    }
                                                    else
                                                    {
                                                        int KeyReportId = int.Parse(_config.GetSection("KeyReportApp").GetSection("Exceptions_Fields").GetSection("KeyControlID").Value);
                                                        var controlIdRef = Exception.Field<AppItemField>(KeyReportId);
                                                        List<int> listControlId = new List<int>();
                                                    }
                                                }


                                                int NameofIUC = int.Parse(_config.GetSection("KeyReportApp").GetSection("Exceptions_Fields").GetSection("KeyReport/IUC").Value);
                                                var NameofIUCref = Exception.Field<AppItemField>(NameofIUC);
                                                List<int> listNameofIUC = new List<int>();
                                                listNameofIUC.Add(_KeyReportID);
                                                NameofIUCref.ItemIds = listNameofIUC;

                                                var ExceptionID = await podio.ItemService.AddNewItem(Int32.Parse(PodioAppKey.AppId), Exception);// uncomment after dev
                                                //Console.WriteLine("TestStatusTracker ID =>", ExceptionID);
                                                soxTracker.PodioItemId = ExceptionID;

                                                //Insert 
                                                KeyReportExcepcion _KeyReportExcepcion = new KeyReportExcepcion();

                                                _KeyReportExcepcion.ClientName = soxTracker.ClientName;
                                                _KeyReportExcepcion.ClientCode = soxTracker.ClientCode;
                                                _KeyReportExcepcion.ClientItemId = soxTracker.ClientItemId;
                                                _KeyReportExcepcion.NameOfIUC = soxTracker.KeyReportName;
                                                _KeyReportExcepcion.KeyControlId = soxTracker.ControlId;
                                                _KeyReportExcepcion.PodioItemId = ExceptionID;
                                                _KeyReportExcepcion.PodioUniqueId = soxTracker.PodioUniqueId; // should be yes referrence on podio
                                                _KeyReportExcepcion.PodioRevision = soxTracker.PodioRevision;
                                                _KeyReportExcepcion.PodioLink = soxTracker.PodioLink;
                                                _KeyReportExcepcion.CreatedBy = soxTracker.CreatedBy;

                                                using (var context = _soxContext.Database.BeginTransaction())
                                                {
                                                    _soxContext.Add(_KeyReportExcepcion);
                                                    await _soxContext.SaveChangesAsync();
                                                    context.Commit();
                                                }
                                                //Console.WriteLine("_KeyReportExcepcion inserted");
                                            }
                                        }





                                    }
                                    catch (Exception ex)
                                    {
                                        //Console.WriteLine(ex.ToString());
                                        //ErrorLog.Write(ex);
                                        FileLog.Write(ex.ToString(), "ErrorUpdatePodioSoxTrackerAsync");
                                        AdminService adminService = new AdminService(_config);
                                        adminService.SendAlert(true, true, ex.ToString(), "UpdatePodioSoxTrackerAsync");
                                        if (_environment.IsDevelopment())
                                        {
                                            //Console.WriteLine(BadRequest(ex.ToString()));
                                            return BadRequest(ex.ToString());
                                        }
                                        else
                                        {
                                            BadRequest();
                                            //Console.WriteLine(BadRequest());
                                            return BadRequest();
                                        }

                                    }
                                }
                                else
                                { //Update Condition here//Update Condition here//Update Condition here//Update Condition here//Update Condition here//Update Condition here
                                    try
                                    {
                                        PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                                        PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;

                                        //==============================================================================KeyReport Update Keyreport==================================================================================

                                        PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("KeyReportNameId").Value;
                                        PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("KeyReportNameToken").Value;
                                        if (PodioAppKey.AppId != string.Empty && PodioAppKey.AppToken != string.Empty)
                                        {
                                            ServicePointManager.Expect100Continue = true;
                                            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                                            await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                                            if (podio.IsAuthenticated() && soxTracker != null)
                                            {
                                                Item KeyReportItem = new Item();


                                                //FY
                                                if (soxTracker.KeyReportName != string.Empty && soxTracker.KeyReportName != null)
                                                {
                                                    int KeyReportNameField = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportNameFieldId").GetSection("Field1").Value);
                                                    var KeyReportNameValue = KeyReportItem.Field<TextItemField>(KeyReportNameField);
                                                    KeyReportNameValue.Value = soxTracker.KeyReportName;
                                                }
                                                else
                                                {
                                                    int KeyReportNameField = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportNameFieldId").GetSection("Field1").Value);
                                                    var KeyReportNameValue = KeyReportItem.Field<TextItemField>(KeyReportNameField);
                                                    KeyReportNameValue.Value = " ";
                                                }

                                                if (soxTracker.Status != string.Empty && soxTracker.Status != null)
                                                {
                                                    soxTracker.Status = KeyReportDatbaseItem.TagStatus;
                                                }

                                                //Search update
                                                KeyReportUserInput _KeyReportUserInput = new KeyReportUserInput();

                                                _KeyReportUserInput.StrAnswer = "";
                                                _KeyReportUserInput.StrAnswer2 = "";
                                                _KeyReportUserInput.StrQuestion = "";
                                                _KeyReportUserInput.Description = "";
                                                _KeyReportUserInput.Position = 0;
                                                _KeyReportUserInput.AppId = PodioAppKey.AppId;
                                                _KeyReportUserInput.FieldId = 0;
                                                _KeyReportUserInput.ItemId = KeyReportDatbaseItem.ItemId;
                                                _KeyReportUserInput.Type = "";
                                                _KeyReportUserInput.Tag = "";
                                                _KeyReportUserInput.Link = soxTracker.PodioLink;
                                                _KeyReportUserInput.TagFY = soxTracker.FY;
                                                _KeyReportUserInput.TagClientName = soxTracker.ClientName;
                                                _KeyReportUserInput.TagReportName = soxTracker.KeyReportName;
                                                _KeyReportUserInput.TagControlId = soxTracker.ControlId;
                                                _KeyReportUserInput.TagStatus = "Active";

                                                using (var context = _soxContext.Database.BeginTransaction())
                                                {
                                                    _KeyReportUserInput.Id = KeyReportDatbaseItem.Id;
                                                    _soxContext.Entry(KeyReportDatbaseItem).CurrentValues.SetValues(_KeyReportUserInput);
                                                    await _soxContext.SaveChangesAsync();
                                                    context.Commit();
                                                }

                                                KeyReportItem.ItemId = KeyReportDatbaseItem.ItemId;
                                                var KeyReportNameID = await podio.ItemService.UpdateItem(KeyReportItem); //uncomment after dev
                                                _KeyReportID = KeyReportDatbaseItem.ItemId;
                                            }
                                        }

                                        //===================================================================================================Update Consolidated App ===========================================================================================

                                        PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("ConsolOrigFormatId").Value;
                                        PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("ConsolOrigFormatToken").Value;

                                        if (PodioAppKey.AppId != string.Empty && PodioAppKey.AppToken != string.Empty)
                                        {
                                            ServicePointManager.Expect100Continue = true;
                                            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                                            await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                                            if (podio.IsAuthenticated() && soxTracker != null)
                                            {
                                                Item KeyReportConsolidated = new Item();


                                                //========fy=================//
                                                if (soxTracker.FY != string.Empty && soxTracker.FY != null)
                                                {
                                                    int FyField = int.Parse(_config.GetSection("KeyReportApp").GetSection("Consolidated_Fields").GetSection("FY").Value);
                                                    var FyFieldValue = KeyReportConsolidated.Field<TextItemField>(FyField);
                                                    FyFieldValue.Value = soxTracker.FY;
                                                }
                                                else
                                                {
                                                    int FyField = int.Parse(_config.GetSection("KeyReportApp").GetSection("Consolidated_Fields").GetSection("FY").Value);
                                                    var FyFieldValue = KeyReportConsolidated.Field<TextItemField>(FyField);
                                                    FyFieldValue.Value = " ";
                                                }


                                                //========screenshot=================//
                                                int Num2Screenshot = int.Parse(_config.GetSection("KeyReportApp").GetSection("Consolidated_Fields").GetSection("Screenshot").Value);
                                                ImageItemField imageField = KeyReportConsolidated.Field<ImageItemField>(Num2Screenshot);
                                                imageField.FileIds = new List<int>();



                                                //========Client Name=================//
                                                if (soxTracker.ClientName != string.Empty && soxTracker.ClientName != null)
                                                {
                                                    var checkClientItemId = _soxContext.ClientSs.Where(x => x.ClientName.Equals(soxTracker.ClientName)).Select(x => x.ItemId).FirstOrDefault();
                                                    //Console.WriteLine(checkClientItemId);
                                                    if (checkClientItemId != null)
                                                    {

                                                        int ClientNameField = int.Parse(_config.GetSection("KeyReportApp").GetSection("Consolidated_Fields").GetSection("ClientName").Value);
                                                        var ClientNameValue = KeyReportConsolidated.Field<AppItemField>(ClientNameField);
                                                        List<int> listClient = new List<int>();
                                                        listClient.Add(checkClientItemId.Value);
                                                        ClientNameValue.ItemIds = listClient;
                                                    }
                                                }
                                                else
                                                {
                                                    int ClientNameField = int.Parse(_config.GetSection("KeyReportApp").GetSection("Consolidated_Fields").GetSection("ClientName").Value);
                                                    var ClientNameValue = KeyReportConsolidated.Field<AppItemField>(ClientNameField);
                                                    List<int> listId = new List<int>();
                                                    ClientNameValue.ItemIds = listId;

                                                }


                                                // RCM CONTROL ID
                                                if (soxTracker.ControlId != string.Empty && soxTracker.ControlId != null)
                                                {
                                                    var checkRcmControlId = _soxContext.RcmControlId.Where(x => x.ControlId.Equals(soxTracker.ControlId)).FirstOrDefault();
                                                    if (checkRcmControlId != null)
                                                    {
                                                        int KeyReportId = int.Parse(_config.GetSection("KeyReportApp").GetSection("Consolidated_Fields").GetSection("KeyControlId").Value);
                                                        var controlIdRef = KeyReportConsolidated.Field<AppItemField>(KeyReportId);
                                                        List<int> listControlId = new List<int>();
                                                        listControlId.Add(checkRcmControlId.PodioItemId);
                                                        controlIdRef.ItemIds = listControlId;
                                                    }
                                                    else
                                                    {
                                                        int KeyReportId = int.Parse(_config.GetSection("KeyReportApp").GetSection("Consolidated_Fields").GetSection("KeyControlId").Value);
                                                        var controlIdRef = KeyReportConsolidated.Field<AppItemField>(KeyReportId);
                                                        List<int> listControlId = new List<int>();
                                                    }
                                                }


                                                int NameofIUC = int.Parse(_config.GetSection("KeyReportApp").GetSection("Consolidated_Fields").GetSection("KeyReport_IUC").Value);
                                                var NameofIUCref = KeyReportConsolidated.Field<AppItemField>(NameofIUC);
                                                List<int> listNameofIUC = new List<int>();
                                                listNameofIUC.Add(_KeyReportID);
                                                NameofIUCref.ItemIds = listNameofIUC;

                                                if (soxTracker.Process != string.Empty && soxTracker.Process != null)
                                                {

                                                    var checkRcmProcess = _soxContext.RcmProcess.Where(x => x.Process.Equals(soxTracker.Process)).FirstOrDefault();
                                                    if (checkRcmProcess != null)
                                                    {
                                                        int Process = int.Parse(_config.GetSection("KeyReportApp").GetSection("Consolidated_Fields").GetSection("SourceProcess").Value);
                                                        var processRef = KeyReportConsolidated.Field<AppItemField>(Process);
                                                        List<int> listProcess = new List<int>();
                                                        listProcess.Add(checkRcmProcess.PodioItemId);
                                                        processRef.ItemIds = listProcess;
                                                    }
                                                }



                                                //search and update
                                                var OrigFormatDatabaseItem = _soxContext.KeyReportOrigFormat
                                                                     .Where(x =>
                                                                     x.ClientName.Equals(soxTracker.ClientName) &&
                                                                     x.ClientCode.Equals(soxTracker.ClientCode) &&
                                                                     x.NameOfKeyReport.Equals(soxTracker.KeyReportName) &&
                                                                     x.KeyControlId.Equals(soxTracker.ControlId)).FirstOrDefault();



                                                KeyReportOrigFormat _KeyReportOrigFormat = new KeyReportOrigFormat();

                                                _KeyReportOrigFormat.ClientName = soxTracker.ClientName;
                                                _KeyReportOrigFormat.ClientCode = soxTracker.ClientCode;
                                                _KeyReportOrigFormat.ClientItemId = soxTracker.ClientItemId;
                                                _KeyReportOrigFormat.KeyControlId = soxTracker.ControlId;
                                                _KeyReportOrigFormat.ControlActivity = "";
                                                _KeyReportOrigFormat.SKey = "";
                                                _KeyReportOrigFormat.NameOfKeyReport = soxTracker.KeyReportName;
                                                _KeyReportOrigFormat.KeyReport = soxTracker.KeyReport;
                                                _KeyReportOrigFormat.IUCType = "";
                                                _KeyReportOrigFormat.Notes = "Active";
                                                _KeyReportOrigFormat.PodioItemId = OrigFormatDatabaseItem.PodioItemId;
                                                _KeyReportOrigFormat.PodioUniqueId = soxTracker.PodioUniqueId;
                                                _KeyReportOrigFormat.PodioRevision = soxTracker.PodioRevision;
                                                _KeyReportOrigFormat.PodioLink = soxTracker.PodioLink;
                                                _KeyReportOrigFormat.CreatedBy = soxTracker.CreatedBy;


                                                using (var context = _soxContext.Database.BeginTransaction())
                                                {
                                                    _KeyReportOrigFormat.Id = OrigFormatDatabaseItem.Id;
                                                    _soxContext.Entry(OrigFormatDatabaseItem).CurrentValues.SetValues(_KeyReportOrigFormat);
                                                    await _soxContext.SaveChangesAsync();
                                                    context.Commit();

                                                }

                                                KeyReportConsolidated.ItemId = OrigFormatDatabaseItem.PodioItemId;
                                                var ConsolidatedID = await podio.ItemService.UpdateItem(KeyReportConsolidated); //uncomment after dev
                                            }
                                        }


                                        //========================================================================================AllUIC App =================================================================================
                                        PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("AllIUCId").Value;
                                        PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("AllIUCToken").Value;

                                        if (PodioAppKey.AppId != string.Empty && PodioAppKey.AppToken != string.Empty)
                                        {
                                            ServicePointManager.Expect100Continue = true;
                                            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                                            await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                                            if (podio.IsAuthenticated() && soxTracker != null)
                                            {
                                                Item KeyReportAllIUC = new Item();

                                                if (soxTracker.FY != string.Empty && soxTracker.FY != null)
                                                {
                                                    int FyField = int.Parse(_config.GetSection("KeyReportApp").GetSection("AllIUC_Fields").GetSection("FY").Value);
                                                    var FyFieldValue = KeyReportAllIUC.Field<TextItemField>(FyField);
                                                    FyFieldValue.Value = soxTracker.FY;
                                                }
                                                else
                                                {
                                                    int FyField = int.Parse(_config.GetSection("KeyReportApp").GetSection("AllIUC_Fields").GetSection("FY").Value);
                                                    var FyFieldValue = KeyReportAllIUC.Field<TextItemField>(FyField);
                                                    FyFieldValue.Value = " ";
                                                }

                                                int Num2Screenshot = int.Parse(_config.GetSection("KeyReportApp").GetSection("AllIUC_Fields").GetSection("Screenshot").Value);
                                                ImageItemField imageField = KeyReportAllIUC.Field<ImageItemField>(Num2Screenshot);
                                                imageField.FileIds = new List<int>();

                                                if (soxTracker.ClientName != string.Empty && soxTracker.ClientName != null)
                                                {
                                                    var checkClientItemId = _soxContext.ClientSs.Where(x => x.ClientName.Equals(soxTracker.ClientName)).Select(x => x.ItemId).FirstOrDefault();
                                                    //Console.WriteLine(checkClientItemId);
                                                    if (checkClientItemId != null)
                                                    {

                                                        int ClientNameField = int.Parse(_config.GetSection("KeyReportApp").GetSection("AllIUC_Fields").GetSection("ClientName").Value);
                                                        var ClientNameValue = KeyReportAllIUC.Field<AppItemField>(ClientNameField);
                                                        List<int> listClient = new List<int>();
                                                        listClient.Add(checkClientItemId.Value);
                                                        ClientNameValue.ItemIds = listClient;
                                                    }
                                                }
                                                else
                                                {
                                                    int ClientNameField = int.Parse(_config.GetSection("KeyReportApp").GetSection("AllIUC_Fields").GetSection("ClientName").Value);
                                                    var ClientNameValue = KeyReportAllIUC.Field<AppItemField>(ClientNameField);
                                                    List<int> listId = new List<int>();
                                                    ClientNameValue.ItemIds = listId;

                                                }

                                                if (soxTracker.ControlId != string.Empty && soxTracker.ControlId != null)
                                                {
                                                    var checkRcmControlId = _soxContext.RcmControlId.Where(x => x.ControlId.Equals(soxTracker.ControlId)).FirstOrDefault();
                                                    if (checkRcmControlId != null)
                                                    {
                                                        int KeyReportId = int.Parse(_config.GetSection("KeyReportApp").GetSection("AllIUC_Fields").GetSection("KeyControl").Value);
                                                        var controlIdRef = KeyReportAllIUC.Field<AppItemField>(KeyReportId);
                                                        List<int> listControlId = new List<int>();
                                                        listControlId.Add(checkRcmControlId.PodioItemId);
                                                        controlIdRef.ItemIds = listControlId;
                                                    }
                                                    else
                                                    {
                                                        int KeyReportId = int.Parse(_config.GetSection("KeyReportApp").GetSection("AllIUC_Fields").GetSection("KeyControl").Value);
                                                        var controlIdRef = KeyReportAllIUC.Field<AppItemField>(KeyReportId);
                                                        List<int> listControlId = new List<int>();
                                                    }
                                                }


                                                int NameofIUC = int.Parse(_config.GetSection("KeyReportApp").GetSection("AllIUC_Fields").GetSection("KeyReport/IUC").Value);
                                                var NameofIUCref = KeyReportAllIUC.Field<AppItemField>(NameofIUC);
                                                List<int> listNameofIUC = new List<int>();
                                                listNameofIUC.Add(_KeyReportID);
                                                NameofIUCref.ItemIds = listNameofIUC;

                                                if (soxTracker.KeyReport != string.Empty && soxTracker.KeyReport != null)
                                                {

                                                    var KeyreportYesNo = _soxContext.KeyReportUniqueKeyReport.Where(x => x.UniqueKeyReport.Equals(soxTracker.KeyReport)).FirstOrDefault();
                                                    if (KeyreportYesNo != null)
                                                    {
                                                        int UniqueKeyReport = int.Parse(_config.GetSection("KeyReportApp").GetSection("AllIUC_Fields").GetSection("KeyReport").Value);
                                                        var UniqueKeyReportref = KeyReportAllIUC.Field<AppItemField>(UniqueKeyReport);
                                                        List<int> ListUniqueKeyReport = new List<int>();
                                                        ListUniqueKeyReport.Add(KeyreportYesNo.PodioItemId);
                                                        UniqueKeyReportref.ItemIds = ListUniqueKeyReport;
                                                    }
                                                    else
                                                    {
                                                        int UniqueKeyReport = int.Parse(_config.GetSection("KeyReportApp").GetSection("AllIUC_Fields").GetSection("KeyReport").Value);
                                                        var UniqueKeyReportref = KeyReportAllIUC.Field<AppItemField>(UniqueKeyReport);
                                                        List<int> ListUniqueKeyReport = new List<int>();
                                                    }
                                                }


                                                if (soxTracker.Process != string.Empty && soxTracker.Process != null)
                                                {

                                                    var checkRcmProcess = _soxContext.RcmProcess.Where(x => x.Process.Equals(soxTracker.Process)).FirstOrDefault();
                                                    if (checkRcmProcess != null)
                                                    {
                                                        int Process = int.Parse(_config.GetSection("KeyReportApp").GetSection("AllIUC_Fields").GetSection("SourceProcess").Value);
                                                        var processRef = KeyReportAllIUC.Field<AppItemField>(Process);
                                                        List<int> listProcess = new List<int>();
                                                        listProcess.Add(checkRcmProcess.PodioItemId);
                                                        processRef.ItemIds = listProcess;
                                                    }
                                                }


                                                //search and update
                                                var AllIUCDatabaseItem = _soxContext.KeyReportAllIUC
                                                                     .Where(x =>
                                                                     x.ClientName.Equals(soxTracker.ClientName) &&
                                                                     x.ClientCode.Equals(soxTracker.ClientCode) &&
                                                                     x.NameOfIUC.Equals(soxTracker.KeyReportName) &&
                                                                     x.KeyControlId.Equals(soxTracker.ControlId)).FirstOrDefault();


                                                KeyReportAllIUC _KeyReportAllIUC = new KeyReportAllIUC();

                                                _KeyReportAllIUC.ClientName = soxTracker.ClientName;
                                                _KeyReportAllIUC.ClientCode = soxTracker.ClientCode;
                                                _KeyReportAllIUC.ClientItemId = soxTracker.ClientItemId;
                                                _KeyReportAllIUC.Num = 0;
                                                _KeyReportAllIUC.KeyControlId = soxTracker.ControlId;
                                                _KeyReportAllIUC.ControlActivity = "";
                                                _KeyReportAllIUC.SKey = soxTracker.KeyReport;
                                                _KeyReportAllIUC.NameOfIUC = soxTracker.KeyReportName;
                                                _KeyReportAllIUC.KeyReport = soxTracker.KeyReport;
                                                _KeyReportAllIUC.UniqueKeyReport = soxTracker.KeyReport; // should be yes referrence on podio
                                                _KeyReportAllIUC.PodioItemId = AllIUCDatabaseItem.PodioItemId;
                                                _KeyReportAllIUC.Process = soxTracker.Process;
                                                _KeyReportAllIUC.PodioUniqueId = soxTracker.PodioUniqueId;
                                                _KeyReportAllIUC.PodioRevision = soxTracker.PodioRevision;
                                                _KeyReportAllIUC.PodioLink = soxTracker.PodioLink;
                                                _KeyReportAllIUC.CreatedBy = soxTracker.CreatedBy;


                                                using (var context = _soxContext.Database.BeginTransaction())
                                                {
                                                    _KeyReportAllIUC.Id = AllIUCDatabaseItem.Id;
                                                    _soxContext.Entry(AllIUCDatabaseItem).CurrentValues.SetValues(_KeyReportAllIUC);
                                                    await _soxContext.SaveChangesAsync();
                                                    context.Commit();

                                                }
                                                KeyReportAllIUC.ItemId = AllIUCDatabaseItem.PodioItemId;
                                                var KeyReportAllIUCID = await podio.ItemService.UpdateItem(KeyReportAllIUC);

                                            }
                                        }

                                        //========================================================================================Test Status Tracker App Update=================================================================================
                                        PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("TestStatusTrackerId").Value;
                                        PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("TestStatusTrackerToken").Value;

                                        if (PodioAppKey.AppId != string.Empty && PodioAppKey.AppToken != string.Empty)
                                        {
                                            ServicePointManager.Expect100Continue = true;
                                            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                                            await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                                            if (podio.IsAuthenticated() && soxTracker != null)
                                            {
                                                Item KeyReportTestStatus = new Item();

                                                if (soxTracker.FY != string.Empty && soxTracker.FY != null)
                                                {
                                                    int FyField = int.Parse(_config.GetSection("KeyReportApp").GetSection("TestStatusTracker_Fields").GetSection("FY").Value);
                                                    var FyFieldValue = KeyReportTestStatus.Field<TextItemField>(FyField);
                                                    FyFieldValue.Value = soxTracker.FY;
                                                }
                                                else
                                                {
                                                    int FyField = int.Parse(_config.GetSection("KeyReportApp").GetSection("TestStatusTracker_Fields").GetSection("FY").Value);
                                                    var FyFieldValue = KeyReportTestStatus.Field<TextItemField>(FyField);
                                                    FyFieldValue.Value = " ";
                                                }

                                                int Num2Screenshot = int.Parse(_config.GetSection("KeyReportApp").GetSection("TestStatusTracker_Fields").GetSection("1Screenshot").Value);
                                                ImageItemField imageField = KeyReportTestStatus.Field<ImageItemField>(Num2Screenshot);
                                                imageField.FileIds = new List<int>();

                                                if (soxTracker.ClientName != string.Empty && soxTracker.ClientName != null)
                                                {
                                                    var checkClientItemId = _soxContext.ClientSs.Where(x => x.ClientName.Equals(soxTracker.ClientName)).Select(x => x.ItemId).FirstOrDefault();
                                                    //Console.WriteLine(checkClientItemId);
                                                    if (checkClientItemId != null)
                                                    {

                                                        int ClientNameField = int.Parse(_config.GetSection("KeyReportApp").GetSection("TestStatusTracker_Fields").GetSection("ClientName").Value);
                                                        var ClientNameValue = KeyReportTestStatus.Field<AppItemField>(ClientNameField);
                                                        List<int> listClient = new List<int>();
                                                        listClient.Add(checkClientItemId.Value);
                                                        ClientNameValue.ItemIds = listClient;
                                                    }
                                                }
                                                else
                                                {
                                                    int ClientNameField = int.Parse(_config.GetSection("KeyReportApp").GetSection("TestStatusTracker_Fields").GetSection("ClientName").Value);
                                                    var ClientNameValue = KeyReportTestStatus.Field<AppItemField>(ClientNameField);
                                                    List<int> listId = new List<int>();
                                                    ClientNameValue.ItemIds = listId;

                                                }

                                                if (soxTracker.ControlId != string.Empty && soxTracker.ControlId != null)
                                                {
                                                    var checkRcmControlId = _soxContext.RcmControlId.Where(x => x.ControlId.Equals(soxTracker.ControlId)).FirstOrDefault();
                                                    if (checkRcmControlId != null)
                                                    {
                                                        int KeyReportId = int.Parse(_config.GetSection("KeyReportApp").GetSection("TestStatusTracker_Fields").GetSection("KeyControlID").Value);
                                                        var controlIdRef = KeyReportTestStatus.Field<AppItemField>(KeyReportId);
                                                        List<int> listControlId = new List<int>();
                                                        listControlId.Add(checkRcmControlId.PodioItemId);
                                                        controlIdRef.ItemIds = listControlId;
                                                    }
                                                    else
                                                    {
                                                        int KeyReportId = int.Parse(_config.GetSection("KeyReportApp").GetSection("TestStatusTracker_Fields").GetSection("KeyControlID").Value);
                                                        var controlIdRef = KeyReportTestStatus.Field<AppItemField>(KeyReportId);
                                                        List<int> listControlId = new List<int>();
                                                    }
                                                }


                                                int NameofIUC = int.Parse(_config.GetSection("KeyReportApp").GetSection("TestStatusTracker_Fields").GetSection("KeyReport/IUC").Value);
                                                var NameofIUCref = KeyReportTestStatus.Field<AppItemField>(NameofIUC);
                                                List<int> listNameofIUC = new List<int>();
                                                listNameofIUC.Add(_KeyReportID);
                                                NameofIUCref.ItemIds = listNameofIUC;

                                                if (soxTracker.Process != string.Empty && soxTracker.Process != null)
                                                {

                                                    var checkRcmProcess = _soxContext.RcmProcess.Where(x => x.Process.Equals(soxTracker.Process)).FirstOrDefault();
                                                    if (checkRcmProcess != null)
                                                    {
                                                        int Process = int.Parse(_config.GetSection("KeyReportApp").GetSection("TestStatusTracker_Fields").GetSection("SourceProcess").Value);
                                                        var processRef = KeyReportTestStatus.Field<AppItemField>(Process);
                                                        List<int> listProcess = new List<int>();
                                                        listProcess.Add(checkRcmProcess.PodioItemId);
                                                        processRef.ItemIds = listProcess;
                                                    }
                                                }


                                                var TestStatusTracker = _soxContext.KeyReportTestStatusTracker
                                                                    .Where(x =>
                                                                    x.ClientName.Equals(soxTracker.ClientName) &&
                                                                    x.ClientCode.Equals(soxTracker.ClientCode) &&
                                                                    x.NameOfIUC.Equals(soxTracker.KeyReportName) &&
                                                                    x.KeyControlId.Equals(soxTracker.ControlId)).FirstOrDefault();


                                                //Search Update
                                                KeyReportTestStatusTracker _KeyReportTestStatusTracker = new KeyReportTestStatusTracker();

                                                _KeyReportTestStatusTracker.ClientName = soxTracker.ClientName;
                                                _KeyReportTestStatusTracker.ClientCode = soxTracker.ClientCode;
                                                _KeyReportTestStatusTracker.ClientItemId = soxTracker.ClientItemId;
                                                _KeyReportTestStatusTracker.NameOfIUC = soxTracker.KeyReportName;
                                                _KeyReportTestStatusTracker.Process = soxTracker.Process;
                                                _KeyReportTestStatusTracker.KeyControlId = soxTracker.ControlId;
                                                _KeyReportTestStatusTracker.KeyReport = soxTracker.KeyReport;
                                                _KeyReportTestStatusTracker.UniqueKeyReport = soxTracker.KeyReport;
                                                _KeyReportTestStatusTracker.PBCStatus = soxTracker.Status;
                                                _KeyReportTestStatusTracker.TestingStatus = soxTracker.WTTestingStatus;
                                                _KeyReportTestStatusTracker.PodioItemId = TestStatusTracker.PodioItemId;
                                                _KeyReportTestStatusTracker.PodioUniqueId = soxTracker.PodioUniqueId; // should be yes referrence on podio
                                                _KeyReportTestStatusTracker.PodioRevision = soxTracker.PodioRevision;
                                                _KeyReportTestStatusTracker.PodioLink = soxTracker.PodioLink;
                                                _KeyReportTestStatusTracker.CreatedBy = soxTracker.CreatedBy;

                                                using (var context = _soxContext.Database.BeginTransaction())
                                                {
                                                    _KeyReportTestStatusTracker.Id = TestStatusTracker.Id;
                                                    _soxContext.Entry(TestStatusTracker).CurrentValues.SetValues(_KeyReportTestStatusTracker);
                                                    await _soxContext.SaveChangesAsync();
                                                    context.Commit();

                                                }

                                                KeyReportTestStatus.ItemId = TestStatusTracker.PodioItemId;
                                                var TestTstatusID = await podio.ItemService.UpdateItem(KeyReportTestStatus);// uncomment after dev

                                            }
                                        }

                                        //========================================================================================Exception App =================================================================================
                                        PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("ExceptionId").Value;
                                        PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("ExceptionToken").Value;

                                        if (PodioAppKey.AppId != string.Empty && PodioAppKey.AppToken != string.Empty)
                                        {
                                            ServicePointManager.Expect100Continue = true;
                                            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                                            await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                                            if (podio.IsAuthenticated() && soxTracker != null)
                                            {
                                                Item Exception = new Item();

                                                if (soxTracker.FY != string.Empty && soxTracker.FY != null)
                                                {
                                                    int FyField = int.Parse(_config.GetSection("KeyReportApp").GetSection("Exceptions_Fields").GetSection("FY").Value);
                                                    var FyFieldValue = Exception.Field<TextItemField>(FyField);
                                                    FyFieldValue.Value = soxTracker.FY;
                                                }
                                                else
                                                {
                                                    int FyField = int.Parse(_config.GetSection("KeyReportApp").GetSection("Exceptions_Fields").GetSection("FY").Value);
                                                    var FyFieldValue = Exception.Field<TextItemField>(FyField);
                                                    FyFieldValue.Value = " ";
                                                }

                                                int Num2Screenshot = int.Parse(_config.GetSection("KeyReportApp").GetSection("Exceptions_Fields").GetSection("1Screenshot").Value);
                                                ImageItemField imageField = Exception.Field<ImageItemField>(Num2Screenshot);
                                                imageField.FileIds = new List<int>();

                                                if (soxTracker.ClientName != string.Empty && soxTracker.ClientName != null)
                                                {
                                                    var checkClientItemId = _soxContext.ClientSs.Where(x => x.ClientName.Equals(soxTracker.ClientName)).Select(x => x.ItemId).FirstOrDefault();
                                                    //Console.WriteLine(checkClientItemId);
                                                    if (checkClientItemId != null)
                                                    {

                                                        int ClientNameField = int.Parse(_config.GetSection("KeyReportApp").GetSection("Exceptions_Fields").GetSection("ClientName").Value);
                                                        var ClientNameValue = Exception.Field<AppItemField>(ClientNameField);
                                                        List<int> listClient = new List<int>();
                                                        listClient.Add(checkClientItemId.Value);
                                                        ClientNameValue.ItemIds = listClient;
                                                    }
                                                }
                                                else
                                                {
                                                    int ClientNameField = int.Parse(_config.GetSection("KeyReportApp").GetSection("Exceptions_Fields").GetSection("ClientName").Value);
                                                    var ClientNameValue = Exception.Field<AppItemField>(ClientNameField);
                                                    List<int> listId = new List<int>();
                                                    ClientNameValue.ItemIds = listId;

                                                }

                                                if (soxTracker.ControlId != string.Empty && soxTracker.ControlId != null)
                                                {
                                                    var checkRcmControlId = _soxContext.RcmControlId.Where(x => x.ControlId.Equals(soxTracker.ControlId)).FirstOrDefault();
                                                    if (checkRcmControlId != null)
                                                    {
                                                        int KeyReportId = int.Parse(_config.GetSection("KeyReportApp").GetSection("Exceptions_Fields").GetSection("KeyControlID").Value);
                                                        var controlIdRef = Exception.Field<AppItemField>(KeyReportId);
                                                        List<int> listControlId = new List<int>();
                                                        listControlId.Add(checkRcmControlId.PodioItemId);
                                                        controlIdRef.ItemIds = listControlId;
                                                    }
                                                    else
                                                    {
                                                        int KeyReportId = int.Parse(_config.GetSection("KeyReportApp").GetSection("Exceptions_Fields").GetSection("KeyControlID").Value);
                                                        var controlIdRef = Exception.Field<AppItemField>(KeyReportId);
                                                        List<int> listControlId = new List<int>();
                                                    }
                                                }


                                                int NameofIUC = int.Parse(_config.GetSection("KeyReportApp").GetSection("Exceptions_Fields").GetSection("KeyReport/IUC").Value);
                                                var NameofIUCref = Exception.Field<AppItemField>(NameofIUC);
                                                List<int> listNameofIUC = new List<int>();
                                                listNameofIUC.Add(_KeyReportID);
                                                NameofIUCref.ItemIds = listNameofIUC;


                                                var ExceptionDatabaseItem = _soxContext.KeyReportExcepcion
                                                                 .Where(x =>
                                                                 x.ClientName.Equals(soxTracker.ClientName) &&
                                                                 x.ClientCode.Equals(soxTracker.ClientCode) &&
                                                                 x.NameOfIUC.Equals(soxTracker.KeyReportName) &&
                                                                 x.KeyControlId.Equals(soxTracker.ControlId)).FirstOrDefault();

                                                //Search Update
                                                KeyReportExcepcion _KeyReportExcepcion = new KeyReportExcepcion();

                                                _KeyReportExcepcion.ClientName = soxTracker.ClientName;
                                                _KeyReportExcepcion.ClientCode = soxTracker.ClientCode;
                                                _KeyReportExcepcion.ClientItemId = soxTracker.ClientItemId;
                                                _KeyReportExcepcion.NameOfIUC = soxTracker.KeyReportName;
                                                _KeyReportExcepcion.KeyControlId = soxTracker.ControlId;
                                                _KeyReportExcepcion.PodioItemId = ExceptionDatabaseItem.PodioItemId;
                                                _KeyReportExcepcion.PodioUniqueId = soxTracker.PodioUniqueId; // should be yes referrence on podio
                                                _KeyReportExcepcion.PodioRevision = soxTracker.PodioRevision;
                                                _KeyReportExcepcion.PodioLink = soxTracker.PodioLink;
                                                _KeyReportExcepcion.CreatedBy = soxTracker.CreatedBy;

                                                using (var context = _soxContext.Database.BeginTransaction())
                                                {
                                                    _KeyReportExcepcion.Id = ExceptionDatabaseItem.Id;
                                                    _soxContext.Entry(ExceptionDatabaseItem).CurrentValues.SetValues(_KeyReportExcepcion);
                                                    await _soxContext.SaveChangesAsync();
                                                    context.Commit();
                                                }
                                                Exception.ItemId = ExceptionDatabaseItem.PodioItemId;
                                                var ExceptionID = await podio.ItemService.UpdateItem(Exception); //comment after d
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        //Console.WriteLine(ex.ToString());
                                        //ErrorLog.Write(ex);
                                        FileLog.Write(ex.ToString(), "ErrorUpdatePodioSoxTrackerAsync");
                                        if (_environment.IsDevelopment())
                                        {
                                            //Console.WriteLine(BadRequest(ex.ToString()));
                                            return BadRequest(ex.ToString());
                                        }
                                        else
                                        {
                                            BadRequest();
                                            //Console.WriteLine(BadRequest());
                                            return BadRequest();
                                        }

                                    }
                                }

                            }
                            else
                            {
                                PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("AllIUCId").Value;
                                PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("AllIUCToken").Value;

                                if (PodioAppKey.AppId != string.Empty && PodioAppKey.AppToken != string.Empty)
                                {
                                    ServicePointManager.Expect100Continue = true;
                                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                                    await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                                    if (podio.IsAuthenticated() && soxTracker != null)
                                    {
                                        Item KeyReportAllIUC = new Item();
                                        if (soxTracker.KeyReport != string.Empty && soxTracker.KeyReport != null)
                                        {

                                            var KeyreportYesNo = _soxContext.KeyReportUniqueKeyReport.Where(x => x.UniqueKeyReport.Equals(soxTracker.KeyReport)).FirstOrDefault();
                                            if (KeyreportYesNo != null)
                                            {
                                                int UniqueKeyReport = int.Parse(_config.GetSection("KeyReportApp").GetSection("AllIUC_Fields").GetSection("KeyReport").Value);
                                                var UniqueKeyReportref = KeyReportAllIUC.Field<AppItemField>(UniqueKeyReport);
                                                List<int> ListUniqueKeyReport = new List<int>();
                                                ListUniqueKeyReport.Add(KeyreportYesNo.PodioItemId);
                                                UniqueKeyReportref.ItemIds = ListUniqueKeyReport;
                                            }
                                            else
                                            {
                                                int UniqueKeyReport = int.Parse(_config.GetSection("KeyReportApp").GetSection("AllIUC_Fields").GetSection("KeyReport").Value);
                                                var UniqueKeyReportref = KeyReportAllIUC.Field<AppItemField>(UniqueKeyReport);
                                                List<int> ListUniqueKeyReport = new List<int>();
                                            }
                                        }

                                        var AllIUCDatabaseItem = _soxContext.KeyReportAllIUC
                                                   .Where(x =>
                                                   x.ClientName.Equals(soxTracker.ClientName) &&
                                                   x.ClientCode.Equals(soxTracker.ClientCode) &&
                                                   x.NameOfIUC.Equals(soxTracker.KeyReportName) &&
                                                   x.KeyControlId.Equals(soxTracker.ControlId)).FirstOrDefault();


                                        KeyReportAllIUC _KeyReportAllIUC = new KeyReportAllIUC();
                                        _KeyReportAllIUC.UniqueKeyReport = soxTracker.KeyReport;
                                        using (var context = _soxContext.Database.BeginTransaction())

                                        {
                                            _soxContext.Add(_KeyReportAllIUC);
                                            await _soxContext.SaveChangesAsync();
                                            context.Commit();
                                        }

                                        KeyReportAllIUC.ItemId = AllIUCDatabaseItem.PodioItemId;
                                        var KeyReportAllIUCID = await podio.ItemService.UpdateItem(KeyReportAllIUC);

                                    }
                                }

                            }
                        }

                        status = true;
                    }
                }

            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorUpdatePodioSoxTrackerAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "UpdatePodioSoxTrackerAsync");
                if (_environment.IsDevelopment())
                {
                    return BadRequest(ex.ToString());
                }
                else
                {
                    return BadRequest();
                }

            }

            //for testing
            status = true;
            if (status)
            {
                return Ok(soxTracker);
            }
            else
            {
                return NoContent();
            }

        }


        private void PopulateDataConsolidated(Item KeyReportItem, SoxTracker soxTracker, bool update)
        {

            //FY
            if (soxTracker.FY != string.Empty && soxTracker.FY != null)
            {
                int FyField = int.Parse(_config.GetSection("KeyReportApp").GetSection("Exceptions_Fields").GetSection("FY").Value);
                var FyFieldValue = KeyReportItem.Field<TextItemField>(FyField);
                FyFieldValue.Value = soxTracker.FY;
            }
            else
            {
                int FyField = int.Parse(_config.GetSection("KeyReportApp").GetSection("Exceptions_Fields").GetSection("FY").Value);
                var FyFieldValue = KeyReportItem.Field<TextItemField>(FyField);
                FyFieldValue.Value = " ";
            }

            // if (soxTracker.KeyReport != string.Empty)
            // {
            //     var KeyReportChoice = soxTrackerItem.Field<TextItemField>(q15EField);
            //     KeyReportChoice.Value = soxTracker.KeyReport;
            //     Console.WriteLine(soxTracker.KeyReport);
            // }


            //Client Name
            if (soxTracker.ClientName != string.Empty && soxTracker.ClientName != null)
            {
                var checkClientItemId = _soxContext.ClientSs.Where(x => x.ClientName.Equals(soxTracker.ClientName)).Select(x => x.ItemId).FirstOrDefault();
                //Console.WriteLine(checkClientItemId);
                if (checkClientItemId != null)
                {

                    int ClientNameField = int.Parse(_config.GetSection("KeyReportApp").GetSection("Exceptions_Fields").GetSection("ClientName").Value);
                    var ClientNameValue = KeyReportItem.Field<AppItemField>(ClientNameField);
                    List<int> listClient = new List<int>();
                    listClient.Add(checkClientItemId.Value);
                    ClientNameValue.ItemIds = listClient;
                }
            }
            else
            {
                int ClientNameField = int.Parse(_config.GetSection("KeyReportApp").GetSection("Exceptions_Fields").GetSection("ClientName").Value);
                var ClientNameValue = KeyReportItem.Field<AppItemField>(ClientNameField);
                List<int> listId = new List<int>();
                ClientNameValue.ItemIds = listId;
                /*
                 *  int ClientId = (int)soxTracker.ClientItemId;
                    ClientNameValue.ItemIds = new List<int>
                {

                };
                    var pbcOwnerIdRef = soxTrackerItem.Field<AppItemField>(q7Field);
                    List<int> listId = new List<int>();
                    pbcOwnerIdRef.ItemIds = listId;
                 */

            }

        }

        [AllowAnonymous]
        [HttpPost("save")]
        public async Task<IActionResult> SaveToDatabase([FromBody] SoxTracker soxTracker)
        {
            bool status = false;
            try
            {
                using (var context = _soxContext.Database.BeginTransaction())
                {

                    //Console.WriteLine(soxTracker);

                    /*
                    PodioApiKey PodioKey = new PodioApiKey();
                    PodioAppKey PodioAppKey = new PodioAppKey();
                    PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                    PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                    PodioAppKey.AppId = _config.GetSection("SoxTrackerApp").GetSection("AppId").Value;
                    PodioAppKey.AppToken = _config.GetSection("SoxTrackerApp").GetSection("AppToken").Value;

                    var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                    ServicePointManager.Expect100Continue = true;
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                    await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                    if (podio.IsAuthenticated() && soxTracker != null)
                    {
                        var podioItem = podio.ItemService.GetItem(soxTracker.PodioItemId);
                        var rawResponse = Newtonsoft.Json.JsonConvert.SerializeObject(podioItem.Result);
                        FileLog.Write($"{rawResponse}", "SaveToDatabase");
                        soxTracker.PodioItemId = (int)podioItem.Result.ItemId;
                        soxTracker.PodioUniqueId = podioItem.Result.AppItemIdFormatted.ToString();
                        soxTracker.PodioRevision = podioItem.Result.CurrentRevision.Revision;
                        soxTracker.PodioLink = podioItem.Result.Link.ToString();
                        soxTracker.CreatedBy = podioItem.Result.CreatedBy.Name.ToString();
                        soxTracker.CreatedOn = DateTime.Parse(podioItem.Result.CreatedOn.ToString());
                    }
                    */
                    soxTracker.Status = "Active";
                    //bool isExists = false;  

                    //validate created date
                    if (DateTime.TryParse(soxTracker.CreatedOn.Value.DateTime.ToString(), out DateTime dtCreated) && soxTracker.CreatedOn.Value.DateTime.ToString() != "01/01/0001 00:00:00")
                    {
                        soxTracker.CreatedOn = dtCreated;
                    }
                    else
                    {
                        soxTracker.CreatedOn = DateTime.Now;
                    }

                    if (DateTime.TryParse(soxTracker.LastUpdate.DateTime.ToString(), out DateTime dtLastUpdate) && soxTracker.LastUpdate.DateTime.ToString() != "01/01/0001 00:00:00")
                    {
                        soxTracker.LastUpdate = dtLastUpdate;
                    }
                    else
                    {
                        soxTracker.LastUpdate = DateTime.Now;
                    }
                    //Get client code and client item id
                    var clientCheck = _soxContext.ClientSs.Where(x => x.ClientName.Equals(soxTracker.ClientName)).FirstOrDefault();
                    if (clientCheck != null)
                    {
                        soxTracker.ClientCode = clientCheck.ClientCode;
                        soxTracker.ClientItemId = clientCheck.ClientItemId;
                    }

                    //check for previous entry, we do upsert
                    var soxTrackerCheck = _soxContext.SoxTracker.FirstOrDefault(id => id.PodioItemId.Equals(soxTracker.PodioItemId));
                    if (soxTrackerCheck != null)
                    {
                        //sox tracker already exists and needs to update
                        soxTracker.Id = soxTrackerCheck.Id;
                        _soxContext.Entry(soxTrackerCheck).CurrentValues.SetValues(soxTracker);
                        await _soxContext.SaveChangesAsync();
                        context.Commit();
                    }
                    else
                    {
                        //sox tracker is new and needs to be added
                        _soxContext.Add(soxTracker);
                        await _soxContext.SaveChangesAsync();
                        context.Commit();
                    }
                }
                using (var context = _soxContext.Database.BeginTransaction())
                {
                    //check for previous entry, we do upsert
                    var rcmCheck = _soxContext.Rcm
                                                .Where(x => x.ClientName.Equals(soxTracker.ClientName)
                                                           && x.Process.Equals(soxTracker.Process)
                                                           && x.Subprocess.Equals(soxTracker.Subprocess)
                                                           && x.FY.Equals(soxTracker.FY) 
                                                           && x.ControlId.Equals(soxTracker.ControlId))
                                                .FirstOrDefault();
                    if (rcmCheck != null)
                    {
                        //sox tracker already exists and needs to update
                        rcmCheck.PbcList = soxTracker.PBC;
                        _soxContext.Update(rcmCheck);
                        await _soxContext.SaveChangesAsync();
                        context.Commit();
                    }
                    else
                    {
                        //sox tracker is new and needs to be added
                        _soxContext.Add(soxTracker);
                        await _soxContext.SaveChangesAsync();
                        context.Commit();
                    }
                }

                status = true;
            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSaveToDatabase");
                if (_environment.IsDevelopment())
                {
                    return BadRequest(ex.ToString());
                }
                else
                {
                    return BadRequest();
                }

            }
            if (status)
            {
                return Ok();
            }
            else
            {
                return NoContent();
            }

        }

        [AllowAnonymous]
        private void PopulatePodioFields(Item soxTrackerItem, SoxTracker soxTracker, bool update)
        {
            #region Podio Fields

            int q6Field = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q6").Value);
            int q7Field = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q7").Value);
            int q7OtherField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q7Other").Value);
            int q8Field = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q8").Value);
            int q9Field = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q9").Value);
            int q10Field = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q10").Value);
            int q11Field = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q11").Value);
            int q12AField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q12A").Value);
            int q12BField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q12B").Value);
            int q12CField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q12C").Value);
            int q12DField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q12D").Value);
            int q13AField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q13A").Value);
            int q13BField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q13B").Value);
            int q13CField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q13C").Value);
            int q13DField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q13D").Value);
            int q13EField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q13E").Value);
            int q13FField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q13F").Value);
            int q13GField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q13G").Value);
            int q13HField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q13H").Value);
            int q13IField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q13I").Value);
            int q13JField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q13J").Value);
            int q13KField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q13K").Value);
            int q13LField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q13L").Value);
            int q14AField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q14A").Value);
            int q14BField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q14B").Value);
            int q14CField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q14C").Value);
            int q14DField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q14D").Value);
            int q14EField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q14E").Value);

            int q15EField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q15").Value);
            int q16EField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q16").Value);
            int duration = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Duration").Value);

            if (soxTracker.PBC != string.Empty && soxTracker.PBC != null)
            {

                var text6 = soxTrackerItem.Field<TextItemField>(q6Field);
                text6.Value = soxTracker.PBC;
            }
            else //Remove value
            {
                if (update)
                {
                    var text6 = soxTrackerItem.Field<TextItemField>(q6Field);
                    text6.Value = " ";
                }
            }

            if (soxTracker.KeyReportName != string.Empty && soxTracker.KeyReportName != null)
            {

                var text16 = soxTrackerItem.Field<TextItemField>(q16EField);
                text16.Value = soxTracker.KeyReportName;
            }
            else //Remove value
            {
                if (update)
                {
                    var text16 = soxTrackerItem.Field<TextItemField>(q16EField);
                    text16.Value = " ";
                }
            }

            if (soxTracker.KeyReport != string.Empty && soxTracker.KeyReport != null)
            {

                var text15 = soxTrackerItem.Field<TextItemField>(q15EField);
                text15.Value = soxTracker.KeyReport;
            }
            else //Remove value
            {
                if (update)
                {
                    var text15 = soxTrackerItem.Field<TextItemField>(q15EField);
                    text15.Value = " ";
                }
            }

            if (soxTracker.PBCOwner != string.Empty && soxTracker.PBCOwner != null && soxTracker.PBCOwner != "Select Option")
            {
                var checkPbcOwner = _soxContext.RcmControlOwner.Where(x => x.ControlOwner.Equals(soxTracker.PBCOwner)).FirstOrDefault();
                if (checkPbcOwner != null)
                {

                    var pbcOwnerIdRef = soxTrackerItem.Field<AppItemField>(q7Field);
                    List<int> listPbcOwner = new List<int>();
                    listPbcOwner.Add(checkPbcOwner.PodioItemId);
                    pbcOwnerIdRef.ItemIds = listPbcOwner;
                }
            }
            else //Remove value
            {
                if (update)
                {
                    var pbcOwnerIdRef = soxTrackerItem.Field<AppItemField>(q7Field);
                    List<int> listId = new List<int>();
                    pbcOwnerIdRef.ItemIds = listId;
                }
            }

            if (soxTracker.PBCOwnerOther != string.Empty && soxTracker.PBCOwnerOther != null)
            {

                var text7a = soxTrackerItem.Field<TextItemField>(q7OtherField);
                text7a.Value = soxTracker.PBCOwnerOther;
            }
            else //Remove value
            {
                if (update)
                {
                    var text7a = soxTrackerItem.Field<TextItemField>(q7OtherField);
                    text7a.Value = " ";
                }
            }

            if (soxTracker.PopulationFileRequest != string.Empty && soxTracker.PopulationFileRequest != null)
            {

                var popFileReq = soxTrackerItem.Field<CategoryItemField>(q8Field);
                popFileReq.OptionText = soxTracker.PopulationFileRequest;
            }
            else //Remove value
            {
                if (update)
                {
                    var popFileReq = soxTrackerItem.Field<CategoryItemField>(q8Field);
                    List<int> listId = new List<int>();
                    popFileReq.OptionIds = listId;
                }

            }

            if (soxTracker.SampleSelection != string.Empty && soxTracker.SampleSelection != null)
            {

                var sampleSel = soxTrackerItem.Field<CategoryItemField>(q9Field);
                sampleSel.OptionText = soxTracker.SampleSelection;
            }
            else //Remove value
            {
                if (update)
                {
                    var sampleSel = soxTrackerItem.Field<CategoryItemField>(q9Field);
                    List<int> listId = new List<int>();
                    sampleSel.OptionIds = listId;
                }

            }

            if (soxTracker.ExternalAuditorSample != string.Empty && soxTracker.ExternalAuditorSample != null)
            {

                var extAuditorR3 = soxTrackerItem.Field<CategoryItemField>(q10Field);
                extAuditorR3.OptionText = soxTracker.ExternalAuditorSample;
            }
            else //Remove value
            {
                if (update)
                {
                    var extAuditorR3 = soxTrackerItem.Field<CategoryItemField>(q10Field);
                    List<int> listId = new List<int>();
                    extAuditorR3.OptionIds = listId;
                }

            }

            if (soxTracker.R3Sample != string.Empty && soxTracker.R3Sample != null && soxTracker.R3Sample != "Select Option")
            {

                var r3Txt = soxTrackerItem.Field<TextItemField>(q11Field);
                r3Txt.Value = soxTracker.PBC;
            }
            else //Remove value
            {
                if (update)
                {
                    var r3Txt = soxTrackerItem.Field<TextItemField>(q11Field);
                    r3Txt.Value = " ";
                }

            }

            if (soxTracker.WTPBC != string.Empty && soxTracker.WTPBC != null && soxTracker.WTPBC != "Select Option")
            {
                List<int> listWtPBC = new List<int>();

                var checWTPPBC = _soxContext.SoxTrackerAppRelationship.Where(x => x.Title.Equals(soxTracker.WTPBC)).FirstOrDefault();
                if (checWTPPBC != null)
                {
                    var wtPBC = soxTrackerItem.Field<AppItemField>(q12AField);
                    listWtPBC.Add(checWTPPBC.PodioItemId);
                    wtPBC.ItemIds = listWtPBC;
                }
            }
            else //Remove value
            {
                if (update)
                {
                    var wtPBC = soxTrackerItem.Field<AppItemField>(q12AField);
                    List<int> listId = new List<int>();
                    wtPBC.ItemIds = listId;
                }
            }

            if (soxTracker.R1PBC != string.Empty && soxTracker.R1PBC != null && soxTracker.R1PBC != "Select Option")
            {


                var checkItem = _soxContext.SoxTrackerAppRelationship.Where(x => x.Title.Equals(soxTracker.R1PBC)).FirstOrDefault();
                if (checkItem != null)
                {
                    List<int> listId = new List<int>();
                    var appItem = soxTrackerItem.Field<AppItemField>(q12BField);
                    listId.Add(checkItem.PodioItemId);
                    appItem.ItemIds = listId;
                }
            }
            else //Remove value
            {
                if (update)
                {
                    var appItem = soxTrackerItem.Field<AppItemField>(q12BField);
                    List<int> listId = new List<int>();
                    appItem.ItemIds = listId;
                }
            }

            if (soxTracker.R2PBC != string.Empty && soxTracker.R2PBC != null && soxTracker.R2PBC != "Select Option")
            {


                var checkItem = _soxContext.SoxTrackerAppRelationship.Where(x => x.Title.Equals(soxTracker.R2PBC)).FirstOrDefault();
                if (checkItem != null)
                {
                    List<int> listId = new List<int>();
                    var appItem = soxTrackerItem.Field<AppItemField>(q12CField);
                    listId.Add(checkItem.PodioItemId);
                    appItem.ItemIds = listId;
                }
            }
            else //Remove value
            {
                if (update)
                {
                    var appItem = soxTrackerItem.Field<AppItemField>(q12CField);
                    List<int> listId = new List<int>();
                    appItem.ItemIds = listId;
                }
            }

            if (soxTracker.R3PBC != string.Empty && soxTracker.R3PBC != null && soxTracker.R3PBC != "Select Option")
            {


                var checkItem = _soxContext.SoxTrackerAppRelationship.Where(x => x.Title.Equals(soxTracker.R3PBC)).FirstOrDefault();
                if (checkItem != null)
                {
                    List<int> listId = new List<int>();
                    var appItem = soxTrackerItem.Field<AppItemField>(q12DField);
                    listId.Add(checkItem.PodioItemId);
                    appItem.ItemIds = listId;
                }
            }
            else //Remove value
            {
                if (update)
                {
                    var appItem = soxTrackerItem.Field<AppItemField>(q12DField);
                    List<int> listId = new List<int>();
                    appItem.ItemIds = listId;
                }
            }

            if (soxTracker.WTTester != string.Empty && soxTracker.WTTester != null && soxTracker.WTTester != "Select Option")
            {
                List<int> listItem = new List<int>();

                var checkItem = _soxContext.SoxTrackerAppRelationship.Where(x => x.Title.Equals(soxTracker.WTTester)).FirstOrDefault();
                if (checkItem != null)
                {
                    var appItem = soxTrackerItem.Field<AppItemField>(q13AField);
                    listItem.Add(checkItem.PodioItemId);
                    appItem.ItemIds = listItem;
                }
            }
            else //Remove value
            {
                if (update)
                {
                    var appItem = soxTrackerItem.Field<AppItemField>(q13AField);
                    List<int> listId = new List<int>();
                    appItem.ItemIds = listId;
                }
            }

            if (soxTracker.R1Tester != string.Empty && soxTracker.R1Tester != null && soxTracker.R1Tester != "Select Option")
            {


                var checkItem = _soxContext.SoxTrackerAppRelationship.Where(x => x.Title.Equals(soxTracker.R1Tester)).FirstOrDefault();
                if (checkItem != null)
                {
                    List<int> listId = new List<int>();
                    var appItem = soxTrackerItem.Field<AppItemField>(q13BField);
                    listId.Add(checkItem.PodioItemId);
                    appItem.ItemIds = listId;
                }
            }
            else //Remove value
            {
                if (update)
                {
                    var appItem = soxTrackerItem.Field<AppItemField>(q13BField);
                    List<int> listId = new List<int>();
                    appItem.ItemIds = listId;
                }
            }

            if (soxTracker.R2Tester != string.Empty && soxTracker.R2Tester != null && soxTracker.R2Tester != "Select Option")
            {


                var checkItem = _soxContext.SoxTrackerAppRelationship.Where(x => x.Title.Equals(soxTracker.R2Tester)).FirstOrDefault();
                if (checkItem != null)
                {
                    List<int> listId = new List<int>();
                    var appItem = soxTrackerItem.Field<AppItemField>(q13CField);
                    listId.Add(checkItem.PodioItemId);
                    appItem.ItemIds = listId;
                }
            }
            else //Remove value
            {
                if (update)
                {
                    var appItem = soxTrackerItem.Field<AppItemField>(q13CField);
                    List<int> listId = new List<int>();
                    appItem.ItemIds = listId;
                }
            }

            if (soxTracker.R3Tester != string.Empty && soxTracker.R3Tester != null && soxTracker.R3Tester != "Select Option")
            {


                var checkItem = _soxContext.SoxTrackerAppRelationship.Where(x => x.Title.Equals(soxTracker.R3Tester)).FirstOrDefault();
                if (checkItem != null)
                {
                    List<int> listId = new List<int>();
                    var appItem = soxTrackerItem.Field<AppItemField>(q13DField);
                    listId.Add(checkItem.PodioItemId);
                    appItem.ItemIds = listId;
                }
            }
            else //Remove value
            {
                if (update)
                {
                    var appItem = soxTrackerItem.Field<AppItemField>(q13DField);
                    List<int> listId = new List<int>();
                    appItem.ItemIds = listId;
                }
            }

            if (soxTracker.WT1LReviewer != string.Empty && soxTracker.WT1LReviewer != null && soxTracker.WT1LReviewer != "Select Option")
            {
                List<int> listWtPBC = new List<int>();

                var checWTPPBC = _soxContext.SoxTrackerAppRelationship.Where(x => x.Title.Equals(soxTracker.WT1LReviewer)).FirstOrDefault();
                if (checWTPPBC != null)
                {
                    var wtPBC = soxTrackerItem.Field<AppItemField>(q13EField);
                    listWtPBC.Add(checWTPPBC.PodioItemId);
                    wtPBC.ItemIds = listWtPBC;
                }
            }
            else //Remove value
            {
                if (update)
                {
                    var wtPBC = soxTrackerItem.Field<AppItemField>(q13EField);
                    List<int> listId = new List<int>();
                    wtPBC.ItemIds = listId;
                }
            }


            if (soxTracker.R11LReviewer != string.Empty && soxTracker.R11LReviewer != null && soxTracker.R11LReviewer != "Select Option")
            {


                var checkItem = _soxContext.SoxTrackerAppRelationship.Where(x => x.Title.Equals(soxTracker.R11LReviewer)).FirstOrDefault();
                if (checkItem != null)
                {
                    List<int> listId = new List<int>();
                    var appItem = soxTrackerItem.Field<AppItemField>(q13FField);
                    listId.Add(checkItem.PodioItemId);
                    appItem.ItemIds = listId;
                }
            }
            else //Remove value
            {
                if (update)
                {
                    var appItem = soxTrackerItem.Field<AppItemField>(q13FField);
                    List<int> listId = new List<int>();
                    appItem.ItemIds = listId;
                }
            }
            
            if (soxTracker.R21LReviewer != string.Empty && soxTracker.R21LReviewer != null && soxTracker.R21LReviewer != "Select Option")
            {


                var checkItem = _soxContext.SoxTrackerAppRelationship.Where(x => x.Title.Equals(soxTracker.R21LReviewer)).FirstOrDefault();
                if (checkItem != null)
                {
                    List<int> listId = new List<int>();
                    var appItem = soxTrackerItem.Field<AppItemField>(q13GField);
                    listId.Add(checkItem.PodioItemId);
                    appItem.ItemIds = listId;
                }
            }
            else //Remove value
            {
                if (update)
                {
                    var appItem = soxTrackerItem.Field<AppItemField>(q13GField);
                    List<int> listId = new List<int>();
                    appItem.ItemIds = listId;
                }
            }


            if (soxTracker.R31LReviewer != string.Empty && soxTracker.R31LReviewer != null && soxTracker.R31LReviewer != "Select Option")
            {


                var checkItem = _soxContext.SoxTrackerAppRelationship.Where(x => x.Title.Equals(soxTracker.R31LReviewer)).FirstOrDefault();
                if (checkItem != null)
                {
                    List<int> listId = new List<int>();
                    var appItem = soxTrackerItem.Field<AppItemField>(q13HField);
                    listId.Add(checkItem.PodioItemId);
                    appItem.ItemIds = listId;
                }
            }
            else //Remove value
            {
                if (update)
                {
                    var appItem = soxTrackerItem.Field<AppItemField>(q13HField);
                    List<int> listId = new List<int>();
                    appItem.ItemIds = listId;
                }
            }

            //debugging
            if (soxTracker.WT2LReviewer != string.Empty && soxTracker.WT2LReviewer != null && soxTracker.WT2LReviewer != "Select Option")
            {
                List<int> listWtPBC = new List<int>();

                var checWTPPBC = _soxContext.SoxTrackerAppRelationship.Where(x => x.Title.Equals(soxTracker.WT2LReviewer)).FirstOrDefault();
                if (checWTPPBC != null)
                {
                    var wtPBC = soxTrackerItem.Field<AppItemField>(q13IField);
                    listWtPBC.Add(checWTPPBC.PodioItemId);
                    wtPBC.ItemIds = listWtPBC;
                }
            }
            else //Remove value
            {
                if (update)
                {
                    var wtPBC = soxTrackerItem.Field<AppItemField>(q13IField);
                    List<int> listId = new List<int>();
                    wtPBC.ItemIds = listId;
                }
            }


            if (soxTracker.R12LReviewer != string.Empty && soxTracker.R12LReviewer != null && soxTracker.R12LReviewer != "Select Option")
            {


                var checkItem = _soxContext.SoxTrackerAppRelationship.Where(x => x.Title.Equals(soxTracker.R12LReviewer)).FirstOrDefault();
                if (checkItem != null)
                {
                    List<int> listId = new List<int>();
                    var appItem = soxTrackerItem.Field<AppItemField>(q13JField);
                    listId.Add(checkItem.PodioItemId);
                    appItem.ItemIds = listId;
                }
            }
            else //Remove value
            {
                if (update)
                {
                    var appItem = soxTrackerItem.Field<AppItemField>(q13JField);
                    List<int> listId = new List<int>();
                    appItem.ItemIds = listId;
                }
            }

            if (soxTracker.R22LReviewer != string.Empty && soxTracker.R22LReviewer != null && soxTracker.R22LReviewer != "Select Option")
            {


                var checkItem = _soxContext.SoxTrackerAppRelationship.Where(x => x.Title.Equals(soxTracker.R22LReviewer)).FirstOrDefault();
                if (checkItem != null)
                {
                    List<int> listId = new List<int>();
                    var appItem = soxTrackerItem.Field<AppItemField>(q13KField);
                    listId.Add(checkItem.PodioItemId);
                    appItem.ItemIds = listId;
                }
            }
            else //Remove value
            {
                if (update)
                {
                    var appItem = soxTrackerItem.Field<AppItemField>(q13KField);
                    List<int> listId = new List<int>();
                    appItem.ItemIds = listId;
                }
            }


            if (soxTracker.R32LReviewer != string.Empty && soxTracker.R32LReviewer != null && soxTracker.R32LReviewer != "Select Option")
            {


                var checkItem = _soxContext.SoxTrackerAppRelationship.Where(x => x.Title.Equals(soxTracker.R32LReviewer)).FirstOrDefault();
                if (checkItem != null)
                {
                    List<int> listId = new List<int>();
                    var appItem = soxTrackerItem.Field<AppItemField>(q13LField);
                    listId.Add(checkItem.PodioItemId);
                    appItem.ItemIds = listId;
                }
            }
            else //Remove value
            {
                if (update)
                {
                    var appItem = soxTrackerItem.Field<AppItemField>(q13LField);
                    List<int> listId = new List<int>();
                    appItem.ItemIds = listId;
                }
            }

            List<string> listRevChecklist = new List<string>();
            if (soxTracker.RCRWT != 0) { listRevChecklist.Add("Walkthrough"); }
            if (soxTracker.RCRR1 != 0) { listRevChecklist.Add("Round 1"); }
            if (soxTracker.RCRR2 != 0) { listRevChecklist.Add("Round 2"); }
            if (soxTracker.RCRR3 != 0) { listRevChecklist.Add("Round 3"); }

            if (listRevChecklist.Any())
            {
                var revChecklist = soxTrackerItem.Field<CategoryItemField>(q14AField);
                revChecklist.OptionTexts = listRevChecklist;
            }
            else //Remove value
            {
                if (update)
                {
                    var revChecklist = soxTrackerItem.Field<CategoryItemField>(q14AField);
                    List<int> listId = new List<int>();
                    revChecklist.OptionIds = listId;
                }

            }

            if (soxTracker.WTTestingStatus != string.Empty && soxTracker.WTTestingStatus != null && soxTracker.WTTestingStatus != "Select Option")
            {
                List<int> listWtPBC = new List<int>();

                var checWTPPBC = _soxContext.SoxTrackerAppRelationship.Where(x => x.Title.Equals(soxTracker.WTTestingStatus)).FirstOrDefault();
                if (checWTPPBC != null)
                {
                    var wtPBC = soxTrackerItem.Field<AppItemField>(q14BField);
                    listWtPBC.Add(checWTPPBC.PodioItemId);
                    wtPBC.ItemIds = listWtPBC;
                }
            }
            else //Remove value
            {
                if (update)
                {
                    var wtPBC = soxTrackerItem.Field<AppItemField>(q14BField);
                    List<int> listId = new List<int>();
                    wtPBC.ItemIds = listId;
                }
            }

            if (soxTracker.R1TestingStatus != string.Empty && soxTracker.R1TestingStatus != null && soxTracker.R1TestingStatus != "Select Option")
            {


                var checkItem = _soxContext.SoxTrackerAppRelationship.Where(x => x.Title.Equals(soxTracker.R1TestingStatus)).FirstOrDefault();
                if (checkItem != null)
                {
                    List<int> listId = new List<int>();
                    var appItem = soxTrackerItem.Field<AppItemField>(q14CField);
                    listId.Add(checkItem.PodioItemId);
                    appItem.ItemIds = listId;
                }
            }
            else //Remove value
            {
                if (update)
                {
                    var appItem = soxTrackerItem.Field<AppItemField>(q14CField);
                    List<int> listId = new List<int>();
                    appItem.ItemIds = listId;
                }
            }

            if (soxTracker.R2TestingStatus != string.Empty && soxTracker.R2TestingStatus != null && soxTracker.R2TestingStatus != "Select Option")
            {


                var checkItem = _soxContext.SoxTrackerAppRelationship.Where(x => x.Title.Equals(soxTracker.R2TestingStatus)).FirstOrDefault();
                if (checkItem != null)
                {
                    List<int> listId = new List<int>();
                    var appItem = soxTrackerItem.Field<AppItemField>(q14DField);
                    listId.Add(checkItem.PodioItemId);
                    appItem.ItemIds = listId;
                }
            }
            else //Remove value
            {
                if (update)
                {
                    var appItem = soxTrackerItem.Field<AppItemField>(q14DField);
                    List<int> listId = new List<int>();
                    appItem.ItemIds = listId;
                }
            }

            if (soxTracker.R3TestingStatus != string.Empty && soxTracker.R3TestingStatus != null && soxTracker.R3TestingStatus != "Select Option")
            {


                var checkItemR3TestingStat = _soxContext.SoxTrackerAppRelationship.Where(x => x.Title.Equals(soxTracker.R3TestingStatus)).FirstOrDefault();
                if (checkItemR3TestingStat != null)
                {
                    List<int> listId = new List<int>();
                    var appItem = soxTrackerItem.Field<AppItemField>(q14EField);
                    listId.Add(checkItemR3TestingStat.PodioItemId);
                    appItem.ItemIds = listId;
                }
            }
            else //Remove value
            {
                if (update)
                {
                    var appItem = soxTrackerItem.Field<AppItemField>(q14EField);
                    List<int> listId = new List<int>();
                    appItem.ItemIds = listId;
                }
            }

            if (soxTracker.Duration != null && soxTracker.Duration.HasValue)
            {
                DurationItemField durationField = soxTrackerItem.Field<DurationItemField>(duration);
                durationField.Value = soxTracker.Duration;
            }
            

            #endregion
        }

        //[AllowAnonymous]
        [HttpGet("view_sox_tracker/{fy}/{ClientName}")]
        public IActionResult fetch_records_sox_tracker(String fy, String ClientName)
        {

            //return "test";
            List<SoxTracker> _soxtracker = new List<SoxTracker>();
            try
            {
                using (var context = _soxContext.Database.BeginTransaction())
                {

                    var listSoxTracker = _soxContext.SoxTracker.Where(x =>
                        x.FY.Equals(fy)
                        && x.ClientName.Equals(ClientName)

                    )
                    .OrderBy(x => x.Process)
                    .ThenBy(x => x.ControlId);

                    if (listSoxTracker != null)
                    {
                        _soxtracker = listSoxTracker.ToList();
                    }
                }
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetListClientAsync {ex}", "ErrorGetListClientAsync");
            }

            if (_soxtracker != null)
            {
                return Ok(_soxtracker.ToArray());
            }
            else
            {
                return NoContent();
            }

        }


    }
}
