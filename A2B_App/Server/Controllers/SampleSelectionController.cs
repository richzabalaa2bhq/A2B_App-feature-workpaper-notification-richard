using A2B_App.Server.Data;
//using A2B_App.Server.Log;
using A2B_App.Server.Services;
using A2B_App.Shared.Podio;
using A2B_App.Shared.Sox;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using PodioAPI;
using PodioAPI.Models;
using PodioAPI.Utils.ItemFields;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading.Tasks;

namespace A2B_App.Server.Controllers
{
    [Authorize]
    [ApiController]
    [Route("api/[controller]")]
    
    public class SampleSelectionController : ControllerBase
    {

        private readonly IConfiguration _config;
        private readonly ILogger<SampleSelectionController> _logger;
        private readonly IWebHostEnvironment _environment;
        private readonly SoxContext _soxContext;

        public SampleSelectionController(IConfiguration config,
            ILogger<SampleSelectionController> logger,
            IWebHostEnvironment environment,
            SoxContext soxContext)
        {
            _config = config;
            _logger = logger;
            _environment = environment;
            _soxContext = soxContext;
        }

        [AllowAnonymous]
        [HttpGet("download/{filename}")]
        public async Task<IActionResult> GetDownloadAsync(string filename)
        {
            //test link = https://localhost:44344/SampleSelection/download/Draft-TestRound-20200408_183306.xlsx

            try
            {
                string startupPath = Directory.GetCurrentDirectory();
                string path = Path.Combine(startupPath, "include", "sampleselection", "download", filename);

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
                FileLog.Write($"Error GetDownloadAsync {ex}", "ErrorGetDownloadAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetDownloadAsync");
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

        [AllowAnonymous]
        [HttpGet("download/{folder}/{filename}")]
        public async Task<IActionResult> GetDownloadAsync2(string folder, string filename)
        {
            //test link = https://localhost:44344/SampleSelection/download/Draft-TestRound-20200408_183306.xlsx

            try
            {
                string startupPath = Directory.GetCurrentDirectory();
                string path = Path.Combine(startupPath, "include", folder, filename);

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
                FileLog.Write($"Error GetDownloadAsync2 {ex}", "ErrorGetDownloadAsync2");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetDownloadAsync2");
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

        [HttpGet("data/rcm/sampleselection/{itemId}")]
        public IActionResult GetSampleSelection(string itemId)
        {
            try
            {
                var sampleSelectionCheck = _soxContext.SampleSelection
                    .OrderByDescending(x => x.Id)
                    .Where(id => id.PodioItemId == int.Parse(itemId))
                    .Include(x => x.ListTestRound)
                    .FirstOrDefault();
                if (sampleSelectionCheck != null)
                {
                    return Ok(sampleSelectionCheck);
                }
                else
                {
                    return BadRequest($"Bad request..");
                }

            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetSampleSelection {ex}", "ErrorGetSampleSelection");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetSampleSelection");
                return BadRequest($"Bad request..");
            }

        }

        //[AllowAnonymous]
        [HttpGet("data/samplesize")]
        public IEnumerable<SampleSize> GetSampleSize()
        {
            List<SampleSize> listSampleSize = new List<SampleSize>();
            try
            {
                var matrix = _soxContext.Matrix.Select(x => new {
                    x.ExternalAuditor,
                    x.Frequency,
                    x.Risk,
                    x.SizeValue,
                    x.StartPopulation,
                    x.EndPopulation
                });

                foreach (var item in matrix)
                {
                    SampleSize sampleSize = new SampleSize();
                    sampleSize.ExternalAuditor = item.ExternalAuditor;
                    sampleSize.Frequency = item.Frequency;
                    sampleSize.Risk = item.Risk;
                    sampleSize.SizeValue = item.SizeValue;
                    sampleSize.StartPopulation = item.StartPopulation;
                    sampleSize.EndPopulation = item.EndPopulation;
                    listSampleSize.Add(sampleSize);
                }


            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetSampleSize {ex}", "ErrorGetSampleSize");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetSampleSize");

            }
            return listSampleSize.ToArray();
        }

        [HttpGet("data/clientsamplesize/{clientName}")]
        public IEnumerable<SampleSize> GetClientSampleSize(string clientName)
        {
            List<SampleSize> listSampleSize = new List<SampleSize>();
            try
            {
                var matrix = _soxContext.Matrix
                    .Where(x => x.ClientName.ToLower().Equals(clientName.ToLower()))
                    .Select(x => new {
                        x.ExternalAuditor,
                        x.Frequency,
                        x.Risk,
                        x.SizeValue,
                        x.StartPopulation,
                        x.EndPopulation
                });

                foreach (var item in matrix)
                {
                    SampleSize sampleSize = new SampleSize();
                    sampleSize.ExternalAuditor = item.ExternalAuditor;
                    sampleSize.Frequency = item.Frequency;
                    sampleSize.Risk = item.Risk;
                    sampleSize.SizeValue = item.SizeValue;
                    sampleSize.StartPopulation = item.StartPopulation;
                    sampleSize.EndPopulation = item.EndPopulation;
                    listSampleSize.Add(sampleSize);
                }


            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetClientSampleSize {ex}", "ErrorGetClientSampleSize");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetClientSampleSize");

            }
            return listSampleSize.ToArray();
        }

        [HttpGet("data/dropdown")]
        public IEnumerable<DropDown> GetDropDown()
        {

            List<DropDown> listDropDown = new List<DropDown>();

            try
            {

                ////Set directory and get matrix.xlsx using file info
                //string startupPath = Directory.GetCurrentDirectory();
                //string strSourceMatrix = Path.Combine(startupPath, "include", "sampleselection", "matrix", "dropdown.xlsx");

                //var fi = new FileInfo(strSourceMatrix);

                //// If you use EPPlus in a noncommercial context
                //// according to the Polyform Noncommercial license:
                //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                //ExcelPackage p = new ExcelPackage(fi);

                //ExcelWorksheet ws = p.Workbook.Worksheets["dropdown"];
                //int colCount = ws.Dimension.End.Column;
                //int rowCount = ws.Dimension.End.Row;
                //for (int row = 2; row <= rowCount; row++)
                //{
                //    if (ws.Cells[row, 1].Value != null)
                //    {
                //        DropDown dropdown = new DropDown();
                //        for (int col = 1; col <= colCount; col++)
                //        {
                //            switch (col)
                //            {

                //                case 1:
                //                    if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                //                    {
                //                        dropdown.ExternalAuditor = ws.Cells[row, col].Value?.ToString();
                //                    }
                //                    break;
                //                case 2:
                //                    if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                //                    {
                //                        dropdown.Percent = int.Parse(ws.Cells[row, col].Value?.ToString());
                //                    }
                //                    break;
                //            }
                //        }
                //        if (dropdown != null)
                //        {
                //            listDropDown.Add(dropdown);
                //        }

                //    }
                //}


                var clientSs = _soxContext.ClientSs.Select(x => new { x.ExternalAuditor, x.Percent, x.PercentRound2 });

                foreach (var item in clientSs)
                {
                    DropDown dropDown = new DropDown();
                    dropDown.ExternalAuditor = item.ExternalAuditor;
                    dropDown.Percent = item.Percent;
                    dropDown.PercentRound2 = item.PercentRound2;
                    listDropDown.Add(dropDown);
                }
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetDropDown {ex}", "ErrorGetDropDown");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetDropDown");
            }

            return listDropDown.ToArray();

        }

        [HttpGet("data/client")]
        public IEnumerable<ClientSs> GetClient()
        {
            //List<ClientSS> listClient = null;
            List<ClientSs> listClientSs = null;
            try
            {

                //listClient = new List<ClientSS>
                //{
                //    new ClientSS
                //    {
                //        ClientName = "",
                //        ExternalAuditor = string.Empty,
                //        ItemId = null
                //    },
                //    new ClientSS
                //    {
                //        ClientName= "Assetmark",
                //        ExternalAuditor= "KPMG",
                //        ItemId = 1384964278
                //    },
                //    new ClientSS
                //    {
                //        ClientName= "ERI",
                //        ExternalAuditor= "DT-ERI",
                //        ItemId = 1384965458
                //    },
                //    new ClientSS
                //    {
                //        ClientName= "Kindred Bio",
                //        ExternalAuditor= "KMJ",
                //        ItemId = 1384965878
                //    },
                //    new ClientSS
                //    {
                //        ClientName= "ViewRay",
                //        ExternalAuditor= "DT-VR",
                //        ItemId = 1384966696
                //    },
                //    new ClientSS
                //    {
                //        ClientName= "McGrath",
                //        ExternalAuditor= "Grant Thornton",
                //        ItemId = 1421150584
                //    }
                //};

                //listClient = new List<ClientSS>();

                listClientSs = _soxContext.ClientSs
                    .OrderBy(x => x.ClientName)
                    .Select(
                        x => new ClientSs
                        {
                            ClientName = x.ClientName,
                            ClientCode = x.ClientCode,
                            ClientItemId = x.ClientItemId,
                            ItemId = x.ItemId,
                            ExternalAuditor = x.ExternalAuditor,
                            Percent = x.Percent,
                            PercentRound2 = x.PercentRound2,
                            PodioItemId = x.PodioItemId
                        }).ToList();

            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetClient {ex}", "ErrorGetClient");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetClient");
            }

            //return listClient.ToArray();
            return listClientSs.ToArray();


        }

        [HttpGet("data/clientdetails/{clientName}")]
        public IEnumerable<ClientSs> GetClientDetails(string clientName)
        {
            //List<ClientSS> listClient = null;
            List<ClientSs> listClientSs = null;
            try
            {

                //listClient = new List<ClientSS>
                //{
                //    new ClientSS
                //    {
                //        ClientName = "",
                //        ExternalAuditor = string.Empty,
                //        ItemId = null
                //    },
                //    new ClientSS
                //    {
                //        ClientName= "Assetmark",
                //        ExternalAuditor= "KPMG",
                //        ItemId = 1384964278
                //    },
                //    new ClientSS
                //    {
                //        ClientName= "ERI",
                //        ExternalAuditor= "DT-ERI",
                //        ItemId = 1384965458
                //    },
                //    new ClientSS
                //    {
                //        ClientName= "Kindred Bio",
                //        ExternalAuditor= "KMJ",
                //        ItemId = 1384965878
                //    },
                //    new ClientSS
                //    {
                //        ClientName= "ViewRay",
                //        ExternalAuditor= "DT-VR",
                //        ItemId = 1384966696
                //    },
                //    new ClientSS
                //    {
                //        ClientName= "McGrath",
                //        ExternalAuditor= "Grant Thornton",
                //        ItemId = 1421150584
                //    }
                //};

                //listClient = new List<ClientSS>();

                listClientSs = _soxContext.ClientSs
                    .OrderBy(x => x.ClientName)
                    .Where(x => x.ClientName.ToLower().Equals(clientName.ToLower()))
                    .Select(
                        x => new ClientSs
                        {
                            ClientName = x.ClientName,
                            ClientCode = x.ClientCode,
                            ClientItemId = x.ClientItemId,
                            ItemId = x.ItemId,
                            ExternalAuditor = x.ExternalAuditor,
                            Percent = x.Percent,
                            PercentRound2 = x.PercentRound2,
                            PodioItemId = x.PodioItemId
                        }).ToList();

            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetClientDetails {ex}", "ErrorGetClientDetails");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetClientDetails");
            }

            //return listClient.ToArray();
            return listClientSs.ToArray();


        }

        [HttpGet("data/frequency")]
        public IEnumerable<Frequency> GetFrequency()
        {
            List<Frequency> listFrequency = null;
            try
            {
                listFrequency = new List<Frequency>
                {
                    new Frequency{ Freq = "", IntValue = null},
                    new Frequency{ Freq = "Daily", IntValue = 366},
                    new Frequency{ Freq = "Weekly", IntValue = 52},
                    new Frequency{ Freq = "Monthly", IntValue = 12},
                };
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetFrequency {ex}", "ErrorGetFrequency");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetFrequency");
            }

            return listFrequency.ToArray();

        }

        [HttpGet("download/population/{filename}/{round}")]
        public IEnumerable<Population> GetPopulation(string filename, int round)
        {
            List<Population> listPopulation = new List<Population>();

            try
            {
                //Set directory and {filename}.xlsx using file info
                string startupPath = Directory.GetCurrentDirectory();
                string path = Path.Combine(startupPath, "include", "upload", filename);

                var fi = new FileInfo(path);

                using (ExcelPackage p = new ExcelPackage(fi))
                {
                    // If you use EPPlus in a noncommercial context
                    // according to the Polyform Noncommercial license:
                    //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    //Console.WriteLine($"Worksheet Count: {p.Workbook.Worksheets.Count}");
                    ExcelWorksheet ws = p.Workbook.Worksheets[round - 1];
                    listPopulation = ExcelPopulation(ws, round);
                }


                //ExcelWorksheet wsR2 = p.Workbook.Worksheets[2];
                //listPopulation.AddRange(ExcelPopulation(wsR2));

                //ExcelWorksheet wsR3 = p.Workbook.Worksheets["Population File (R3)"];
                //listPopulation.AddRange(ExcelPopulation(wsR3, 3));

            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetPopulation {ex}", "ErrorGetPopulation");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetPopulation");
            }

            return listPopulation.ToArray();

        }

        //[AllowAnonymous]
        [HttpGet("download/population2/{filename}/{round}")]
        public IEnumerable<List<string>> GetPopulation2(string filename, int round)
        {
            List<List<string>> listPopulation = new List<List<string>>();

            try
            {
                //Set directory and {filename}.xlsx using file info
                string startupPath = Directory.GetCurrentDirectory();
                string path = Path.Combine(startupPath, "include", "upload", filename);

                var fi = new FileInfo(path);

                using (ExcelPackage p = new ExcelPackage(fi))
                {
                    // If you use EPPlus in a noncommercial context
                    // according to the Polyform Noncommercial license:
                    //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    //Console.WriteLine($"Worksheet Count: {p.Workbook.Worksheets.Count}");
                    ExcelWorksheet ws = p.Workbook.Worksheets[round - 1];
                    listPopulation = ExcelPopulation2(ws, round);
                }

            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetPopulation2 {ex}", "ErrorGetPopulation2");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetPopulation2");
            }

            return listPopulation;

        }

        [HttpGet("data/risk")]
        public IEnumerable<string> GetRisk()
        {
            List<string> listRisk = null;
            try
            {
                listRisk = new List<string> { "", "High", "Medium", "Low" };
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetRisk {ex}", "ErrorGetRisk");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetRisk");
            }
            return listRisk.ToArray();
        }

        [HttpGet("data/q4r3")]
        public IEnumerable<string> GetQ4R3()
        {
            List<string> listQ4R3 = null;
            try
            {
                listQ4R3 = new List<string> { "", "Yes", "No" };
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetQ4R3 {ex}", "ErrorGetQ4R3");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetQ4R3");
            }
            return listQ4R3.ToArray();
        }

        [HttpGet("data/listsampleselection")]
        public IEnumerable<SampleSelection> GetListSampleSelection()
        {
            List<SampleSelection> listSampleSelection = null;
            try
            {
                listSampleSelection = _soxContext.SampleSelection.ToList();
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetListSampleSelection {ex}", "ErrorGetListSampleSelection");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetListSampleSelection");
            }
            return listSampleSelection.ToArray();
        }

        //[AllowAnonymous]
        [HttpPost("excel/create")]
        public string CreateExcel([FromBody] SampleSelection sampleSelection)
        {
            //List<string> excelFilename = new List<string>();
            string excelFilename = string.Empty;

            try
            {

                //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage xls = new ExcelPackage())
                {

                    var ws = xls.Workbook.Worksheets.Add("TestRounds");

                    ws.Column(1).Width = 13;
                    ws.Column(2).Width = 16;
                    ws.Column(3).Width = 25;
                    ws.Column(4).Width = 25;
                    ws.Column(5).Width = 25;
                    ws.Column(6).Width = 25;
                    ws.Column(7).Width = 25;
                    ws.Column(8).Width = 25;
                    ws.Column(9).Width = 25;
                    ws.Column(10).Width = 25;
                    ws.View.ShowGridLines = false;

                    #region Sample Selection Input
                    ws.Row(1).Height = 63;
                    ws.Row(2).Height = 14;
                    ws.Row(3).Height = 47;
                    ws.Row(4).Height = 5;
                    ws.Row(5).Height = 17;
                    ws.Row(6).Height = 17;
                    ws.Row(7).Height = 17;
                    ws.Row(8).Height = 17;
                    ws.Row(9).Height = 30;
                    ws.Row(10).Height = 30;
                    ws.Row(11).Height = 30;
                    ws.Row(12).Height = (sampleSelection.Version != "3" ? 15 : 52);
                    ws.Row(13).Height = (sampleSelection.Version != "3" ? 15 : 52);
                    ws.Row(14).Height = 15;

                    ws.Cells["A1"].Value = "Client";
                    ws.Cells["B1"].Value = sampleSelection.ClientName;
                    ExcelSetBorder(ws, "A1:B1");
                    ExcelSetAlignCenter(ws, "A1:B1");
                    ExcelSetArialSize12(ws, "A1:B1");
                    ExcelSetFontColorRed(ws, "A1");
                    ExcelSetBackgroundGreen(ws, "A1:B1");
                    ExcelSetFontColorGreen(ws, "B1");
                    ExcelSetFontBold(ws, "A1:B1");

                    ws.Cells["D1"].Value = "External Auditor";
                    ws.Cells["E1"].Value = sampleSelection.ExternalAuditor;
                    ExcelSetBackgroundBlue(ws, "D1:E1");
                    ExcelSetBorder(ws, "D1:E1");
                    ExcelSetAlignCenter(ws, "D1:E1");
                    ExcelSetArialSize12(ws, "D1:E1");
                    ExcelSetFontColorRed(ws, "E1");
                    ExcelSetFontBold(ws, "D1:E1");

                    ws.Cells["G1"].Value = "Q4 (R3) Sample Required?";
                    ws.Cells["H1"].Value = sampleSelection.Q4R3SampleRequired;
                    ws.Cells["I1"].Value = "How many samples to be tested in Q4?";
                    ws.Cells["J1"].Value = sampleSelection.CountSampleQ4R3;
                    ExcelSetBackgroundBlue(ws, "G1:J1");
                    ExcelSetBackgroundGreen(ws, "H1");
                    ExcelSetBackgroundGreen(ws, "J1");
                    ExcelSetBorder(ws, "G1:J1");
                    ExcelSetAlignCenter(ws, "G1:J1");
                    ExcelSetArialSize12(ws, "G1:J1");
                    ExcelSetFontColorGreen(ws, "H1");
                    ExcelSetFontColorGreen(ws, "J1");
                    ExcelSetFontBold(ws, "G1:J1");
                    ExcelWrapText(ws, "G1:J1");

                    ws.Cells["A3"].Value = "Risk";
                    ws.Cells["B3"].Value = sampleSelection.Risk;
                    ws.Cells["C3"].Value = "Annual Population";
                    ws.Cells["D3"].Value = sampleSelection.AnnualPopulation;
                    ws.Cells["E3"].Value = "Annual Sample Size";
                    ws.Cells["F3"].Value = sampleSelection.AnnualSampleSize;
                    ws.Cells["G3"].Value = "Frequency";
                    ws.Cells["H3"].Value = sampleSelection.Frequency;
                    ExcelSetBackgroundBlue(ws, "C3:H3");
                    ExcelSetBackgroundGreen(ws, "A3:B3");
                    ExcelSetBackgroundGreen(ws, "H3");
                    ExcelSetFontColorRed(ws, "A3");
                    ExcelSetFontColorRed(ws, "D3");
                    ExcelSetFontColorRed(ws, "F3");
                    ExcelSetFontColorGreen(ws, "B3");
                    ExcelSetFontColorGreen(ws, "H3");
                    ExcelSetBorder(ws, "A3:H3");
                    ExcelSetAlignCenter(ws, "A3:H3");
                    ExcelSetArialSize12(ws, "A3:H3");
                    ExcelSetFontBold(ws, "A3:H3");
                    ExcelWrapText(ws, "A3:H3");


                    ws.Cells["C5"].Value = "Round 1";
                    ws.Cells["D5"].Value = "Round 2";
                    ws.Cells["E5"].Value = "Round 3";
                    ws.Cells["F5"].Value = "Total";
                    ExcelSetBackgroundBlue(ws, "C5:F5");
                    ExcelSetBorder(ws, "C5:F5");
                    ExcelSetAlignCenter(ws, "C5:F5");
                    ExcelSetArialSize12(ws, "C5:F5");
                    ExcelSetFontBold(ws, "C5:F5");
                    ExcelWrapText(ws, "C5:F5");

                    ws.Cells["A6"].Value = "Population by Start Date";
                    ws.Cells["A6:B6"].Merge = true;
                    ws.Cells["C6"].Value = (sampleSelection.Round1Start != null ? sampleSelection.Round1Start.Value.ToString("MM/dd/yy") : "");
                    ws.Cells["D6"].Value = (sampleSelection.Round2Start != null ? sampleSelection.Round2Start.Value.ToString("MM/dd/yy") : "");
                    ws.Cells["E6"].Value = (sampleSelection.Round3Start != null ? sampleSelection.Round3Start.Value.ToString("MM/dd/yy") : "");
                    ExcelSetBackgroundGreen(ws, "A6:E6");
                    ExcelSetFontColorRed(ws, "A6:B6");
                    ExcelSetFontColorGreen(ws, "C6:E6");
                    ExcelSetBorder(ws, "A6:F6");
                    ExcelSetAlignCenter(ws, "A6:F6");
                    ExcelSetArialSize12(ws, "A6:F6");
                    ExcelSetFontBold(ws, "A6:F6");
                    ExcelWrapText(ws, "A6:F6");

                    ws.Cells["A7"].Value = "Population by End Date";
                    ws.Cells["A7:B7"].Merge = true;
                    ws.Cells["F6:F7"].Merge = true;
                    ws.Cells["C7"].Value = (sampleSelection.Round1End != null ? sampleSelection.Round1End.Value.ToString("MM/dd/yy") : "");
                    ws.Cells["D7"].Value = (sampleSelection.Round2End != null ? sampleSelection.Round2End.Value.ToString("MM/dd/yy") : "");
                    ws.Cells["E7"].Value = (sampleSelection.Round3End != null ? sampleSelection.Round3End.Value.ToString("MM/dd/yy") : "");
                    ws.Cells["F6"].Value = (sampleSelection.DaysPeriodRoundTot.ToString() != string.Empty ? sampleSelection.DaysPeriodRoundTot.ToString() : "");
                    ExcelSetBackgroundGreen(ws, "A7:E7");
                    ExcelSetFontColorRed(ws, "A7:B7");
                    ExcelSetFontColorGreen(ws, "C7:E7");
                    ExcelSetBorder(ws, "A7:F7");
                    ExcelSetAlignCenter(ws, "A7:F7");
                    ExcelSetArialSize12(ws, "A7:F7");
                    ExcelSetFontBold(ws, "A7:F7");
                    ExcelWrapText(ws, "A7:F7");
                    ExcelSetBackgroundBlue(ws, "F6:F7");

                    ws.Cells["A8"].Value = "Population by Round";
                    ws.Cells["A8:B8"].Merge = true;
                    ws.Cells["C8"].Value = sampleSelection.PopulationByRound1;
                    ws.Cells["D8"].Value = sampleSelection.PopulationByRound2;
                    ws.Cells["E8"].Value = sampleSelection.PopulationByRound3;
                    ws.Cells["F8"].Value = sampleSelection.PopulationByRoundTot;
                    ExcelSetBackgroundBlue(ws, "A8:F8");
                    ExcelSetBorder(ws, "A8:F8");
                    ExcelSetAlignCenter(ws, "A8:F8");
                    ExcelSetArialSize12(ws, "A8:F8");
                    ExcelSetFontBold(ws, "A8:F8");
                    ExcelWrapText(ws, "A8:F8");

                    ws.Cells["A9"].Value = "Samples by Round";
                    ws.Cells["A9:B9"].Merge = true;
                    ws.Cells["C9"].Value = sampleSelection.SamplesByRound1;
                    ws.Cells["D9"].Value = sampleSelection.SamplesByRound2;
                    ws.Cells["E9"].Value = sampleSelection.SamplesByRound3;
                    ws.Cells["F9"].Value = sampleSelection.SamplesByRoundTot;
                    ExcelSetBackgroundBlue(ws, "A9:F9");
                    ExcelSetBorder(ws, "A9:F9");
                    ExcelSetAlignCenter(ws, "A9:F9");
                    ExcelSetArialSize12(ws, "A9:F9");
                    ExcelSetFontBold(ws, "A9:F9");
                    ExcelWrapText(ws, "A9:F9");

                    ws.Cells["A10"].Value = "Samples Closed";
                    ws.Cells["A10:B10"].Merge = true;
                    ws.Cells["C10"].Value = sampleSelection.SamplesCloseRound1;
                    ws.Cells["D10"].Value = sampleSelection.SamplesCloseRound2;
                    ws.Cells["E10"].Value = sampleSelection.SamplesCloseRound3;
                    ws.Cells["F10"].Value = sampleSelection.SamplesCloseRoundTot;
                    ExcelSetBackgroundBlue(ws, "A10:F10");
                    ExcelSetBorder(ws, "A10:F10");
                    ExcelSetAlignCenter(ws, "A10:F10");
                    ExcelSetArialSize12(ws, "A10:F10");
                    ExcelSetFontBold(ws, "A10:F10");
                    ExcelWrapText(ws, "A10:F10");

                    ws.Cells["A11"].Value = "Samples Remaining";
                    ws.Cells["A11:B11"].Merge = true;

                    //Change from value to formula
                    //ws.Cells["C11"].Value = sampleSelection.SamplesRemainingRound1;
                    //ws.Cells["D11"].Value = sampleSelection.SamplesRemainingRound2;
                    //ws.Cells["E11"].Value = sampleSelection.SamplesRemainingRound3;
                    //ws.Cells["F11"].Value = sampleSelection.SamplesRemainingRoundTot;

                    ws.Cells["C11"].Formula = "C9-C10";
                    ws.Cells["D11"].Formula = "D9-D10";
                    ws.Cells["E11"].Formula = "E9-E10";
                    ws.Cells["F11"].Formula = "SUM(C11:E11)";
                    ws.Workbook.CalcMode = ExcelCalcMode.Automatic;
                    ws.Calculate();

                    ExcelSetBackgroundBlue(ws, "A11:F11");
                    ExcelSetBorder(ws, "A11:F11");
                    ExcelSetAlignCenter(ws, "A11:F11");
                    ExcelSetArialSize12(ws, "A11:F11");
                    ExcelSetFontBold(ws, "A11:F11");
                    ExcelWrapText(ws, "A11:F11");



                    //transactional and materiliaty
                    if (sampleSelection.Version == "3")
                    {
                        ws.Cells["A12"].Value = "Does the sample selection need to consider materiality?";
                        ws.Cells["A12:B12"].Merge = true;
                        ws.Cells["C12"].Value = (sampleSelection.IsMateriality != "Yes" ? "No" : sampleSelection.IsMateriality);
                        ExcelSetBackgroundBlue(ws, "A12:B12");
                        ExcelSetFontColorRed(ws, "A12:B12");
                        ExcelSetFontColorGreen(ws, "C12");
                        ExcelSetBackgroundGreen(ws, "C12");
                        ExcelSetBorder(ws, "A12:C12");
                        ExcelSetAlignCenter(ws, "A12:E12");
                        ExcelSetArialSize12(ws, "A12:E12");
                        ExcelSetFontBold(ws, "A12:E12");
                        ExcelWrapText(ws, "A12:E12");

                        ws.Cells["A13"].Value = "What to consider on materiality?";
                        ws.Cells["A13:B13"].Merge = true;
                        ws.Cells["C13"].Value = (sampleSelection.ConsiderMateriality1 != string.Empty ? sampleSelection.ConsiderMateriality1 : "");
                        ws.Cells["D13"].Value = (sampleSelection.ConsiderMateriality2 != string.Empty ? sampleSelection.ConsiderMateriality2 : "");
                        ws.Cells["E13"].Value = (sampleSelection.ConsiderMateriality3 != string.Empty ? sampleSelection.ConsiderMateriality3 : "");
                        ExcelSetBackgroundBlue(ws, "A13:B13");
                        ExcelSetFontColorRed(ws, "A13:B13");
                        ExcelSetFontColorGreen(ws, "C13:E13");
                        ExcelSetBackgroundGreen(ws, "C13:E13");
                        ExcelSetBorder(ws, "A13:E13");
                        ExcelSetAlignCenter(ws, "A13:E13");
                        ExcelSetArialSize12(ws, "A13:E13");
                        ExcelSetFontBold(ws, "A13:E13");
                        ExcelWrapText(ws, "A13:E13");
                    }

                    #endregion

                    #region Testing Round


                    if (sampleSelection.Version != "3")
                    {
                        //Generate Daily/Weekly/Monthly Test Round
                        ExcelDailyWeeklyMonty(ws, sampleSelection);
                    }
                    else if (sampleSelection.Version == "3")
                    {
                        //Generate Transactional Test Round
                        ExcelTransactional(ws, sampleSelection);
                    }


                    #endregion

                    string startupPath = Environment.CurrentDirectory;
                    //string strSourceDownload = startupPath + "\\include\\sampleselection\\download\\";
                    string strSourceDownload = Path.Combine(startupPath, "include", "sampleselection", "download");

                    if (!Directory.Exists(strSourceDownload))
                    {
                        Directory.CreateDirectory(strSourceDownload);
                    }
                    var ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    string filename = $"Draft-{sampleSelection.ExternalAuditor}-TestRound-{ts}.xlsx";
                    string strOutput = Path.Combine(strSourceDownload, filename);

                    //Check if file not exists
                    if (System.IO.File.Exists(strOutput))
                    {
                        System.IO.File.Delete(strOutput);
                    }

                    xls.SaveAs(new FileInfo(strOutput));
                    excelFilename = filename;
                }

            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error CreateExcel {ex}", "ErrorCreateExcel");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "CreateExcel");
            }

            return excelFilename;
        }

        [HttpPost("excel/test")]
        public SampleSelection CreateExcelTest([FromBody] SampleSelection sampleSelection)
        {
            return sampleSelection;
        }

        [HttpPost("podio/create/testround")]
        public async Task<IActionResult> CreatePodioTestRoundAsync([FromBody] SampleSelection sampleSelection)
        {
            List<TestRound> listTestRound = new List<TestRound>();
            bool status = false;
            try
            {
                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("SampleSelectionPodioApp").GetSection("RoundAppId").Value;
                PodioAppKey.AppToken = _config.GetSection("SampleSelectionPodioApp").GetSection("RoundAppToken").Value;
                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                if (podio.IsAuthenticated() && sampleSelection.ListTestRound.Count != 0)
                {
                    //Create testing rounds podio items
                    //int count = 1;
                    foreach (var item in sampleSelection.ListTestRound)
                    {

                        Item testingRoundItem = new Item();
                        var testingRound = testingRoundItem.Field<TextItemField>("title");
                        testingRound.Value = item.TestingRound;

                        if (item.A2Q2Samples != null)
                        {
                            var a2q2Samples = testingRoundItem.Field<TextItemField>("a2q2-samples");
                            a2q2Samples.Value = item.A2Q2Samples;
                        }

                        if (item.Date != null)
                        {
                            //date field for Population by End Date Round 3
                            var dtStart = testingRoundItem.Field<DateItemField>("date");
                            dtStart.Start = item.Date.Value.DateTime;

                        }
                        if (item.Comment != null)
                        {
                            var comment = testingRoundItem.Field<TextItemField>("a2q2-comments");
                            comment.Value = item.Comment;
                        }

                        if (item.Status != null)
                        {
                            var catStatus = testingRoundItem.Field<CategoryItemField>("status");
                            catStatus.OptionText = item.Status;
                        }


                        #region transactional and materiality


                        if (item.HeaderRoundDisplay1 != null)
                        {
                            var testingTMHeader1 = testingRoundItem.Field<TextItemField>("materiality-value-1");
                            testingTMHeader1.Value = item.HeaderRoundDisplay1;
                        }

                        if (item.HeaderRoundDisplay2 != null)
                        {
                            var testingTMHeader2 = testingRoundItem.Field<TextItemField>("materiality-value-2");
                            testingTMHeader2.Value = item.HeaderRoundDisplay2;
                        }

                        if (item.HeaderRoundDisplay3 != null)
                        {
                            var testingTMHeader3 = testingRoundItem.Field<TextItemField>("materiality-value-3");
                            testingTMHeader3.Value = item.HeaderRoundDisplay3;
                        }

                        if (item.HeaderRoundDisplay4 != null)
                        {
                            var testingTMHeader4 = testingRoundItem.Field<TextItemField>("header-display-4");
                            testingTMHeader4.Value = item.HeaderRoundDisplay4;
                        }

                        if (item.HeaderRoundDisplay5 != null)
                        {
                            var testingTMHeader5 = testingRoundItem.Field<TextItemField>("header-display-5");
                            testingTMHeader5.Value = item.HeaderRoundDisplay5;
                        }
                        if (item.HeaderRoundDisplay6 != null)
                        {
                            var testingTMHeader6 = testingRoundItem.Field<TextItemField>("header-display-6");
                            testingTMHeader6.Value = item.HeaderRoundDisplay6;
                        }
                        if (item.HeaderRoundDisplay7 != null)
                        {
                            var testingTMHeader7 = testingRoundItem.Field<TextItemField>("header-display-7");
                            testingTMHeader7.Value = item.HeaderRoundDisplay7;
                        }
                        if (item.HeaderRoundDisplay8 != null)
                        {
                            var testingTMHeader8 = testingRoundItem.Field<TextItemField>("header-display-8");
                            testingTMHeader8.Value = item.HeaderRoundDisplay8;
                        }
                        if (item.HeaderRoundDisplay9 != null)
                        {
                            var testingTMHeader9 = testingRoundItem.Field<TextItemField>("header-display-9");
                            testingTMHeader9.Value = item.HeaderRoundDisplay9;
                        }
                        if (item.HeaderRoundDisplay10 != null)
                        {
                            var testingTMHeader10 = testingRoundItem.Field<TextItemField>("header-display-10");
                            testingTMHeader10.Value = item.HeaderRoundDisplay10;
                        }

                        if (item.HeaderRoundDisplay11 != null)
                        {
                            var testingTMHeader11 = testingRoundItem.Field<TextItemField>("header-display-11");
                            testingTMHeader11.Value = item.HeaderRoundDisplay11;
                        }

                        if (item.HeaderRoundDisplay12 != null)
                        {
                            var testingTMHeader12 = testingRoundItem.Field<TextItemField>("header-display-12");
                            testingTMHeader12.Value = item.HeaderRoundDisplay12;
                        }
                        if (item.HeaderRoundDisplay13 != null)
                        {
                            var testingTMHeader13 = testingRoundItem.Field<TextItemField>("header-display-13");
                            testingTMHeader13.Value = item.HeaderRoundDisplay13;
                        }
                        if (item.HeaderRoundDisplay14 != null)
                        {
                            var testingTMHeader14 = testingRoundItem.Field<TextItemField>("header-display-14");
                            testingTMHeader14.Value = item.HeaderRoundDisplay14;
                        }
                        if (item.HeaderRoundDisplay15 != null)
                        {
                            var testingTMHeader15 = testingRoundItem.Field<TextItemField>("header-display-15");
                            testingTMHeader15.Value = item.HeaderRoundDisplay15;
                        }
                        if (item.HeaderRoundDisplay16 != null)
                        {
                            var testingTMHeader16 = testingRoundItem.Field<TextItemField>("header-display-16");
                            testingTMHeader16.Value = item.HeaderRoundDisplay16;
                        }
                        if (item.HeaderRoundDisplay17 != null)
                        {
                            var testingTMHeader17 = testingRoundItem.Field<TextItemField>("header-display-17");
                            testingTMHeader17.Value = item.HeaderRoundDisplay17;
                        }
                        if (item.HeaderRoundDisplay18 != null)
                        {
                            var testingTMHeader18 = testingRoundItem.Field<TextItemField>("header-display-18");
                            testingTMHeader18.Value = item.HeaderRoundDisplay18;
                        }
                        if (item.HeaderRoundDisplay19 != null)
                        {
                            var testingTMHeader19 = testingRoundItem.Field<TextItemField>("header-display-19");
                            testingTMHeader19.Value = item.HeaderRoundDisplay19;
                        }
                        if (item.HeaderRoundDisplay20 != null)
                        {
                            var testingTMHeader20 = testingRoundItem.Field<TextItemField>("header-display-20");
                            testingTMHeader20.Value = item.HeaderRoundDisplay20;
                        }


                        if (item.ContentDisplay1 != null)
                        {
                            var testingTMValue1 = testingRoundItem.Field<TextItemField>("content-display-1");
                            testingTMValue1.Value = item.ContentDisplay1;
                        }

                        if (item.ContentDisplay2 != null)
                        {
                            var testingTMValue1 = testingRoundItem.Field<TextItemField>("content-display-2");
                            testingTMValue1.Value = item.ContentDisplay2;
                        }

                        if (item.ContentDisplay3 != null)
                        {
                            var testingTMValue1 = testingRoundItem.Field<TextItemField>("content-display-3");
                            testingTMValue1.Value = item.ContentDisplay3;
                        }

                        if (item.ContentDisplay4 != null)
                        {
                            var testingTMValue1 = testingRoundItem.Field<TextItemField>("content-display-4");
                            testingTMValue1.Value = item.ContentDisplay4;
                        }

                        if (item.ContentDisplay5 != null)
                        {
                            var testingTMValue1 = testingRoundItem.Field<TextItemField>("content-display-5");
                            testingTMValue1.Value = item.ContentDisplay5;
                        }
                        if (item.ContentDisplay6 != null)
                        {
                            var testingTMValue6 = testingRoundItem.Field<TextItemField>("content-display-6");
                            testingTMValue6.Value = item.ContentDisplay6;
                        }
                        if (item.ContentDisplay7 != null)
                        {
                            var testingTMValue7 = testingRoundItem.Field<TextItemField>("content-display-7");
                            testingTMValue7.Value = item.ContentDisplay7;
                        }
                        if (item.ContentDisplay8 != null)
                        {
                            var testingTMValue8 = testingRoundItem.Field<TextItemField>("content-display-8");
                            testingTMValue8.Value = item.ContentDisplay8;
                        }
                        if (item.ContentDisplay9 != null)
                        {
                            var testingTMValue9 = testingRoundItem.Field<TextItemField>("content-display-9");
                            testingTMValue9.Value = item.ContentDisplay9;
                        }
                        if (item.ContentDisplay10 != null)
                        {
                            var testingTMValue10 = testingRoundItem.Field<TextItemField>("content-display-10");
                            testingTMValue10.Value = item.ContentDisplay10;
                        }
                        if (item.ContentDisplay11 != null)
                        {
                            var testingTMValue11 = testingRoundItem.Field<TextItemField>("content-display-11");
                            testingTMValue11.Value = item.ContentDisplay11;
                        }
                        if (item.ContentDisplay12 != null)
                        {
                            var testingTMValue12 = testingRoundItem.Field<TextItemField>("content-display-12");
                            testingTMValue12.Value = item.ContentDisplay12;
                        }
                        if (item.ContentDisplay13 != null)
                        {
                            var testingTMValue13 = testingRoundItem.Field<TextItemField>("content-display-13");
                            testingTMValue13.Value = item.ContentDisplay13;
                        }
                        if (item.ContentDisplay14 != null)
                        {
                            var testingTMValue14 = testingRoundItem.Field<TextItemField>("content-display-14");
                            testingTMValue14.Value = item.ContentDisplay14;
                        }
                        if (item.ContentDisplay15 != null)
                        {
                            var testingTMValue15 = testingRoundItem.Field<TextItemField>("content-display-15");
                            testingTMValue15.Value = item.ContentDisplay15;
                        }
                        if (item.ContentDisplay16 != null)
                        {
                            var testingTMValue16 = testingRoundItem.Field<TextItemField>("content-display-16");
                            testingTMValue16.Value = item.ContentDisplay16;
                        }
                        if (item.ContentDisplay17 != null)
                        {
                            var testingTMValue17 = testingRoundItem.Field<TextItemField>("content-display-17");
                            testingTMValue17.Value = item.ContentDisplay17;
                        }
                        if (item.ContentDisplay18 != null)
                        {
                            var testingTMValue18 = testingRoundItem.Field<TextItemField>("content-display-18");
                            testingTMValue18.Value = item.ContentDisplay18;
                        }
                        if (item.ContentDisplay19 != null)
                        {
                            var testingTMValue19 = testingRoundItem.Field<TextItemField>("content-display-19");
                            testingTMValue19.Value = item.ContentDisplay19;
                        }
                        if (item.ContentDisplay20 != null)
                        {
                            var testingTMValue20 = testingRoundItem.Field<TextItemField>("content-display-20");
                            testingTMValue20.Value = item.ContentDisplay20;
                        }


                        #endregion


                        var roundId = await podio.ItemService.AddNewItem(Int32.Parse(PodioAppKey.AppId), testingRoundItem);
                        TestRound createdItemId = new TestRound();
                        createdItemId = item;
                        createdItemId.PodioItemId = int.Parse(roundId.ToString());
                        listTestRound.Add(createdItemId);
                    }

                    status = true;
                }

            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex.ToString());
                //ErrorLog.Write(ex);
                FileLog.Write($"Error CreatePodioTestRoundAsync {ex}", "ErrorCreatePodioTestRoundAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "CreatePodioTestRoundAsync");
            }
            if (status)
            {
                return Ok(listTestRound.ToArray());
            }
            else
            {
                return NoContent();
            }



        }

        [HttpPost("podio/create/sampleselection")]
        public async Task<IActionResult> CreatePodioSampleSelectionAsync([FromBody] SampleSelection sampleSelection)
        {
            List<string> itemId = new List<string>();
            bool status = false;
            try
            {
                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("SampleSelectionPodioApp").GetSection("AppId").Value;
                PodioAppKey.AppToken = _config.GetSection("SampleSelectionPodioApp").GetSection("AppToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                if (podio.IsAuthenticated() && sampleSelection.ListTestRound.Count != 0)
                {
                    //Create sample selection
                    Item sampleSelectionItem = new Item();

                    if (sampleSelection.ClientId != null)
                    {
                        var clientField = sampleSelectionItem.Field<AppItemField>("client");
                        //Item id's to reference
                        clientField.ItemIds = new List<int> { sampleSelection.ClientId.Value };
                    }

                    if (sampleSelection.Q4R3SampleRequired != null && sampleSelection.Q4R3SampleRequired != string.Empty)
                    {
                        var catSampReq = sampleSelectionItem.Field<CategoryItemField>("q4-r3-sample-required");
                        catSampReq.OptionText = sampleSelection.Q4R3SampleRequired;
                    }

                    if (sampleSelection.CountSampleQ4R3 != null)
                    {
                        var txtCountTestQ4 = sampleSelectionItem.Field<TextItemField>("how-many-samples-to-be-tested-in-q4");
                        txtCountTestQ4.Value = sampleSelection.CountSampleQ4R3.Value.ToString();
                    }

                    if (sampleSelection.Risk != null && sampleSelection.Risk != string.Empty)
                    {
                        var catRisk = sampleSelectionItem.Field<CategoryItemField>("risk");
                        catRisk.OptionText = sampleSelection.Risk;
                    }

                    if (sampleSelection.AnnualPopulation != null)
                    {
                        var txtAnnualPop = sampleSelectionItem.Field<TextItemField>("annual-population");
                        txtAnnualPop.Value = sampleSelection.AnnualPopulation.ToString();
                    }

                    if (sampleSelection.AnnualSampleSize != null)
                    {
                        var txtAnnualSize = sampleSelectionItem.Field<TextItemField>("annual-sample-size");
                        txtAnnualSize.Value = sampleSelection.AnnualSampleSize.ToString();
                    }

                    if (sampleSelection.Frequency != null && sampleSelection.Frequency != string.Empty)
                    {
                        var catFrequency = sampleSelectionItem.Field<CategoryItemField>("frequency");
                        catFrequency.OptionText = sampleSelection.Frequency;
                    }

                    if (sampleSelection.Round1Start != null)
                    {
                        //date field for Population by Start Date Round 1
                        var dtPopStartRound1 = sampleSelectionItem.Field<DateItemField>("round-1");
                        dtPopStartRound1.Start = DateTime.Parse(sampleSelection.Round1Start.Value.ToString("yyyy-MM-dd"));
                    }

                    if (sampleSelection.Round2Start != null)
                    {
                        //date field for Population by Start Date Round 2
                        var dtPopStartRound2 = sampleSelectionItem.Field<DateItemField>("round-1-2");
                        dtPopStartRound2.Start = DateTime.Parse(sampleSelection.Round2Start.Value.ToString("yyyy-MM-dd"));
                    }

                    if (sampleSelection.Round3Start != null)
                    {
                        //date field for Population by Start Date Round 3
                        var dtPopStartRound3 = sampleSelectionItem.Field<DateItemField>("population-by-start-date-round-3");
                        dtPopStartRound3.Start = DateTime.Parse(sampleSelection.Round3Start.Value.ToString("yyyy-MM-dd"));
                    }

                    if (sampleSelection.Round1End != null)
                    {
                        //date field for Population by End Date Round 1
                        var dtPopEndRound1 = sampleSelectionItem.Field<DateItemField>("round-2");
                        dtPopEndRound1.Start = DateTime.Parse(sampleSelection.Round1End.Value.ToString("yyyy-MM-dd"));
                    }

                    if (sampleSelection.Round2End != null)
                    {
                        //date field for Population by End Date Round 2
                        var dtPopEndRound2 = sampleSelectionItem.Field<DateItemField>("round-2-2");
                        dtPopEndRound2.Start = DateTime.Parse(sampleSelection.Round2End.Value.ToString("yyyy-MM-dd"));
                    }

                    if (sampleSelection.Round3End != null)
                    {
                        //date field for Population by End Date Round 3
                        var dtPopEndRound3 = sampleSelectionItem.Field<DateItemField>("population-by-end-date-round-3");
                        dtPopEndRound3.Start = DateTime.Parse(sampleSelection.Round3End.Value.ToString("yyyy-MM-dd"));
                    }

                    NumericItemField popByRound1 = sampleSelectionItem.Field<NumericItemField>("population-by-round-1");
                    popByRound1.Value = sampleSelection.PopulationByRound1;

                    NumericItemField popByRound2 = sampleSelectionItem.Field<NumericItemField>("population-by-round-2");
                    popByRound2.Value = sampleSelection.PopulationByRound2;

                    NumericItemField popByRound3 = sampleSelectionItem.Field<NumericItemField>("population-by-round-3");
                    popByRound3.Value = sampleSelection.PopulationByRound3;

                    NumericItemField sampByRound1 = sampleSelectionItem.Field<NumericItemField>("samples-by-round-1");
                    sampByRound1.Value = sampleSelection.SamplesByRound1;

                    NumericItemField sampByRound2 = sampleSelectionItem.Field<NumericItemField>("samples-by-round-2");
                    sampByRound2.Value = sampleSelection.SamplesByRound2;

                    NumericItemField sampByRound3 = sampleSelectionItem.Field<NumericItemField>("samples-by-round-3");
                    sampByRound3.Value = sampleSelection.SamplesByRound3;

                    NumericItemField sampClosedByRound1 = sampleSelectionItem.Field<NumericItemField>("samples-closed-round-1");
                    sampClosedByRound1.Value = sampleSelection.SamplesCloseRound1;

                    NumericItemField sampClosedByRound2 = sampleSelectionItem.Field<NumericItemField>("samples-closed-round-2");
                    sampClosedByRound2.Value = sampleSelection.SamplesCloseRound2;

                    NumericItemField sampClosedByRound3 = sampleSelectionItem.Field<NumericItemField>("samples-closed-round-3");
                    sampClosedByRound3.Value = sampleSelection.SamplesCloseRound3;

                    NumericItemField sampRemaining1 = sampleSelectionItem.Field<NumericItemField>("samples-remaining-round-1");
                    sampRemaining1.Value = sampleSelection.SamplesRemainingRound1;

                    NumericItemField sampRemaining2 = sampleSelectionItem.Field<NumericItemField>("samples-remaining-round-2");
                    sampRemaining2.Value = sampleSelection.SamplesRemainingRound2;

                    NumericItemField sampRemaining3 = sampleSelectionItem.Field<NumericItemField>("samples-remaining-round-3");
                    sampRemaining3.Value = sampleSelection.SamplesRemainingRound3;

                    if (sampleSelection.PopulationFile != null && sampleSelection.PopulationFile != string.Empty)
                    {
                        var txtPopulationFilename = sampleSelectionItem.Field<TextItemField>("population-guid-name");
                        txtPopulationFilename.Value = sampleSelection.PopulationFile;
                    }

                    if (sampleSelection.IsMateriality != null && sampleSelection.IsMateriality != string.Empty)
                    {
                        var catMateriality = sampleSelectionItem.Field<CategoryItemField>("is-materiality-required");
                        catMateriality.OptionText = sampleSelection.IsMateriality;
                    }

                    if (sampleSelection.ConsiderMateriality1 != null && sampleSelection.ConsiderMateriality1 != string.Empty)
                    {
                        var txtconsiderMat1 = sampleSelectionItem.Field<TextItemField>("materiality-to-consider-round-1");
                        txtconsiderMat1.Value = sampleSelection.ConsiderMateriality1;
                    }

                    if (sampleSelection.ConsiderMateriality2 != null && sampleSelection.ConsiderMateriality2 != string.Empty)
                    {
                        var txtconsiderMat2 = sampleSelectionItem.Field<TextItemField>("materiality-to-consider-round-2");
                        txtconsiderMat2.Value = sampleSelection.ConsiderMateriality2;
                    }

                    if (sampleSelection.ConsiderMateriality3 != null && sampleSelection.ConsiderMateriality3 != string.Empty)
                    {
                        var txtconsiderMat3 = sampleSelectionItem.Field<TextItemField>("materiality-to-consider-round-3");
                        txtconsiderMat3.Value = sampleSelection.ConsiderMateriality3;
                    }

                    if (sampleSelection != null)
                    {
                        var txtSampleSelectionData = sampleSelectionItem.Field<TextItemField>("json-data");
                        txtSampleSelectionData.Value = JsonConvert.SerializeObject(sampleSelection).ToString();
                    }

                    if (sampleSelection.Version != null && sampleSelection.Version != string.Empty)
                    {
                        var txtVersion = sampleSelectionItem.Field<TextItemField>("version");
                        txtVersion.Value = sampleSelection.Version;
                    }

                    if (sampleSelection.RcmPodioItemId.ToString() != "0")
                    {
                        var txtRcmItemId = sampleSelectionItem.Field<TextItemField>("rcm-item-id");
                        txtRcmItemId.Value = sampleSelection.RcmPodioItemId.ToString();
                    }


                    if (sampleSelection.ListTestRound.Count > 0)
                    {
                        var roundTest = sampleSelectionItem.Field<AppItemField>("round-test");
                        List<int> itemIdTestRound = new List<int>();
                        foreach (var item in sampleSelection.ListTestRound)
                        {
                            itemIdTestRound.Add(item.PodioItemId);
                            Debug.WriteLine(item.PodioItemId);
                        }
                        //Item id's to reference
                        roundTest.ItemIds = itemIdTestRound;

                    }

                    var podioReturn = await podio.ItemService.AddNewItem(Int32.Parse(PodioAppKey.AppId), sampleSelectionItem);
                    itemId.Add(podioReturn.ToString());
                    status = true;
                }

            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex.ToString());
                //ErrorLog.Write(ex);
                FileLog.Write($"Error CreatePodioSampleSelectionAsync {ex}", "ErrorCreatePodioSampleSelectionAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "CreatePodioSampleSelectionAsync");
            }
            if (status)
            {
                return (Ok(itemId));
            }
            else
            {
                return NoContent();
            }
        }

        [HttpPost("data/create")]
        public async Task<IActionResult> SaveSampleSelectionAsync([FromBody] SampleSelection sampleSelection)
        {
            int result = 0;
            try
            {
                SampleSelectionService podioService = new SampleSelectionService(_soxContext, _config);
                result = await podioService.SaveSampleSelectionAsync(sampleSelection);
            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex.ToString());
                //ErrorLog.Write(ex);
                FileLog.Write($"Error SaveSampleSelectionAsync {ex}", "ErrorSaveSampleSelectionAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SaveSampleSelectionAsync");
            }

            if (result > 0)
            {
                return Ok(new { ItemId = result });
            }
            else
            {
                return NoContent();
            }

        }

        //Dictionary use for download files
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
                {".png", "image/png"},
                {".jpg", "image/jpeg"},
            };
        }

        #region Excel Functions

        private List<Population> ExcelPopulation(ExcelWorksheet ws, int type)
        {
            List<Population> listPopulation = new List<Population>();
            int colCount = ws.Dimension.End.Column;
            int rowCount = ws.Dimension.End.Row;
            //int type = 1;
            for (int row = 1; row <= rowCount; row++)
            {
                //if (ws.Cells[row, 2].Value != null){}

                Population population = new Population();
                for (int col = 1; col <= colCount; col++)
                {

                    population.PopType = type;

                    switch (col)
                    {

                        case 1:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.UniqueId = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 2:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.PurchaseOrder = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 3:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.SupplierSched = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 4:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.PoRev = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 5:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.PoLine = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 6:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.Requisition = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 7:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.RequisitionLine = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 8:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.EnteredBy = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 9:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.Status = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 10:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.Buyer = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 11:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.Contact = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 12:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.OrderDate = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 13:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.Supplier = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 14:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.ShipTo = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 15:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.SortName = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 16:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.Telephone = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 17:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.ItemNumber = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 18:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.ProdLine = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 19:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.ProdDescription = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 20:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.Site = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 21:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.Location = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 22:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.ItemRevision = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 23:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.SupplierItem = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 24:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.QuantityOrdered = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 25:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.UnitOfMeasure = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 26:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.UMConversion = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 27:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.QtyOrderedXPOCost = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 28:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.QuantityReceived = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 29:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.QtyOpen = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 30:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.QtyReturned = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 31:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.DueDate = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 32:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.OverDue = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 33:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.PerformanceDate = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 34:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.Currency = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 35:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.StandardCost = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 36:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.PurchasedCost = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 37:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.PurCostBC = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 38:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.OpenPoCost = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 39:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.PpvPerUnit = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 40:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.Type = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 41:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.StdMtlCostNow = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 42:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.WorkOrderId = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 43:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.Operation = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 44:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.PurchAcct = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 45:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.GlAccountDesc = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 46:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.CostCenter = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 47:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.GlDescription = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 48:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.Project = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 49:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.Description = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 50:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.Taxable = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 51:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.Comments = ws.Cells[row, col].Value?.ToString();
                            }
                            break;

                    }
                }

                if (population != null)
                {
                    listPopulation.Add(population);
                }


            }

            return listPopulation;

        }

        private List<List<string>> ExcelPopulation2(ExcelWorksheet ws, int type)
        {
            int colCount = ws.Dimension.End.Column;
            int rowCount = ws.Dimension.End.Row;
            //List<string>[] listPopulation = new List<string>[];
            List<List<string>> parentPopulation = new List<List<string>>();
            bool isEmpty;
            //int type = 1;
            for (int row = 1; row <= rowCount; row++)
            {
                isEmpty = true;
                List<string> childPopulation = new List<string>();
                for (int col = 1; col <= colCount; col++)
                {
                    bool isDateHeader = IsDateHeader(ws.Cells[1, col].Value?.ToString());
                    bool isDateTime = IsDateTime(ws.Cells[row, col].Value?.ToString());
                    bool isDecimal = IsDecimal(ws.Cells[row, col].Value?.ToString());
                    bool isIdentificationHeader = IsIdentificationHeader(ws.Cells[1, col].Value?.ToString());

                    Debug.WriteLine($"Cell[{row},{col}] : {ws.Cells[row, col].Value?.ToString()} - IdentificationHeader({isIdentificationHeader}) | DateHeader({isDateHeader}) | Datetime({isDateTime}) | Decimal({isDecimal})");
                    //Debug.WriteLine($"Cell[{row},{col}] : {ws.Cells[row, col].Value?.ToString()}");

                    if (isDateTime && isDateHeader)
                    {
                        DateTime dtVal = DateTime.Parse(ws.Cells[row, col].Value?.ToString());
                        childPopulation.Add(dtVal.ToString("MM/dd/yyyy"));
                    }
                    else if (isDecimal && !isIdentificationHeader && !isDateHeader)
                    {
                        decimal decVal = decimal.Parse(ws.Cells[row, col].Value?.ToString());
                        childPopulation.Add(decVal.ToString("0.00"));
                    }
                    else
                    {
                        childPopulation.Add(ws.Cells[row, col].Value?.ToString());
                    }

                    if (ws.Cells[row, col].Value?.ToString() != string.Empty && ws.Cells[row, col].Value?.ToString() != null)
                    {
                        isEmpty = false;
                    }
                }

                if (childPopulation != null && !isEmpty)
                {
                    parentPopulation.Add(childPopulation);
                }

            }

            return parentPopulation;

        }

        //Excel function to set boarder range
        private void ExcelSetBorder(ExcelWorksheet ws, string range)
        {
            ws.Cells[range].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells[range].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells[range].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            ws.Cells[range].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
        }

        //Excel function to set align center
        private void ExcelSetAlignCenter(ExcelWorksheet ws, string range)
        {
            ws.Cells[range].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells[range].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        }

        //Excel function to set arial 12
        private void ExcelSetArialSize12(ExcelWorksheet ws, string range)
        {
            ws.Cells[range].Style.Font.SetFromFont(new Font("Arial", 12));
        }

        //Excel function to set background #ccffcc
        private void ExcelSetBackgroundGreen(ExcelWorksheet ws, string range)
        {
            ws.Cells[range].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[range].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#ccffcc"));
        }

        //Excel function to set background #d9e1f2
        private void ExcelSetBackgroundBlue(ExcelWorksheet ws, string range)
        {
            ws.Cells[range].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[range].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#d9e1f2"));
        }

        //Excel function to set font color #c00000
        private void ExcelSetFontColorRed(ExcelWorksheet ws, string range)
        {
            ws.Cells[range].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#c00000"));
        }

        //Excel function to set font color #375623
        private void ExcelSetFontColorGreen(ExcelWorksheet ws, string range)
        {
            ws.Cells[range].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#375623"));
        }

        //Excel function to set font bold
        private void ExcelSetFontBold(ExcelWorksheet ws, string range)
        {
            ws.Cells[range].Style.Font.Bold = true;
        }

        //Excel function to set wraptext true
        private void ExcelWrapText(ExcelWorksheet ws, string range)
        {
            ws.Cells[range].Style.WrapText = true;
        }

        private void ExcelDailyWeeklyMonty(ExcelWorksheet ws, SampleSelection sampleSelection)
        {
            #region Version 1
            int row = 14;
            int colTestRounds = 1;
            int colA2Q2Samples = 2;
            int colDate = 3;
            int colWeekOnly = 4;
            int colStatus = 5;
            int colComment = 6;

            ws.Cells[row, colTestRounds].Value = "Testing Round";
            ws.Cells[row, colA2Q2Samples].Value = "A2Q2 Samples";
            ws.Cells[row, colDate].Value = "Date";
            ws.Cells[row, colWeekOnly].Value = "Weekly Only";
            ws.Cells[row, colStatus].Value = "Status";
            ws.Cells[row, colComment].Value = "A2Q2 Comments";
            ws.Cells["F" + row + ":H" + row].Merge = true;
            ws.Cells["A" + row + ":H" + row].Style.WrapText = true;

            ws.Cells["A" + row + ":H" + row].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + row + ":H" + row].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + row + ":H" + row].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + row + ":H" + row].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + row + ":H" + row].Style.Font.SetFromFont(new Font("Arial", 12));
            ws.Cells["A" + row + ":H" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A" + row + ":H" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["A" + row + ":D" + row].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A" + row + ":D" + row].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#d9e1f2"));
            ws.Cells["E" + row + ":H" + row].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["E" + row + ":H" + row].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#ccffcc"));
            ws.Cells["D" + row + ":H" + row].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#c00000"));
            ws.Cells["A" + row + ":H" + row].Style.Font.Bold = true;
            ws.Row(row).Height = 42;

            row++;
            foreach (var item in sampleSelection.ListTestRound)
            {
                ws.Cells[row, colTestRounds].Value = item.TestingRound;
                ws.Cells[row, colA2Q2Samples].Value = item.A2Q2Samples;
                if (sampleSelection.Version != "3")
                {
                    ws.Cells[row, colDate].Value = (sampleSelection.Frequency == "Monthly") ? item.MonthOnly : item.Date.Value.DateTime.ToString("MM/dd/yyyy");
                }
                else
                {
                    ws.Cells[row, colDate].Value = item.ContentDisplay1;
                }

                ws.Cells[row, colWeekOnly].Value = (sampleSelection.Frequency == "Weekly") ? item.WeeklyOnly : string.Empty;
                ws.Cells[row, colStatus].Value = item.Status;
                ws.Cells[row, colComment].Value = item.Comment;

                #region Format
                ws.Cells["F" + row + ":H" + row].Merge = true;
                ws.Cells["A" + row + ":H" + row].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + row + ":H" + row].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + row + ":H" + row].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + row + ":H" + row].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + row + ":H" + row].Style.Font.SetFromFont(new Font("Arial", 12));
                ws.Cells["A" + row + ":H" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells["A" + row + ":H" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Cells["A" + row + ":D" + row].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells["A" + row + ":D" + row].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#d9e1f2"));
                ws.Cells["E" + row + ":H" + row].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells["E" + row + ":H" + row].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#ccffcc"));
                #endregion

                row++;
            }
            #endregion

        }

        private void ExcelTransactional(ExcelWorksheet ws, SampleSelection sampleSelection)
        {
            #region Round 1
            int row = 15;

            List<int> listColumn = new List<int>();
            List<TestRound> listTestRound = new List<TestRound>();
            string[] header1 = new string[] { string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty };
            int column = 1;
            listColumn.Add(column);
            column++;
            listColumn.Add(column);

            if (sampleSelection.ListTestRound != null && sampleSelection.ListTestRound.Any(item => item.TestingRound == "Round 1"))
            {
                listTestRound.AddRange(sampleSelection.ListTestRound);

                
                header1[0] = listTestRound.Where(x => x.TestingRound == "Round 1").Select(x => x.HeaderRoundDisplay1).FirstOrDefault(); 
                if (header1[0] != string.Empty && header1[0] != null)
                {
                    column++;
                    listColumn.Add(column);
                }


                
                header1[1] = listTestRound.Where(x => x.TestingRound == "Round 1").Select(x => x.HeaderRoundDisplay2).FirstOrDefault();
                if (header1[1] != string.Empty && header1[1] != null)
                {
                    column++;
                    listColumn.Add(column);
                }


                
                header1[2] = listTestRound.Where(x => x.TestingRound == "Round 1").Select(x => x.HeaderRoundDisplay3).FirstOrDefault();
                if (header1[2] != string.Empty && header1[2] != null)
                {
                    column++;
                    listColumn.Add(column);
                }


                
                header1[3] = listTestRound.Where(x => x.TestingRound == "Round 1").Select(x => x.HeaderRoundDisplay4).FirstOrDefault();
                if (header1[3] != string.Empty && header1[3] != null)
                {
                    column++;
                    listColumn.Add(column);
                }

                
                header1[4] = listTestRound.Where(x => x.TestingRound == "Round 1").Select(x => x.HeaderRoundDisplay5).FirstOrDefault();
                if (header1[4] != string.Empty && header1[4] != null)
                {
                    column++;
                    listColumn.Add(column);
                }
                
                header1[5] = listTestRound.Where(x => x.TestingRound == "Round 1").Select(x => x.HeaderRoundDisplay6).FirstOrDefault();
                if (header1[5] != string.Empty && header1[5] != null)
                {
                    column++;
                    listColumn.Add(column);
                }
                
                header1[6] = listTestRound.Where(x => x.TestingRound == "Round 1").Select(x => x.HeaderRoundDisplay7).FirstOrDefault();
                if (header1[6] != string.Empty && header1[6] != null)
                {
                    column++;
                    listColumn.Add(column);
                }
                
                header1[7] = listTestRound.Where(x => x.TestingRound == "Round 1").Select(x => x.HeaderRoundDisplay8).FirstOrDefault();
                if (header1[7] != string.Empty && header1[7] != null)
                {
                    column++;
                    listColumn.Add(column);
                }
               
                header1[8] = listTestRound.Where(x => x.TestingRound == "Round 1").Select(x => x.HeaderRoundDisplay9).FirstOrDefault();
                if (header1[8] != string.Empty && header1[8] != null)
                {
                    column++;
                    listColumn.Add(column);
                }
                
                header1[9] = listTestRound.Where(x => x.TestingRound == "Round 1").Select(x => x.HeaderRoundDisplay10).FirstOrDefault();
                if (header1[9] != string.Empty && header1[9] != null)
                {
                    column++;
                    listColumn.Add(column);
                }


                column++; //status
                listColumn.Add(column);
                column++; //for a2q2 comments
                listColumn.Add(column);

                //Console.WriteLine($"listColumn.Count: {listColumn.Count}");

                ws.Cells[row, listColumn[0]].Value = "Testing Round";
                ws.Cells[row, listColumn[1]].Value = "A2Q2 Samples";

                int columnHeaderStart = 3;
                //int countHeader = 0;
                if (listColumn.Count - 4 > 0)
                {
                    for (int i = 0; i < listColumn.Count - 4; i++)
                    {
                        if (header1[i] != string.Empty)
                        {
                            //countHeader++;
                            ws.Cells[row, columnHeaderStart].Value = header1[i];
                            columnHeaderStart++;
                        }
                    }
                }

                ws.Cells[row, listColumn[listColumn.Count - 2]].Value = "Status";
                ws.Cells[row, listColumn[listColumn.Count - 1]].Value = "A2Q2 Comments";
                ws.Cells[row, listColumn[listColumn.Count - 1], row, listColumn[listColumn.Count - 1] + 2].Merge = true;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.WrapText = true;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Font.SetFromFont(new Font("Arial", 12));
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#d9e1f2"));
                ws.Cells[row, listColumn[listColumn.Count - 2], row, listColumn[listColumn.Count - 1] + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, listColumn[listColumn.Count - 2], row, listColumn[listColumn.Count - 1] + 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#ccffcc"));
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1]].Style.Font.Bold = true;
                ws.Row(row).Height = 42;

                row++;
                foreach (var item in sampleSelection.ListTestRound)
                {
                    if (item.TestingRound == "Round 1")
                    {
                        int columnIndex = 0;
                        ws.Cells[row, listColumn[columnIndex]].Value = item.TestingRound;
                        columnIndex++;
                        ws.Cells[row, listColumn[columnIndex]].Value = item.A2Q2Samples;

                        
                        if (item.ContentDisplay1 != string.Empty && item.ContentDisplay1 != null)
                        {
                            columnIndex++;
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay1;
                        }

                        
                        if (item.ContentDisplay2 != string.Empty && item.ContentDisplay2 != null)
                        {
                            columnIndex++;
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay2;
                        }

                        
                        if (item.ContentDisplay3 != string.Empty && item.ContentDisplay3 != null)
                        {
                            columnIndex++;
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay3;
                        }

                        
                        if (item.ContentDisplay4 != string.Empty && item.ContentDisplay4 != null)
                        {
                            columnIndex++;
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay4;
                        }

                        
                        if (item.ContentDisplay5 != string.Empty && item.ContentDisplay5 != null)
                        {
                            columnIndex++;
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay5;
                        }
                        
                        if (item.ContentDisplay6 != string.Empty && item.ContentDisplay6 != null)
                        {
                            columnIndex++;
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay6;
                        }
                        
                        if (item.ContentDisplay7 != string.Empty && item.ContentDisplay7 != null)
                        {
                            columnIndex++;
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay7;
                        }
                        
                        if (item.ContentDisplay8 != string.Empty && item.ContentDisplay8 != null)
                        {
                            columnIndex++;
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay8;
                        }
                        
                        if (item.ContentDisplay9 != string.Empty && item.ContentDisplay9 != null)
                        {
                            columnIndex++;
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay9;
                        }
                        
                        if (item.ContentDisplay10 != string.Empty && item.ContentDisplay10 != null)
                        {
                            columnIndex++;
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay10;
                        }

                        columnIndex++;
                        ws.Cells[row, columnIndex + 1].Value = item.Status;
                        columnIndex++;
                        ws.Cells[row, columnIndex + 1].Value = item.Comment;

                        #region Format
                        ws.Cells[row, listColumn[listColumn.Count - 1], row, listColumn[listColumn.Count - 1] + 2].Merge = true;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Font.SetFromFont(new Font("Arial", 12));
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#d9e1f2"));
                        ws.Cells[row, listColumn[listColumn.Count - 2], row, listColumn[listColumn.Count - 1] + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[row, listColumn[listColumn.Count - 2], row, listColumn[listColumn.Count - 1] + 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#ccffcc"));
                        #endregion

                        row++;
                    }

                }

            }





            #endregion

            #region Round 2
            row += 2;
            listColumn = new List<int>();
            listTestRound = new List<TestRound>();
            header1 = new string[] { string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty };
            column = 1;
            listColumn.Add(column);
            column++;
            listColumn.Add(column);

            if (sampleSelection.ListTestRound != null && sampleSelection.ListTestRound.Any(item => item.TestingRound == "Round 2"))
            {
                listTestRound.AddRange(sampleSelection.ListTestRound);

                header1[0] = listTestRound.Where(x => x.TestingRound == "Round 2").Select(x => x.HeaderRoundDisplay1).FirstOrDefault();
                
                if (header1[0] != string.Empty && header1[0] != null)
                {
                    column++;
                    listColumn.Add(column);
                }

                
                header1[1] = listTestRound.Where(x => x.TestingRound == "Round 2").Select(x => x.HeaderRoundDisplay2).FirstOrDefault();
                
                if (header1[1] != string.Empty && header1[1] != null)
                {
                    column++;
                    listColumn.Add(column);
                }


                header1[2] = listTestRound.Where(x => x.TestingRound == "Round 2").Select(x => x.HeaderRoundDisplay3).FirstOrDefault();
                
                if (header1[2] != string.Empty && header1[2] != null)
                {
                    column++;
                    listColumn.Add(column);
                }


                header1[3] = listTestRound.Where(x => x.TestingRound == "Round 2").Select(x => x.HeaderRoundDisplay4).FirstOrDefault();
                
                if (header1[3] != string.Empty && header1[3] != null)
                {
                    column++;
                    listColumn.Add(column);
                }


                header1[4] = listTestRound.Where(x => x.TestingRound == "Round 2").Select(x => x.HeaderRoundDisplay5).FirstOrDefault();
                
                if (header1[4] != string.Empty && header1[4] != null)
                {
                    column++;
                    listColumn.Add(column);
                }

                header1[5] = listTestRound.Where(x => x.TestingRound == "Round 2").Select(x => x.HeaderRoundDisplay6).FirstOrDefault();

                if (header1[5] != string.Empty && header1[5] != null)
                {
                    column++;
                    listColumn.Add(column);
                }

                header1[6] = listTestRound.Where(x => x.TestingRound == "Round 2").Select(x => x.HeaderRoundDisplay7).FirstOrDefault();

                if (header1[6] != string.Empty && header1[6] != null)
                {
                    column++;
                    listColumn.Add(column);
                }

                header1[7] = listTestRound.Where(x => x.TestingRound == "Round 2").Select(x => x.HeaderRoundDisplay8).FirstOrDefault();

                if (header1[7] != string.Empty && header1[7] != null)
                {
                    column++;
                    listColumn.Add(column);
                }

                header1[8] = listTestRound.Where(x => x.TestingRound == "Round 2").Select(x => x.HeaderRoundDisplay9).FirstOrDefault();

                if (header1[8] != string.Empty && header1[8] != null)
                {
                    column++;
                    listColumn.Add(column);
                }

                header1[9] = listTestRound.Where(x => x.TestingRound == "Round 2").Select(x => x.HeaderRoundDisplay10).FirstOrDefault();

                if (header1[9] != string.Empty && header1[9] != null)
                {
                    column++;
                    listColumn.Add(column);
                }

                column++; //status
                listColumn.Add(column);
                column++; //for a2q2 comments
                listColumn.Add(column);

                //Console.WriteLine($"listColumn.Count: {listColumn.Count}");

                ws.Cells[row, listColumn[0]].Value = "Testing Round";
                ws.Cells[row, listColumn[1]].Value = "A2Q2 Samples";

                int columnHeaderStart = 3;
                //int countHeader = 0;
                if (listColumn.Count - 4 > 0)
                {
                    for (int i = 0; i < listColumn.Count - 4; i++)
                    {
                        if (header1[i] != string.Empty)
                        {
                            //countHeader++;
                            ws.Cells[row, columnHeaderStart].Value = header1[i];
                            columnHeaderStart++;
                        }
                    }
                }

                ws.Cells[row, listColumn[listColumn.Count - 2]].Value = "Status";
                ws.Cells[row, listColumn[listColumn.Count - 1]].Value = "A2Q2 Comments";
                ws.Cells[row, listColumn[listColumn.Count - 1], row, listColumn[listColumn.Count - 1] + 2].Merge = true;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.WrapText = true;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Font.SetFromFont(new Font("Arial", 12));
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#d9e1f2"));
                ws.Cells[row, listColumn[listColumn.Count - 2], row, listColumn[listColumn.Count - 1] + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, listColumn[listColumn.Count - 2], row, listColumn[listColumn.Count - 1] + 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#ccffcc"));
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1]].Style.Font.Bold = true;
                ws.Row(row).Height = 42;

                row++;
                foreach (var item in sampleSelection.ListTestRound)
                {
                    if (item.TestingRound == "Round 2")
                    {
                        int columnIndex = 0;
                        ws.Cells[row, listColumn[columnIndex]].Value = item.TestingRound;
                        columnIndex++;
                        ws.Cells[row, listColumn[columnIndex]].Value = item.A2Q2Samples;

                        
                        if (item.ContentDisplay1 != string.Empty && item.ContentDisplay1 != null)
                        {
                            columnIndex++;
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay1;
                        }

                        
                        if (item.ContentDisplay2 != string.Empty && item.ContentDisplay2 != null)
                        {
                            columnIndex++;
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay2;
                        }

                        
                        if (item.ContentDisplay3 != string.Empty && item.ContentDisplay3 != null)
                        {
                            columnIndex++;
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay3;
                        }

                        
                        if (item.ContentDisplay4 != string.Empty && item.ContentDisplay4 != null)
                        {
                            columnIndex++;
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay4;
                        }

                        
                        if (item.ContentDisplay5 != string.Empty && item.ContentDisplay5 != null)
                        {
                            columnIndex++;
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay5;
                        }

                        if (item.ContentDisplay6 != string.Empty && item.ContentDisplay6 != null)
                        {
                            columnIndex++;
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay6;
                        }

                        if (item.ContentDisplay7 != string.Empty && item.ContentDisplay7 != null)
                        {
                            columnIndex++;
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay7;
                        }

                        if (item.ContentDisplay8 != string.Empty && item.ContentDisplay8 != null)
                        {
                            columnIndex++;
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay8;
                        }

                        if (item.ContentDisplay9 != string.Empty && item.ContentDisplay9 != null)
                        {
                            columnIndex++;
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay9;
                        }

                        if (item.ContentDisplay10 != string.Empty && item.ContentDisplay10 != null)
                        {
                            columnIndex++;
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay10;
                        }

                        columnIndex++;
                        ws.Cells[row, columnIndex + 1].Value = item.Status;
                        columnIndex++;
                        ws.Cells[row, columnIndex + 1].Value = item.Comment;

                        #region Format
                        ws.Cells[row, listColumn[listColumn.Count - 1], row, listColumn[listColumn.Count - 1] + 2].Merge = true;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Font.SetFromFont(new Font("Arial", 12));
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#d9e1f2"));
                        ws.Cells[row, listColumn[listColumn.Count - 2], row, listColumn[listColumn.Count - 1] + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[row, listColumn[listColumn.Count - 2], row, listColumn[listColumn.Count - 1] + 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#ccffcc"));
                        #endregion

                        row++;
                    }

                }
            }

            #endregion

            #region Round 3
            row += 2;
            listColumn = new List<int>();
            listTestRound = new List<TestRound>();
            header1 = new string[] { string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty };
            column = 1;
            listColumn.Add(column);
            column++;
            listColumn.Add(column);

            if (sampleSelection.ListTestRound != null && sampleSelection.ListTestRound.Any(item => item.TestingRound == "Round 3"))
            {
                listTestRound.AddRange(sampleSelection.ListTestRound);

                header1[0] = listTestRound.Where(x => x.TestingRound == "Round 3").Select(x => x.HeaderRoundDisplay1).FirstOrDefault();
                
                if (header1[0] != string.Empty && header1[0] != null)
                {
                    column++;
                    listColumn.Add(column);
                }


                header1[1] = listTestRound.Where(x => x.TestingRound == "Round 3").Select(x => x.HeaderRoundDisplay2).FirstOrDefault();
                
                if (header1[1] != string.Empty && header1[1] != null)
                {
                    column++;
                    listColumn.Add(column);
                }


                header1[2] = listTestRound.Where(x => x.TestingRound == "Round 3").Select(x => x.HeaderRoundDisplay3).FirstOrDefault();
                
                if (header1[2] != string.Empty && header1[2] != null)
                {
                    column++;
                    listColumn.Add(column);
                }


                header1[3] = listTestRound.Where(x => x.TestingRound == "Round 3").Select(x => x.HeaderRoundDisplay4).FirstOrDefault();
                
                if (header1[3] != string.Empty && header1[3] != null)
                {
                    column++;
                    listColumn.Add(column);
                }


                header1[4] = listTestRound.Where(x => x.TestingRound == "Round 3").Select(x => x.HeaderRoundDisplay5).FirstOrDefault();
                
                if (header1[4] != string.Empty && header1[4] != null)
                {
                    column++;
                    listColumn.Add(column);
                }

                header1[5] = listTestRound.Where(x => x.TestingRound == "Round 3").Select(x => x.HeaderRoundDisplay6).FirstOrDefault();

                if (header1[5] != string.Empty && header1[5] != null)
                {
                    column++;
                    listColumn.Add(column);
                }

                header1[6] = listTestRound.Where(x => x.TestingRound == "Round 3").Select(x => x.HeaderRoundDisplay7).FirstOrDefault();

                if (header1[6] != string.Empty && header1[6] != null)
                {
                    column++;
                    listColumn.Add(column);
                }

                header1[7] = listTestRound.Where(x => x.TestingRound == "Round 3").Select(x => x.HeaderRoundDisplay8).FirstOrDefault();

                if (header1[7] != string.Empty && header1[7] != null)
                {
                    column++;
                    listColumn.Add(column);
                }

                header1[8] = listTestRound.Where(x => x.TestingRound == "Round 3").Select(x => x.HeaderRoundDisplay9).FirstOrDefault();

                if (header1[8] != string.Empty && header1[8] != null)
                {
                    column++;
                    listColumn.Add(column);
                }

                header1[9] = listTestRound.Where(x => x.TestingRound == "Round 3").Select(x => x.HeaderRoundDisplay10).FirstOrDefault();

                if (header1[9] != string.Empty && header1[9] != null)
                {
                    column++;
                    listColumn.Add(column);
                }

                column++; //status
                listColumn.Add(column);
                column++; //for a2q2 comments
                listColumn.Add(column);

                //Console.WriteLine($"listColumn.Count: {listColumn.Count}");

                ws.Cells[row, listColumn[0]].Value = "Testing Round";
                ws.Cells[row, listColumn[1]].Value = "A2Q2 Samples";

                int columnHeaderStart = 3;
                int countHeader = 0;
                if (listColumn.Count - 4 > 0)
                {
                    for (int i = 0; i < listColumn.Count - 4; i++)
                    {
                        if (header1[i] != string.Empty)
                        {
                            countHeader++;
                            ws.Cells[row, columnHeaderStart].Value = header1[i];
                            columnHeaderStart++;
                        }
                    }
                }

                ws.Cells[row, listColumn[listColumn.Count - 2]].Value = "Status";
                ws.Cells[row, listColumn[listColumn.Count - 1]].Value = "A2Q2 Comments";
                ws.Cells[row, listColumn[listColumn.Count - 1], row, listColumn[listColumn.Count - 1] + 2].Merge = true;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.WrapText = true;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Font.SetFromFont(new Font("Arial", 12));
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#d9e1f2"));
                ws.Cells[row, listColumn[listColumn.Count - 2], row, listColumn[listColumn.Count - 1] + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, listColumn[listColumn.Count - 2], row, listColumn[listColumn.Count - 1] + 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#ccffcc"));
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1]].Style.Font.Bold = true;
                ws.Row(row).Height = 42;

                row++;
                foreach (var item in sampleSelection.ListTestRound)
                {
                    if (item.TestingRound == "Round 3")
                    {
                        int columnIndex = 0;
                        ws.Cells[row, listColumn[columnIndex]].Value = item.TestingRound;
                        columnIndex++;
                        ws.Cells[row, listColumn[columnIndex]].Value = item.A2Q2Samples;

                        
                        if (item.ContentDisplay1 != string.Empty && item.ContentDisplay1 != null)
                        {
                            columnIndex++;
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay1;
                        }

                        
                        if (item.ContentDisplay2 != string.Empty && item.ContentDisplay2 != null)
                        {
                            columnIndex++;
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay2;
                        }

                        
                        if (item.ContentDisplay3 != string.Empty && item.ContentDisplay3 != null)
                        {
                            columnIndex++;
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay3;
                        }

                        
                        if (item.ContentDisplay4 != string.Empty && item.ContentDisplay4 != null)
                        {
                            columnIndex++;
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay4;
                        }

                        
                        if (item.ContentDisplay5 != string.Empty && item.ContentDisplay5 != null)
                        {
                            columnIndex++;
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay5;
                        }

                        if (item.ContentDisplay6 != string.Empty && item.ContentDisplay6 != null)
                        {
                            columnIndex++;
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay6;
                        }

                        if (item.ContentDisplay7 != string.Empty && item.ContentDisplay7 != null)
                        {
                            columnIndex++;
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay7;
                        }

                        if (item.ContentDisplay8 != string.Empty && item.ContentDisplay8 != null)
                        {
                            columnIndex++;
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay8;
                        }

                        if (item.ContentDisplay9 != string.Empty && item.ContentDisplay9 != null)
                        {
                            columnIndex++;
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay9;
                        }

                        if (item.ContentDisplay10 != string.Empty && item.ContentDisplay10 != null)
                        {
                            columnIndex++;
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay10;
                        }

                        columnIndex++;
                        ws.Cells[row, columnIndex + 1].Value = item.Status;
                        columnIndex++;
                        ws.Cells[row, columnIndex + 1].Value = item.Comment;

                        #region Format
                        ws.Cells[row, listColumn[listColumn.Count - 1], row, listColumn[listColumn.Count - 1] + 2].Merge = true;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Font.SetFromFont(new Font("Arial", 12));
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#d9e1f2"));
                        ws.Cells[row, listColumn[listColumn.Count - 2], row, listColumn[listColumn.Count - 1] + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[row, listColumn[listColumn.Count - 2], row, listColumn[listColumn.Count - 1] + 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#ccffcc"));
                        #endregion

                        row++;
                    }

                }

            }


            #endregion
        
        }

        #endregion

        private bool IsDateTime(string dateTime)
        {
            bool result = false;
            DateTime dtValue;
            DateTime.TryParse(dateTime, out dtValue);
            if (dtValue.Year != 0001)
                result = true;

            return result;
        }

        private bool IsDecimal(string decimalValue)
        {
            bool result = false;
            decimal intValue;
            decimal.TryParse(decimalValue, out intValue);
            if (intValue != 0)
                result = true;

            return result;
        }

        private bool IsDateHeader(string decimalValue)
        {
            bool result = false;
            if (decimalValue != null)
            {
                if (decimalValue.Contains("date", StringComparison.OrdinalIgnoreCase))
                    result = true;
            }
            return result;
        }

        private bool IsIdentificationHeader(string decimalValue)
        {
            bool result = false;
            if (decimalValue != null)
            {
                if (decimalValue.Contains("id", StringComparison.OrdinalIgnoreCase))
                    result = true;
            }
            return result;
        }

    }
}
