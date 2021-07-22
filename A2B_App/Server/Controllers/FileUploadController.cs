using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using A2B_App.Server.Data;
using A2B_App.Server.JobScheduler;
using A2B_App.Server.Services;
using A2B_App.Shared.Podio;
using A2B_App.Shared.Sox;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using PodioAPI;
using PodioAPI.Models;
using Quartz;

namespace A2B_App.Server.Controllers
{
    [Authorize]
    [ApiController]
    [Route("api/[controller]")]
    public class FileUploadController : ControllerBase
    {
        private readonly IConfiguration _config;
        private readonly ILogger<FileUploadController> _logger;
        private readonly IWebHostEnvironment _environment;
        private readonly SoxContext _soxContext;
        private readonly IScheduler _scheduler;
        public FileUploadController(IConfiguration config, 
            ILogger<FileUploadController> logger, 
            IWebHostEnvironment environment, 
            SoxContext soxContext,
            IScheduler scheduler)
        {
            _config = config;
            _logger = logger;
            _environment = environment;
            _soxContext = soxContext;
            _scheduler = scheduler;
        }

        [HttpPost("upload")]
        public async Task<IActionResult> UploadFileAsync([FromForm] IFormFile file)
        {
            try
            {
                //check if file is valid
                if (file == null || file.Length == 0)
                    return BadRequest("Invalid File");

                string fileName = file.FileName;
                string extension = Path.GetExtension(fileName).ToLowerInvariant(); ;

                //set allowed extensions
                string[] allowedExtensions = { ".xlsx", ".xls" };

                //check if extension is valid
                if (!allowedExtensions.Contains(extension))
                    return BadRequest("Invalid File");

                string startupPath = Directory.GetCurrentDirectory();
                string newFilename = $"{Guid.NewGuid()}{extension}";
                string path = Path.Combine(startupPath, "include", "upload", newFilename);
                string filePath = path;

                using (var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    await file.CopyToAsync(fileStream);
                }

                return Ok($"{newFilename}");
            }
            catch (Exception ex)
            {
                FileLog.Write($"Error UploadFileAsync {ex}", "ErrorUploadFileAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "UploadFileAsync");
                return BadRequest();
            }
           
        }

        //[AllowAnonymous]
        [HttpPost("soxtracker")]
        public async Task<IActionResult> UploadSoxTrackerAsync([FromForm] IFormFile file)
        {
            try
            {
                //check if file is valid
                if (file == null || file.Length == 0)
                    return BadRequest("Invalid File");

                string fileName = file.FileName;
                string extension = Path.GetExtension(fileName).ToLowerInvariant(); ;

                //set allowed extensions
                string[] allowedExtensions = { ".xlsx", ".xls" };

                //check if extension is valid
                if (!allowedExtensions.Contains(extension))
                    return BadRequest("Invalid File");

                string startupPath = Directory.GetCurrentDirectory();
                string newFilename = $"Tracker_{Guid.NewGuid()}{extension}";
                string path = Path.Combine(startupPath, "include", "upload", newFilename);
                string filePath = path;

                using (var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    await file.CopyToAsync(fileStream);
                }

                using (var context = _soxContext.Database.BeginTransaction())
                {
                    var fi = new FileInfo(filePath);

                    using (ExcelPackage p = new ExcelPackage(fi))
                    {
                        // If you use EPPlus in a noncommercial context
                        // according to the Polyform Noncommercial license:
                        //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                        ExcelWorksheet ws = p.Workbook.Worksheets[0];
                        int colCount = ws.Dimension.End.Column;
                        int rowCount = ws.Dimension.End.Row;

                        List<SoxTracker> listSoxTracker = new List<SoxTracker>();
                        string clientName = string.Empty;
                        string fY = string.Empty;

                        if (ws.Cells[1, 1].Value?.ToString() != string.Empty)
                        {
                            clientName = ws.Cells[1, 1].Value?.ToString();
                        }

                        if (ws.Cells[3, 1].Value?.ToString() != string.Empty)
                        {
                            fY = ws.Cells[3, 1].Value?.ToString();
                        }


                        //loop start at row 7
                        for (int row = 7; row <= rowCount; row++)
                        {
                            SoxTracker soxTracker = new SoxTracker();

                            //loop column
                            for (int col = 1; col <= colCount; col++)
                            {

                                switch (col)
                                {
                                    case 1:
                                        if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                                        {
                                            soxTracker.Process = ws.Cells[row, col].Value?.ToString();
                                        }
                                        break;
                                    case 2:
                                        if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                                        {
                                            soxTracker.Subprocess = ws.Cells[row, col].Value?.ToString();
                                        }
                                        break;
                                    case 3:
                                        if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                                        {
                                            soxTracker.ControlId = ws.Cells[row, col].Value?.ToString();
                                        }
                                        break;

                                    case 7:
                                        if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                                        {
                                            soxTracker.PBC = ws.Cells[row, col].Value?.ToString();
                                        }
                                        break;
                                    case 8:
                                        if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                                        {
                                            soxTracker.PBCOwner = ws.Cells[row, col].Value?.ToString();
                                        }
                                        break;
                                    case 9:
                                        if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                                        {
                                            soxTracker.PopulationFileRequest = ws.Cells[row, col].Value?.ToString();
                                        }
                                        break;
                                    case 10:
                                        if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                                        {
                                            soxTracker.SampleSelection = ws.Cells[row, col].Value?.ToString();
                                        }
                                        break;
                                    case 11:
                                        if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                                        {
                                            soxTracker.R3Sample = ws.Cells[row, col].Value?.ToString();
                                        }
                                        break;
                                    case 12:
                                        if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                                        {
                                            soxTracker.WTPBC = ws.Cells[row, col].Value?.ToString();
                                        }
                                        break;
                                    case 13:
                                        if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                                        {
                                            soxTracker.R1PBC = ws.Cells[row, col].Value?.ToString();
                                        }
                                        break;
                                    case 14:
                                        if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                                        {
                                            soxTracker.R2PBC = ws.Cells[row, col].Value?.ToString();
                                        }
                                        break;
                                    case 15:
                                        if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                                        {
                                            soxTracker.R3PBC = ws.Cells[row, col].Value?.ToString();
                                        }
                                        break;
                                    case 16:
                                        if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                                        {
                                            soxTracker.WTTester = ws.Cells[row, col].Value?.ToString();
                                        }
                                        break;
                                    case 17:
                                        if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                                        {
                                            soxTracker.WT1LReviewer = ws.Cells[row, col].Value?.ToString();
                                        }
                                        break;
                                    case 18:
                                        if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                                        {
                                            soxTracker.WT2LReviewer = ws.Cells[row, col].Value?.ToString();
                                        }
                                        break;
                                    case 19:
                                        if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                                        {
                                            soxTracker.WTTestingStatus = ws.Cells[row, col].Value?.ToString();
                                        }
                                        break;
                                    case 20:
                                        if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                                        {
                                            soxTracker.R1Tester = ws.Cells[row, col].Value?.ToString();
                                        }
                                        break;
                                    case 21:
                                        if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                                        {
                                            soxTracker.R11LReviewer = ws.Cells[row, col].Value?.ToString();
                                        }
                                        break;
                                    case 22:
                                        if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                                        {
                                            soxTracker.R12LReviewer = ws.Cells[row, col].Value?.ToString();
                                        }
                                        break;
                                    case 23:
                                        if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                                        {
                                            soxTracker.R1TestingStatus = ws.Cells[row, col].Value?.ToString();
                                        }
                                        break;
                                    case 24:
                                        if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                                        {
                                            soxTracker.R2Tester = ws.Cells[row, col].Value?.ToString();
                                        }
                                        break;
                                    case 25:
                                        if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                                        {
                                            soxTracker.R21LReviewer = ws.Cells[row, col].Value?.ToString();
                                        }
                                        break;
                                    case 26:
                                        if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                                        {
                                            soxTracker.R22LReviewer = ws.Cells[row, col].Value?.ToString();
                                        }
                                        break;
                                    case 27:
                                        if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                                        {
                                            soxTracker.R2TestingStatus = ws.Cells[row, col].Value?.ToString();
                                        }
                                        break;
                                    case 28:
                                        if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                                        {
                                            soxTracker.R3Tester = ws.Cells[row, col].Value?.ToString();
                                        }
                                        break;
                                    case 29:
                                        if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                                        {
                                            soxTracker.R31LReviewer = ws.Cells[row, col].Value?.ToString();
                                        }
                                        break;
                                    case 30:
                                        if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                                        {
                                            soxTracker.R32LReviewer = ws.Cells[row, col].Value?.ToString();
                                        }
                                        break;
                                    case 31:
                                        if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                                        {
                                            soxTracker.R3TestingStatus = ws.Cells[row, col].Value?.ToString();
                                        }
                                        break;

                                }
                            }


                            if (soxTracker != null)
                            {
                                soxTracker.ClientName = clientName;
                                soxTracker.FY = fY;
                                listSoxTracker.Add(soxTracker);
                            }

                        }


                        if (listSoxTracker != null && listSoxTracker.Count > 0)
                        {
                            foreach (var item in listSoxTracker)
                            {
                                var checkSoxTracker = _soxContext.SoxTracker
                                    .Where(x => x.ClientName.Equals(item.ClientName)
                                        && x.ControlId.Equals(item.ControlId)
                                        && x.FY.Equals(item.FY)
                                        && x.Process.Equals(item.Process)
                                        && x.Subprocess.Equals(item.Subprocess)
                                        )
                                    .FirstOrDefault();
                                if (checkSoxTracker != null)
                                {
                                    item.Id = checkSoxTracker.Id;
                                    _soxContext.Entry(checkSoxTracker).CurrentValues.SetValues(item);
                                    await _soxContext.SaveChangesAsync();
                                }
                                else
                                {
                                    _soxContext.Add(item);
                                    await _soxContext.SaveChangesAsync();
                                }
                            }

                            context.Commit();
                        }

                    }
                }

                return Ok($"{newFilename}");
            }
            catch (Exception ex)
            {
                FileLog.Write($"Error UploadSoxTrackerAsync {ex}", "ErrorUploadSoxTrackerAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "UploadSoxTrackerAsync");
                return BadRequest();

            }
           
        }

        //[AllowAnonymous]
        [HttpPost("image")]
        public async Task<IActionResult> UploadImageAsync([FromForm] IFormFile file)
        {
            try
            {
                //check if file is valid
                if (file == null || file.Length == 0)
                    return BadRequest("Invalid File");

                string fileName = file.FileName;
                string extension = Path.GetExtension(fileName).ToLowerInvariant(); ;

                //set allowed extensions
                string[] allowedExtensions = { ".jpg", ".png" };

                //check if extension is valid
                if (!allowedExtensions.Contains(extension))
                    return BadRequest("Invalid File");

                string startupPath = Directory.GetCurrentDirectory();
                string newFilename = $"K_{Guid.NewGuid()}{extension}";
                string path = Path.Combine(startupPath, "include", "upload", "image", newFilename);
                string filePath = path;

                using (var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    await file.CopyToAsync(fileStream);
                }

                return Ok($"{newFilename}");
            }
            catch (Exception ex)
            {
                FileLog.Write($"Error UploadImageAsync {ex}", "ErrorUploadImageAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "UploadImageAsync");
                return BadRequest();

            }
           

        }

        [HttpPost("keyreport")]
        public async Task<IActionResult> UploadKeyReportAsync([FromForm] IFormFile file)
        {
            try
            {
                //check if file is valid
                if (file == null || file.Length == 0)
                    return BadRequest("Invalid File");

                string fileName = file.FileName;
                string extension = Path.GetExtension(fileName).ToLowerInvariant(); ;

                //set allowed extensions
                string[] allowedExtensions = { ".xlsx", ".xls" };

                //check if extension is valid
                if (!allowedExtensions.Contains(extension))
                    return BadRequest("Invalid File");

                string startupPath = Directory.GetCurrentDirectory();
                string newFilename = $"K_{Guid.NewGuid()}{extension}";
                string path = Path.Combine(startupPath, "include", "upload", "image", newFilename);
                string filePath = path;

                using (var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    await file.CopyToAsync(fileStream);
                }

                return Ok($"{newFilename}");
            }
            catch (Exception ex)
            {
                FileLog.Write($"Error UploadKeyReportAsync {ex}", "ErrorUploadKeyReportAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "UploadKeyReportAsync");
                return BadRequest();

            }


        }


        [HttpPost("uploadFileKeyReport")]
        public async Task<IActionResult> UploadFileForKeyReportAsync([FromForm] IFormFile file)
        {
            try
            {
                //check if file is valid
                if (file == null || file.Length == 0)
                    return BadRequest("Invalid File");

                string fileName = file.FileName;
                string extension = Path.GetExtension(fileName).ToLowerInvariant(); ;

                //set allowed extensions
                string[] allowedExtensions = { ".xlsx", ".xls" };

                //check if extension is valid
                if (!allowedExtensions.Contains(extension))
                    return BadRequest("Invalid File");

                string startupPath = Directory.GetCurrentDirectory();
                //string newFilename = $"{Guid.NewGuid()}{extension}";
                string path = Path.Combine(startupPath, "include", "upload/keyreport/kr tracker", fileName);
                string filePath = path;

                using (var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    await file.CopyToAsync(fileStream);

                    var keyreportRaw = await KeyReportFileUpload(path);

                    ImportFields importFields = new ImportFields();
                    importFields.ListExcelColumns = keyreportRaw.Item1;
                    importFields.ListDatabaseColumns = keyreportRaw.Item2;
                    importFields.Filename = fileName;
                    return Ok(importFields);

                }
                //return NoContent();
            }
            catch (Exception ex)
            {
                FileLog.Write($"Error UploadFileForRcmAsync {ex}", "ErrorUploadFileForRcmAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "UploadFileForRcmAsync");
                return BadRequest();
            }

        }


        //[AllowAnonymous]
        [HttpPost("uploadFileRcm")]
        public async Task<IActionResult> UploadFileForRcmAsync([FromForm] IFormFile file)
        {
            try
            {
                //check if file is valid
                if (file == null || file.Length == 0)
                    return BadRequest("Invalid File");

                string fileName = file.FileName;
                string extension = Path.GetExtension(fileName).ToLowerInvariant(); ;

                //set allowed extensions
                string[] allowedExtensions = { ".xlsx", ".xls" };

                //check if extension is valid
                if (!allowedExtensions.Contains(extension))
                    return BadRequest("Invalid File");

                string startupPath = Directory.GetCurrentDirectory();
                string newFilename = $"{Guid.NewGuid()}{extension}";
                string path = Path.Combine(startupPath, "include", "upload", newFilename);
                string filePath = path;

                using (var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    await file.CopyToAsync(fileStream);

                    var rcmRaw = await RcmFileUpload(path);

                    ImportFields importFields = new ImportFields();
                    importFields.ListExcelColumns = rcmRaw.Item1;
                    importFields.ListDatabaseColumns = rcmRaw.Item2;
                    importFields.Filename = newFilename;
                    return Ok(importFields);

                }


                //return NoContent();
            }
            catch (Exception ex)
            {
                FileLog.Write($"Error UploadFileForRcmAsync {ex}", "ErrorUploadFileForRcmAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "UploadFileForRcmAsync");
                return BadRequest();
            }

        }

        //[AllowAnonymous]
        [HttpPost("uploadFileSod")]
        public async Task<IActionResult> UploadFileForSodAsync([FromForm] FileImport fileImport)
        {
            try
            {


                //check if file is valid
                if (fileImport.File == null || fileImport.File.Length == 0)
                    return BadRequest("Invalid File");

                string fileName = fileImport.File.FileName;
                string extension = Path.GetExtension(fileName).ToLowerInvariant(); ;

                //set allowed extensions
                string[] allowedExtensions = { ".xlsx", ".xls" };

                //check if extension is valid
                if (!allowedExtensions.Contains(extension))
                    return BadRequest("Invalid File");

                string startupPath = Directory.GetCurrentDirectory();
                string newFilename = $"{Guid.NewGuid()}{extension}";
                string path = Path.Combine(startupPath, "include", "upload", newFilename);
                string filePath = path;

                using (var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    await fileImport.File.CopyToAsync(fileStream);

                    SodService sodService = new SodService(_soxContext, _config);
                    ServiceResponse response = new ServiceResponse();
                    response = await sodService.ProcessSodFile(newFilename, fileImport);
                    return Ok(response);
                }


                //return NoContent();
            }
            catch (Exception ex)
            {
                FileLog.Write($"Error UploadFileForSodAsync {ex}", "ErrorUploadFileForSodAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "UploadFileForSodAsync");
                return BadRequest();
            }

        }


        [AllowAnonymous]
        [HttpPost("uploadFileSodSoxRox")]
        public async Task<IActionResult> UploadFileForSodSoxRoxAsync([FromForm] SodSoxRoxImport fileImport)
        {
            try
            {
                SodService sodService = new SodService(_soxContext, _config);

                //check if file is valid
                if (fileImport.FileRoleUser == null || fileImport.FileRoleUser.Length == 0 || fileImport.FileRolePerm == null || fileImport.FileRolePerm.Length == 0)
                    return BadRequest("Invalid File for Role User or Role Permission");


                string startupPath = Directory.GetCurrentDirectory();
                string path = string.Empty;
                string newFilename = string.Empty;


                List<SoxRoxFile> listSoxFile = new List<SoxRoxFile>();

                SoxRoxFile soxFile1 = new SoxRoxFile();
                soxFile1.FileName = fileImport.FileRoleUser.FileName ?? string.Empty;
                soxFile1.Position = 0;
                soxFile1.ClientName = fileImport.ClientName;
                listSoxFile.Add(soxFile1);

                SoxRoxFile soxFile2 = new SoxRoxFile();
                soxFile2.FileName = fileImport.FileRolePerm.FileName ?? string.Empty;
                soxFile2.Position = 1;
                soxFile2.ClientName = fileImport.ClientName;
                listSoxFile.Add(soxFile2);

                SoxRoxFile soxFile3 = new SoxRoxFile();
                soxFile3.FileName = fileImport.FileConflictPerm.FileName ?? string.Empty;
                soxFile3.Position = 2;
                soxFile3.ClientName = fileImport.ClientName;
                listSoxFile.Add(soxFile3);

                SoxRoxFile soxFile4 = new SoxRoxFile();
                soxFile4.FileName = fileImport.FileDescToPerm.FileName ?? string.Empty;
                soxFile4.Position = 3;
                soxFile4.ClientName = fileImport.ClientName;
                listSoxFile.Add(soxFile4);


                foreach (var item in listSoxFile)
                {
                    if(item.FileName != string.Empty)
                    {
                        //check file extension
                        string extension = Path.GetExtension(item.FileName).ToLowerInvariant();
                        //set allowed extensions
                        string[] allowedExtensions = { ".xlsx", ".xls" };
                        //check if extension is valid
                        if (!allowedExtensions.Contains(extension))
                            return BadRequest($"Invalid File Extension {item.FileName}");

                        if (item.Position > 1)
                            newFilename = item.FileName; //conflict and description will retain the filename
                        else
                            newFilename = $"{Guid.NewGuid()}{extension}"; // role user and role perm is rename

                        path = Path.Combine(startupPath, "include", "upload", newFilename);

                        using (var fileStream = new FileStream(path, FileMode.Append, FileAccess.Write))
                        {
                            
                            switch (item.Position)
                            {
                                case 0:
                                    Debug.WriteLine($"Importing {fileImport.FileRoleUser.FileName}");
                                    await fileImport.FileRoleUser.CopyToAsync(fileStream);
                                    break;
                                case 1:
                                    Debug.WriteLine($"Importing {fileImport.FileRolePerm.FileName}");
                                    await fileImport.FileRolePerm.CopyToAsync(fileStream);
                                    break;
                                case 2:
                                    Debug.WriteLine($"Importing {fileImport.FileConflictPerm.FileName}");
                                    await fileImport.FileConflictPerm.CopyToAsync(fileStream);
                                    break;
                                case 3:
                                    Debug.WriteLine($"Importing {fileImport.FileDescToPerm.FileName}");
                                    await fileImport.FileDescToPerm.CopyToAsync(fileStream);
                                    break;
                                default:
                                    break;
                            }
                            item.NewFileName = newFilename;

                        }
                    }
                }

                if(listSoxFile != null && listSoxFile.Count > 0)
                {


                    Debug.WriteLine($"Schedule job process for SOD SoxRox");
                    ITrigger trigger = TriggerBuilder.Create()
                        .WithIdentity($"{DateTime.Now}-SODSoxRox")
                        .StartNow()
                        .WithPriority(2)
                        .Build();

                    IJobDetail job = JobBuilder.Create<SodSoxRoxJob>()
                        .WithIdentity($"{DateTime.Now}-SODSoxRox")
                        .Build();
                    job.JobDataMap["object"] = listSoxFile;
                    job.JobDataMap["context"] = _soxContext;
                    job.JobDataMap["config"] = _config;
                    job.JobDataMap["requestedBy"] = fileImport.RequestedBy;

                    await _scheduler.ScheduleJob(job, trigger);




                    //SodSoxRoxInput response = new SodSoxRoxInput();
                    //string clientName = listSoxFile[0].ClientName;

                    //Debug.WriteLine($"Reading SOD SoxRox File | {DateTime.Now}");
                    //#region Reading SOD SoxRox File
                    ////run simultaneously
                    //var taskRoleUser = System.Threading.Tasks.Task.Run(() => sodService.ReadFileSodSoxRoxRoleUser(listSoxFile, clientName));
                    //var taskRolePerm = System.Threading.Tasks.Task.Run(() => sodService.ReadFileSodSoxRoxRolePerm(listSoxFile, clientName));
                    //var taskDescPerm = System.Threading.Tasks.Task.Run(() => sodService.ReadFileSodSoxRoxDescToPerm(listSoxFile, clientName));
                    //var taskConflict = System.Threading.Tasks.Task.Run(() => sodService.ReadFileSodSoxRoxConflictPerm(listSoxFile, clientName));

                    //System.Threading.Tasks.Task.WhenAll(taskRoleUser, taskRolePerm, taskDescPerm, taskConflict).Wait();
                    ////Task.WhenAll(taskDescPerm, taskConflict);

                    //response.ListRoleUser = taskRoleUser.Result;
                    //response.ListRolePerm = taskRolePerm.Result;
                    //response.ListDescriptionToPerm = taskDescPerm.Result;
                    //response.ListConflictPerm = taskConflict.Result;
                    //response.ListRoleUserTrim = sodService.GetSodSoxRoxRoleUser(response);
                    //Debug.WriteLine($"DONE Reading SOD SoxRox File | {DateTime.Now}");
                    //#endregion



                    //Debug.WriteLine($"Processing SOD SoxRox Analysis | {DateTime.Now}");
                    //#region Processing SOD SoxRox Analysis
                    //var sodSoxRoxRaw2 = System.Threading.Tasks.Task.Run(() => sodService.ProcessSoxRoxDataRaw2_3(response));
                    //var sodSoxRoxRaw3 = System.Threading.Tasks.Task.Run(() => sodService.ProcessSoxRoxDataRaw3_3(response));
                    //var taskCreateDescPerm = System.Threading.Tasks.Task.Run(() => sodService.ProcessSoxRoxDescriptionOutput(response));
                    //System.Threading.Tasks.Task.WhenAll(sodSoxRoxRaw2, sodSoxRoxRaw3, taskCreateDescPerm).Wait();
                    //Debug.WriteLine($"DONE Processing SOD SoxRox Analysis | {DateTime.Now}");
                    //#endregion



                    //Debug.WriteLine($"Creating SOD SoxRox Report | {DateTime.Now}");
                    //#region Creating SOD SoxRox Report
                    ////create excel
                    //var taskCreateDescriptionOutput = System.Threading.Tasks.Task.Run(() => sodService.CreateSoxRoxDescriptionFile(taskCreateDescPerm.Result, clientName));
                    //var taskCreatereport = System.Threading.Tasks.Task.Run(() => sodService.CreateSodSoxRoxFile(sodSoxRoxRaw2.Result, sodSoxRoxRaw3.Result, clientName));
                    //System.Threading.Tasks.Task.WhenAll(taskCreateDescriptionOutput, taskCreatereport).Wait();

                    //SodSoxRoxOutputFile sodSoxRoxOutputFile = new SodSoxRoxOutputFile();
                    //sodSoxRoxOutputFile.ReportFileName = taskCreatereport.Result;
                    //sodSoxRoxOutputFile.DescriptionFileName = taskCreateDescriptionOutput.Result;
                    //Debug.WriteLine($"DONE Creating SOD SoxRox Report | {DateTime.Now}");
                    //#endregion


                    ////Debug.WriteLine($"Compress files | {DateTime.Now}");
                    ////#region Compress files
                    ////startupPath = Directory.GetCurrentDirectory();
                    ////string strSourceDownload = Path.Combine(startupPath, "include", "sod");
                    ////string strOutput1 = Path.Combine(strSourceDownload, sodSoxRoxOutputFile.ReportFileName);
                    ////string strOutput2 = Path.Combine(strSourceDownload, sodSoxRoxOutputFile.DescriptionFileName);
                    ////List<string> listSodReport = new List<string>();
                    ////listSodReport.Add(sodSoxRoxOutputFile.ReportFileName);
                    ////listSodReport.Add(sodSoxRoxOutputFile.DescriptionFileName);
                    ////AdminService adminService = new AdminService(_config);
                    ////var taskCompressReportFile = System.Threading.Tasks.Task.Run(() => adminService.CompressItem(strSourceDownload, $"{clientName}_SODReport", listSodReport));
                    ////System.Threading.Tasks.Task.WhenAll(taskCompressReportFile).Wait();
                    ////string zipFileName = taskCompressReportFile.Result;
                    ////Debug.WriteLine($"DONE Compress files | {DateTime.Now}");
                    ////#endregion




                    //SodSoxRoxInput response = new SodSoxRoxInput();

                    ////run synchronously
                    ////var taskRoleUser = sodService.ReadSodSoxRoxRoleUser(listSoxFile, fileImport.ClientName);
                    ////var taskRolePerm = sodService.ReadSodSoxRoxRolePerm(listSoxFile, fileImport.ClientName);
                    ////var taskDescPerm = sodService.ReadSodSoxRoxDescToPerm(listSoxFile, fileImport.ClientName);

                    ////run simultaneously
                    //var taskRoleUser = System.Threading.Tasks.Task.Run(() => sodService.ReadFileSodSoxRoxRoleUser(listSoxFile, fileImport.ClientName));
                    //var taskRolePerm = System.Threading.Tasks.Task.Run(() => sodService.ReadFileSodSoxRoxRolePerm(listSoxFile, fileImport.ClientName));
                    //var taskDescPerm = System.Threading.Tasks.Task.Run(() => sodService.ReadFileSodSoxRoxDescToPerm(listSoxFile, fileImport.ClientName));
                    //var taskConflict = System.Threading.Tasks.Task.Run(() => sodService.ReadFileSodSoxRoxConflictPerm(listSoxFile, fileImport.ClientName));

                    //await System.Threading.Tasks.Task.WhenAll(taskRoleUser, taskRolePerm, taskDescPerm, taskConflict);
                    //await System.Threading.Tasks.Task.WhenAll(taskDescPerm, taskConflict);

                    //response.ListRoleUser = taskRoleUser.Result;
                    //response.ListRolePerm = taskRolePerm.Result;
                    //response.ListDescriptionToPerm = taskDescPerm.Result;
                    //response.ListConflictPerm = taskConflict.Result;
                    //response.ListRoleUserTrim = await sodService.GetSodSoxRoxRoleUser(response);

                    ////var taskCreateDescPerm = System.Threading.Tasks.Task.Run(() => sodService.ProcessSoxRoxDescriptionOutput(response));
                    ////await System.Threading.Tasks.Task.WhenAll(taskCreateDescPerm);

                    //////List<SodSoxRoxDescriptionPermission> listSoxRoxPermA = new List<SodSoxRoxDescriptionPermission>();
                    //////listSoxRoxPermA = taskCreateDescPerm.Result;

                    ////var taskCreateDescriptionOutput = System.Threading.Tasks.Task.Run(() => sodService.CreateSoxRoxDescriptionFile(taskCreateDescPerm.Result, fileImport.ClientName));
                    ////await System.Threading.Tasks.Task.WhenAll(taskCreateDescriptionOutput);

                    ////SodSoxRoxOutputFile soxRoxFile = new SodSoxRoxOutputFile();
                    ////soxRoxFile.DescriptionFileName = taskCreateDescriptionOutput.Result;

                    ////return Ok(response);
                }

                return Ok();

            }
            catch (Exception ex)
            {
                FileLog.Write($"Error UploadFileForSodSoxRoxAsync {ex}", "ErrorUploadFileForSodSoxRoxAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "UploadFileForSodSoxRoxAsync");
                return BadRequest();
            }

        }

        [AllowAnonymous]
        [HttpPost("processSoxRox")]
        public async Task<IActionResult> ProcessSodSoxRoxAsync([FromBody] SodSoxRoxInput sodSoxRox)
        {
            try
            {
                //do analysis
                SodService sodService = new SodService(_soxContext, _config);
                //var sodSoxRoxRaw2 = await sodService.ProcessSoxRoxDataRaw2(sodSoxRox);
                //var sodSoxRoxRaw3 = await sodService.ProcessSoxRoxDataRaw3(sodSoxRox);

                var sodSoxRoxRaw2 = System.Threading.Tasks.Task.Run(() => sodService.ProcessSoxRoxDataRaw2_3(sodSoxRox));
                var sodSoxRoxRaw3 = System.Threading.Tasks.Task.Run(() => sodService.ProcessSoxRoxDataRaw3_3(sodSoxRox));

                await System.Threading.Tasks.Task.WhenAll(sodSoxRoxRaw2, sodSoxRoxRaw3);

                string clientName = sodSoxRox.ListRoleUser[0].Client;

                //create excel
                SodSoxRoxOutputFile sodSoxRoxOutputFile = new SodSoxRoxOutputFile();
                sodSoxRoxOutputFile.ReportFileName = await sodService.CreateSodSoxRoxFile(sodSoxRoxRaw2.Result, sodSoxRoxRaw3.Result, clientName);

                return Ok(sodSoxRoxOutputFile);

            }
            catch (Exception ex)
            {
                FileLog.Write($"Error ProcessSodSoxRoxAsync {ex}", "ErrorProcessSodSoxRoxAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "UploadFileForSodSoxRoxAsync");
                return BadRequest();
            }

        }

        [AllowAnonymous]
        [HttpPost("getSodSoxRoxUser")]
        public IActionResult ProcessSodSoxRoxRoleUserAsync([FromBody] SodSoxRoxInput sodSoxRox)
        {
            try
            {
                SodService sodService = new SodService(_soxContext, _config);
                var sodSoxRoleUser = sodService.GetSodSoxRoxRoleUser(sodSoxRox);

                //var listRoles = sodService.GetMap02Roles(sodSoxRox.ListRolePerm);

                return Ok(sodSoxRoleUser);

            }
            catch (Exception ex)
            {
                FileLog.Write($"Error ProcessSodSoxRoxRoleUserAsync {ex}", "ErrorProcessSodSoxRoxRoleUserAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "ProcessSodSoxRoxRoleUserAsync");
                return BadRequest();
            }

        }


        //[AllowAnonymous]
        [HttpPost("importRcm")]
        public async Task<IActionResult> ImportRcmAsync([FromBody] ImportFields importFields)
        {
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("RcmPodioApp").GetSection("RcmAppId").Value;
                PodioAppKey.AppToken = _config.GetSection("RcmPodioApp").GetSection("RcmAppToken").Value;

                RcmService rcmService = new RcmService(_soxContext, _config);

                WriteLog writeLog = new WriteLog();
                string startupPath = Directory.GetCurrentDirectory();
                string path = Path.Combine(startupPath, "include", "upload", importFields.Filename);
                string filePath = path;

                FileInfo fi = new FileInfo(path);
                List<ColumnVal> columnField = new List<ColumnVal>();
                List<DBColumnVal> dbColumn = new List<DBColumnVal>();
                using (ExcelPackage p = new ExcelPackage(fi))
                {
                    // If you use EPPlus in a noncommercial context
                    // according to the Polyform Noncommercial license:
                    //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    //Console.WriteLine($"Worksheet Count: {p.Workbook.Worksheets.Count}");
                    ExcelWorksheet ws = p.Workbook.Worksheets[0];
                    int colCount = ws.Dimension.End.Column;
                    int rowCount = ws.Dimension.End.Row;

                    //get all value from excel file and store in listRowVal
                    List<RowVal> listRowVal = new List<RowVal>();
                    for (int row = 2; row <= rowCount; row++) //we start at row 2 since row 1 is the header or title
                    {

                        RowVal rowVal = new RowVal();
                        List<ColumnVal> colValue = new List<ColumnVal>();
                        for (int col = 1; col <= colCount; col++)
                        {
                            ColumnVal column = new ColumnVal();
                            column.Index = col;
                            column.ExcelColumnName = ws.Cells[row, col].Value?.ToString();
                            colValue.Add(column);
                        }
                        rowVal.Column = colValue;
                        rowVal.Row = row;
                        listRowVal.Add(rowVal);
                    }

                    foreach (var item in listRowVal)
                    {
                        Rcm rcm = new Rcm();

                        var xlsColClientName = importFields.ListDatabaseColumns.Where(x => x.DbColumnName.Equals("ClientName")).FirstOrDefault();
                        if (xlsColClientName.ExcelColumn.Index != 0)
                        {
                            rcm.ClientName = item.Column.Where(x => x.Index.Equals(xlsColClientName.ExcelColumn.Index)).Select(x => x.ExcelColumnName).FirstOrDefault().ToString();
                            if (rcm.ClientName != string.Empty)
                            {
                                rcm.ClientNameText = rcm.ClientName;
                                var client = _soxContext.ClientSs.Where(x => x.Name.Equals(rcm.ClientName)).FirstOrDefault();
                                if(client != null)
                                {
                                    rcm.ClientItemId = client.ClientItemId;
                                    rcm.ClientCode = client.ClientCode;
                                }
                            }
                        }
                           
                        if(rcm.ClientItemId != null && rcm.ClientItemId != 0)
                        {
                            var xlsColFy = importFields.ListDatabaseColumns.Where(x => x.DbColumnName.Equals("FY")).FirstOrDefault();
                            if (xlsColFy.ExcelColumn.Index != 0)
                            {
                                var Fy = item.Column.Where(x => x.Index.Equals(xlsColFy.ExcelColumn.Index)).Select(x => x.ExcelColumnName).FirstOrDefault();
                                rcm.FY = Fy ?? string.Empty;
                            }
                                

                            var xlsColProcess = importFields.ListDatabaseColumns.Where(x => x.DbColumnName.Equals("Process")).FirstOrDefault();
                            if (xlsColProcess.ExcelColumn.Index != 0)
                            {
                                var Process = item.Column.Where(x => x.Index.Equals(xlsColProcess.ExcelColumn.Index)).Select(x => x.ExcelColumnName).FirstOrDefault();
                                rcm.Process = Process ?? string.Empty;
                            }
                                

                            var xlsColSubprocess = importFields.ListDatabaseColumns.Where(x => x.DbColumnName.Equals("Subprocess")).FirstOrDefault();
                            if (xlsColSubprocess.ExcelColumn.Index != 0)
                            {
                                var Subprocess = item.Column.Where(x => x.Index.Equals(xlsColSubprocess.ExcelColumn.Index)).Select(x => x.ExcelColumnName).FirstOrDefault();
                                rcm.Subprocess = Subprocess ?? string.Empty;
                            }
                                

                            var xlsColControlObjective = importFields.ListDatabaseColumns.Where(x => x.DbColumnName.Equals("ControlObjective")).FirstOrDefault();
                            if (xlsColControlObjective.ExcelColumn.Index != 0)
                            {
                                var ControlObjective = item.Column.Where(x => x.Index.Equals(xlsColControlObjective.ExcelColumn.Index)).Select(x => x.ExcelColumnName).FirstOrDefault();
                                rcm.ControlObjective = ControlObjective ?? string.Empty;
                            }
                                

                            var xlsColSpecificRisk = importFields.ListDatabaseColumns.Where(x => x.DbColumnName.Equals("SpecificRisk")).FirstOrDefault();
                            if (xlsColSpecificRisk.ExcelColumn.Index != 0)
                            {
                                var SpecificRisk = item.Column.Where(x => x.Index.Equals(xlsColSpecificRisk.ExcelColumn.Index)).Select(x => x.ExcelColumnName).FirstOrDefault();
                                rcm.SpecificRisk = SpecificRisk ?? string.Empty;
                            }
                               

                            var xlsColFinStatemenElement = importFields.ListDatabaseColumns.Where(x => x.DbColumnName.Equals("FinStatemenElement")).FirstOrDefault();
                            if (xlsColFinStatemenElement.ExcelColumn.Index != 0)
                            {
                                var FinStatemenElement = item.Column.Where(x => x.Index.Equals(xlsColFinStatemenElement.ExcelColumn.Index)).Select(x => x.ExcelColumnName).FirstOrDefault();
                                rcm.FinStatemenElement = FinStatemenElement ?? string.Empty;
                            }
                                

                            var xlsColCompletenessAccuracy = importFields.ListDatabaseColumns.Where(x => x.DbColumnName.Equals("CompletenessAccuracy")).FirstOrDefault();
                            if (xlsColCompletenessAccuracy.ExcelColumn.Index != 0)
                            {
                                var CompletenessAccuracy = item.Column.Where(x => x.Index.Equals(xlsColCompletenessAccuracy.ExcelColumn.Index)).Select(x => x.ExcelColumnName).FirstOrDefault();
                                rcm.CompletenessAccuracy = CompletenessAccuracy ?? string.Empty;
                            }
                                

                            var xlsColExistenceDisclosure = importFields.ListDatabaseColumns.Where(x => x.DbColumnName.Equals("ExistenceDisclosure")).FirstOrDefault();
                            if (xlsColExistenceDisclosure.ExcelColumn.Index != 0)
                            {
                                var ExistenceDisclosure = item.Column.Where(x => x.Index.Equals(xlsColExistenceDisclosure.ExcelColumn.Index)).Select(x => x.ExcelColumnName).FirstOrDefault();
                                rcm.ExistenceDisclosure = ExistenceDisclosure ?? string.Empty;
                            }
                                

                            var xlsColPresentationDisclosure = importFields.ListDatabaseColumns.Where(x => x.DbColumnName.Equals("PresentationDisclosure")).FirstOrDefault();
                            if (xlsColPresentationDisclosure.ExcelColumn.Index != 0)
                            {
                                var PresentationDisclosure = item.Column.Where(x => x.Index.Equals(xlsColPresentationDisclosure.ExcelColumn.Index)).Select(x => x.ExcelColumnName).FirstOrDefault();
                                rcm.PresentationDisclosure = PresentationDisclosure ?? string.Empty;
                            }
                                

                            var xlsColRigthsObligation = importFields.ListDatabaseColumns.Where(x => x.DbColumnName.Equals("RigthsObligation")).FirstOrDefault();
                            if (xlsColRigthsObligation.ExcelColumn.Index != 0)
                            {
                                var RigthsObligation = item.Column.Where(x => x.Index.Equals(xlsColProcess.ExcelColumn.Index)).Select(x => x.ExcelColumnName).FirstOrDefault();
                                rcm.RigthsObligation = RigthsObligation ?? string.Empty;
                            }
                                

                            var xlsColValuationAlloc = importFields.ListDatabaseColumns.Where(x => x.DbColumnName.Equals("ValuationAlloc")).FirstOrDefault();
                            if (xlsColValuationAlloc.ExcelColumn.Index != 0)
                            {
                                var ValuationAlloc = item.Column.Where(x => x.Index.Equals(xlsColValuationAlloc.ExcelColumn.Index)).Select(x => x.ExcelColumnName).FirstOrDefault();
                                rcm.ValuationAlloc = ValuationAlloc ?? string.Empty;
                            }
                                

                            var xlsColControlId = importFields.ListDatabaseColumns.Where(x => x.DbColumnName.Equals("ControlId")).FirstOrDefault();
                            if (xlsColControlId.ExcelColumn.Index != 0)
                            {
                                var ControlId = item.Column.Where(x => x.Index.Equals(xlsColControlId.ExcelColumn.Index)).Select(x => x.ExcelColumnName).FirstOrDefault();
                                rcm.ControlId = ControlId ?? string.Empty;
                            }
                                

                            var xlsColControlActivityFy19 = importFields.ListDatabaseColumns.Where(x => x.DbColumnName.Equals("ControlActivityFy19")).FirstOrDefault();
                            if (xlsColControlActivityFy19.ExcelColumn.Index != 0)
                            {
                                var ControlActivityFy19 = item.Column.Where(x => x.Index.Equals(xlsColControlActivityFy19.ExcelColumn.Index)).Select(x => x.ExcelColumnName).FirstOrDefault();
                                rcm.ControlActivityFy19 = ControlActivityFy19 ?? string.Empty;
                            }
                                

                            var xlsColControlPlaceDate = importFields.ListDatabaseColumns.Where(x => x.DbColumnName.Equals("ControlPlaceDate")).FirstOrDefault();
                            if (xlsColControlPlaceDate.ExcelColumn.Index != 0)
                            {
                                var ControlPlaceDate = item.Column.Where(x => x.Index.Equals(xlsColProcess.ExcelColumn.Index)).Select(x => x.ExcelColumnName).FirstOrDefault();
                                rcm.ControlPlaceDate = ControlPlaceDate ?? string.Empty;
                            }
                                

                            var xlsColControlOwner = importFields.ListDatabaseColumns.Where(x => x.DbColumnName.Equals("ControlOwner")).FirstOrDefault();
                            if (xlsColControlOwner.ExcelColumn.Index != 0)
                            {
                                var ControlOwner = item.Column.Where(x => x.Index.Equals(xlsColControlOwner.ExcelColumn.Index)).Select(x => x.ExcelColumnName).FirstOrDefault();
                                rcm.ControlOwner = ControlOwner ?? string.Empty;
                            }
                                

                            var xlsColControlFrequency = importFields.ListDatabaseColumns.Where(x => x.DbColumnName.Equals("ControlFrequency")).FirstOrDefault();
                            if (xlsColControlFrequency.ExcelColumn.Index != 0)
                            {
                                var ControlFrequency = item.Column.Where(x => x.Index.Equals(xlsColControlFrequency.ExcelColumn.Index)).Select(x => x.ExcelColumnName).FirstOrDefault();
                                rcm.ControlFrequency = ControlFrequency ?? string.Empty;
                            }
                                

                            var xlsColKey = importFields.ListDatabaseColumns.Where(x => x.DbColumnName.Equals("Key")).FirstOrDefault();
                            if (xlsColKey.ExcelColumn.Index != 0)
                            {
                                var Key = item.Column.Where(x => x.Index.Equals(xlsColKey.ExcelColumn.Index)).Select(x => x.ExcelColumnName).FirstOrDefault();
                                rcm.Key = Key ?? string.Empty;
                            }
                                

                            var xlsColControlType = importFields.ListDatabaseColumns.Where(x => x.DbColumnName.Equals("ControlType")).FirstOrDefault();
                            if (xlsColControlType.ExcelColumn.Index != 0)
                            {
                                var ControlType = item.Column.Where(x => x.Index.Equals(xlsColControlType.ExcelColumn.Index)).Select(x => x.ExcelColumnName).FirstOrDefault();
                                rcm.ControlType = ControlType ?? string.Empty;
                            }
                                

                            var xlsColNatureProc = importFields.ListDatabaseColumns.Where(x => x.DbColumnName.Equals("NatureProc")).FirstOrDefault();
                            if (xlsColNatureProc.ExcelColumn.Index != 0)
                            {
                                var NatureProc = item.Column.Where(x => x.Index.Equals(xlsColNatureProc.ExcelColumn.Index)).Select(x => x.ExcelColumnName).FirstOrDefault();
                                rcm.NatureProc = NatureProc ?? string.Empty;
                            }
                                

                            var xlsColFraudControl = importFields.ListDatabaseColumns.Where(x => x.DbColumnName.Equals("FraudControl")).FirstOrDefault();
                            if (xlsColFraudControl.ExcelColumn.Index != 0)
                            {
                                var FraudControl = item.Column.Where(x => x.Index.Equals(xlsColFraudControl.ExcelColumn.Index)).Select(x => x.ExcelColumnName).FirstOrDefault();
                                rcm.FraudControl = FraudControl ?? string.Empty;
                            }
                                

                            var xlsColRiskLvl = importFields.ListDatabaseColumns.Where(x => x.DbColumnName.Equals("RiskLvl")).FirstOrDefault();
                            if (xlsColRiskLvl.ExcelColumn.Index != 0)
                            {
                                var RiskLvl = item.Column.Where(x => x.Index.Equals(xlsColRiskLvl.ExcelColumn.Index)).Select(x => x.ExcelColumnName).FirstOrDefault();
                                rcm.RiskLvl = RiskLvl ?? string.Empty;
                            }
                                

                            var xlsColManagementRevControl = importFields.ListDatabaseColumns.Where(x => x.DbColumnName.Equals("ManagementRevControl")).FirstOrDefault();
                            if (xlsColManagementRevControl.ExcelColumn.Index != 0)
                            {
                                var ManagementRevControl = item.Column.Where(x => x.Index.Equals(xlsColManagementRevControl.ExcelColumn.Index)).Select(x => x.ExcelColumnName).FirstOrDefault();
                                rcm.ManagementRevControl = ManagementRevControl ?? string.Empty;
                            }
                                

                            var xlsColPbcList = importFields.ListDatabaseColumns.Where(x => x.DbColumnName.Equals("PbcList")).FirstOrDefault();
                            if (xlsColPbcList.ExcelColumn.Index != 0)
                            {
                                var PbcList = item.Column.Where(x => x.Index.Equals(xlsColPbcList.ExcelColumn.Index)).Select(x => x.ExcelColumnName).FirstOrDefault();
                                rcm.PbcList = PbcList ?? string.Empty;
                            }
                                

                            var xlsColTestProc = importFields.ListDatabaseColumns.Where(x => x.DbColumnName.Equals("TestProc")).FirstOrDefault();
                            if (xlsColTestProc.ExcelColumn.Index != 0)
                            {
                                var TestProc = item.Column.Where(x => x.Index.Equals(xlsColTestProc.ExcelColumn.Index)).Select(x => x.ExcelColumnName).FirstOrDefault();
                                rcm.TestProc = TestProc ?? string.Empty;
                            }
                                

                            var xlsColTestingPeriod = importFields.ListDatabaseColumns.Where(x => x.DbColumnName.Equals("TestingPeriod")).FirstOrDefault();
                            if (xlsColTestingPeriod.ExcelColumn.Index != 0)
                            {
                                var TestingPeriod = item.Column.Where(x => x.Index.Equals(xlsColTestingPeriod.ExcelColumn.Index)).Select(x => x.ExcelColumnName).FirstOrDefault();
                                rcm.TestingPeriod = TestingPeriod ?? string.Empty;
                            }
                                

                            var xlsColPopulationFileRequired = importFields.ListDatabaseColumns.Where(x => x.DbColumnName.Equals("PopulationFileRequired")).FirstOrDefault();
                            if (xlsColPopulationFileRequired.ExcelColumn.Index != 0)
                            {
                                var PopulationFileRequired = item.Column.Where(x => x.Index.Equals(xlsColPopulationFileRequired.ExcelColumn.Index)).Select(x => x.ExcelColumnName).FirstOrDefault();
                                rcm.PopulationFileRequired = PopulationFileRequired ?? string.Empty;
                            }
                                

                            var xlsColMethodUsed = importFields.ListDatabaseColumns.Where(x => x.DbColumnName.Equals("MethodUsed")).FirstOrDefault();
                            if (xlsColMethodUsed.ExcelColumn.Index != 0)
                            {
                                var MethodUsed = item.Column.Where(x => x.Index.Equals(xlsColMethodUsed.ExcelColumn.Index)).Select(x => x.ExcelColumnName).FirstOrDefault();
                                rcm.MethodUsed = MethodUsed ?? string.Empty;
                            }
                                

                            var xlsColTestValidation = importFields.ListDatabaseColumns.Where(x => x.DbColumnName.Equals("TestValidation")).FirstOrDefault();
                            if (xlsColTestValidation.ExcelColumn.Index != 0)
                            {
                                var TestValidation = item.Column.Where(x => x.Index.Equals(xlsColTestValidation.ExcelColumn.Index)).Select(x => x.ExcelColumnName).FirstOrDefault();
                                rcm.TestValidation = TestValidation ?? string.Empty;
                            }
                                

                            var xlsColControlActivity = importFields.ListDatabaseColumns.Where(x => x.DbColumnName.Equals("ControlActivity")).FirstOrDefault();
                            if (xlsColControlActivity.ExcelColumn.Index != 0)
                            {
                                var ControlActivity = item.Column.Where(x => x.Index.Equals(xlsColControlActivity.ExcelColumn.Index)).Select(x => x.ExcelColumnName).FirstOrDefault();
                                rcm.ControlActivity = ControlActivity ?? string.Empty;
                            }

                            rcm.Status = "Active";
                            //start saving items in Podio
                            rcm = await rcmService.SavePodioRcm(rcm); //return podio item id if success
                            
                            if (rcm.PodioItemId != 0)
                            {
                                //save item in database
                                await rcmService.SaveRcmToDatabase(rcm);
                            }
                        }
                        else
                        {
                            Exception ex = new Exception($"Error: client item id not found. Row {item.Row} : Filename {importFields.Filename}");
                            FileLog.Write(ex.ToString(), "ErrorImportRcmAsync");
                            AdminService adminService = new AdminService(_config);
                            adminService.SendAlert(true, true, ex.ToString(), "ImportFileAsync");
                        }



                        //writeLog.Display(rcm);

                    }



                    






                    return Ok(listRowVal);
                }





            }
            catch (Exception ex)
            {
                FileLog.Write($"Error ImportFileAsync {ex}", "ErrorImportFileAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "ImportFileAsync");
                return BadRequest();
            }

        }

        [AllowAnonymous]
        [HttpPost("trigger")]
        public async Task<IActionResult> TriggerJob()
        {
            ITrigger trigger = TriggerBuilder.Create()
                .WithIdentity($"{DateTime.Now}-SODSoxRox")
                .StartNow()
                .WithPriority(1)
                .Build();

            IJobDetail job = JobBuilder.Create<SodSoxRoxJob>()
                .WithIdentity($"{DateTime.Now}-SODSoxRox")                
                .Build();
            job.JobDataMap["object"] = "";

            await _scheduler.ScheduleJob(job, trigger);

            return Ok();
        }

        [AllowAnonymous]
        [HttpPost("listjobs")]
        public async Task<IActionResult> ListRunningJobs()
        {

            var listJobs = await _scheduler.GetCurrentlyExecutingJobs();

            return Ok(listJobs.ToList());
        }


        [AllowAnonymous]
        [HttpPost("compress")]
        public IActionResult CompressTest(string filename1, string filename2)
        {
            try
            {
                AdminService adminService = new AdminService(_config);
                Debug.WriteLine($"Compress files | {DateTime.Now}");
                #region Compress files
                string startupPath = Directory.GetCurrentDirectory();
                string strSourceDownload = Path.Combine(startupPath, "include", "sod");
                //string strOutput1 = Path.Combine(strSourceDownload, filename1);
                //string strOutput2 = Path.Combine(strSourceDownload, filename2);
                List<string> listSodReport = new List<string>();
                listSodReport.Add(filename1);
                listSodReport.Add(filename2);
                var taskCompressReportFile = System.Threading.Tasks.Task.Run(() => adminService.CompressItem(strSourceDownload, "SodSoxRoxReport", listSodReport));
                System.Threading.Tasks.Task.WhenAll(taskCompressReportFile).Wait();
                string zipFileName = taskCompressReportFile.Result;
                Debug.WriteLine($"DONE Compress files | {DateTime.Now}");
                #endregion


                return Ok(zipFileName);
            }
            catch (Exception ex)
            {

                return BadRequest(ex);
            }
            
            
        }

        private List<string> ImportExcludeDBColumn()
        {
            List<string> listExclude = new List<string>();
            listExclude.AddRange(new List<string>{
                "Id",
                "ClientCode",
                "ClientItemId",
                "ClientNameText",
                "PodioItemId",
                "PodioUniqueId",
                "PodioRevision",
                "PodioLink",
                "CreatedBy",
                "SharefileLink",
                "JsonData",
                "CreatedOn",
                "Status"
            });


            return listExclude;
        }

        //Dictionary mime use for download files
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
                {".mp4", "video/mp4"},
            };
        }

        private async Task<(List<ColumnVal>, List<DBColumnVal>)> RcmFileUpload(string path)
        {
            FileInfo fi = new FileInfo(path);
            List<ColumnVal> columnField = new List<ColumnVal>();
            List<DBColumnVal> dbColumn = new List<DBColumnVal>();
            using (ExcelPackage p = new ExcelPackage(fi))
            {
                // If you use EPPlus in a noncommercial context
                // according to the Polyform Noncommercial license:
                //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                //Console.WriteLine($"Worksheet Count: {p.Workbook.Worksheets.Count}");
                ExcelWorksheet ws = p.Workbook.Worksheets[0];
                int colCount = ws.Dimension.End.Column;
                int rowCount = ws.Dimension.End.Row;

                for (int row = 1; row <= 1; row++)
                {
                    for (int col = 1; col <= colCount; col++)
                    {
                        ColumnVal column = new ColumnVal();
                        column.Index = col;
                        column.ExcelColumnName = ws.Cells[row, col].Value?.ToString();
                        columnField.Add(column);
                    }
                }
            }
            List<string> listExclude = new List<string>();
            listExclude = ImportExcludeDBColumn();
            string[] table;
            int position = 1;
            
            table = typeof(Rcm).GetProperties().Select(property => property.Name).ToArray();
            foreach (var item in table)
            {
                DBColumnVal column = new DBColumnVal();
                ColumnVal xlsColumn = new ColumnVal();
                if (!listExclude.Contains(item))
                {
                    column.DbColumnName = item;
                    column.Position = position;
                    xlsColumn.Index = 0;
                    xlsColumn.ExcelColumnName = string.Empty;
                    column.ExcelColumn = xlsColumn;
                    dbColumn.Add(column);
                    position++;
                }
            }
            await System.Threading.Tasks.Task.Delay(0);
            return (columnField, dbColumn);
        }

        private async Task<(List<ColumnVal>, List<DBColumnVal>)> KeyReportFileUpload(string path)
        {
            FileInfo fi = new FileInfo(path);
            List<ColumnVal> columnField = new List<ColumnVal>();
            List<DBColumnVal> dbColumn = new List<DBColumnVal>();
            using (ExcelPackage p = new ExcelPackage(fi))
            {
                // If you use EPPlus in a noncommercial context
                // according to the Polyform Noncommercial license:
                //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                //Console.WriteLine($"Worksheet Count: {p.Workbook.Worksheets.Count}");
                ExcelWorksheet ws = p.Workbook.Worksheets[0];
                int colCount = ws.Dimension.End.Column;
                int rowCount = ws.Dimension.End.Row;

                for (int row = 1; row <= 1; row++)
                {
                    for (int col = 1; col <= colCount; col++)
                    {
                        ColumnVal column = new ColumnVal();
                        column.Index = col;
                        column.ExcelColumnName = ws.Cells[row, col].Value?.ToString();
                        columnField.Add(column);
                    }
                }
            }
            List<string> listExclude = new List<string>();
            listExclude = ImportExcludeDBColumn();
            string[] table;
            int position = 1;

            table = typeof(Rcm).GetProperties().Select(property => property.Name).ToArray();
            foreach (var item in table)
            {
                DBColumnVal column = new DBColumnVal();
                ColumnVal xlsColumn = new ColumnVal();
                if (!listExclude.Contains(item))
                {
                    column.DbColumnName = item;
                    column.Position = position;
                    xlsColumn.Index = 0;
                    xlsColumn.ExcelColumnName = string.Empty;
                    column.ExcelColumn = xlsColumn;
                    dbColumn.Add(column);
                    position++;
                }
            }
            await System.Threading.Tasks.Task.Delay(0);
            return (columnField, dbColumn);
        }

    }
}
