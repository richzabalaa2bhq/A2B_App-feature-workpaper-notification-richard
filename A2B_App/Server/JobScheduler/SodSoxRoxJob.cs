using A2B_App.Server.Data;
using A2B_App.Server.Services;
using A2B_App.Shared.Sox;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Quartz;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Threading.Tasks;

namespace A2B_App.Server.JobScheduler
{
    public class SodSoxRoxJob : IJob
    {
        //private readonly IConfiguration _config;
        //private readonly ILogger<SodSoxRoxTask> _logger;
        //private readonly IWebHostEnvironment _environment;
        //private readonly SoxContext _soxContext;
        //public SodSoxRoxTask(IConfiguration config,
        //    ILogger<SodSoxRoxTask> logger,
        //    IWebHostEnvironment environment,
        //    SoxContext soxContext)
        //{
        //    _config = config;
        //    _logger = logger;
        //    _environment = environment;
        //    _soxContext = soxContext;
        //}

        public Task Execute(IJobExecutionContext context)
        {

            string emailBody = string.Empty;
            string emailCc = string.Empty;
            var listSoxFile = context.JobDetail.JobDataMap["object"] as List<SoxRoxFile>;
            var _soxContext = context.JobDetail.JobDataMap["context"] as SoxContext;
            var _config = context.JobDetail.JobDataMap["config"] as IConfiguration;
            var requestedBy = context.JobDetail.JobDataMap["requestedBy"] as string;
            SodService sodService = new SodService(_soxContext, _config);
            AdminService adminService = new AdminService(_config);


            try
            {
                emailCc = _config.GetSection("Email").GetSection("SODSoxRox").GetSection("EmailCc").Value;
                Debug.WriteLine($"Starting SOD SoxRox Job | {DateTime.Now}");
                

                SodSoxRoxInput response = new SodSoxRoxInput();
                string clientName = listSoxFile[0].ClientName;



                Debug.WriteLine($"Reading SOD SoxRox File | {DateTime.Now}");
                #region Reading SOD SoxRox File
                //run simultaneously
                var taskRoleUser = Task.Run(() => sodService.ReadFileSodSoxRoxRoleUser(listSoxFile, clientName));
                var taskRolePerm = Task.Run(() => sodService.ReadFileSodSoxRoxRolePerm(listSoxFile, clientName));
                var taskDescPerm = Task.Run(() => sodService.ReadFileSodSoxRoxDescToPerm(listSoxFile, clientName));
                var taskConflict = Task.Run(() => sodService.ReadFileSodSoxRoxConflictPerm(listSoxFile, clientName));

                Task.WhenAll(taskRoleUser, taskRolePerm, taskDescPerm, taskConflict).Wait();
                //Task.WhenAll(taskDescPerm, taskConflict);

                response.ListRoleUser = taskRoleUser.Result;
                response.ListRolePerm = taskRolePerm.Result;
                response.ListDescriptionToPerm = taskDescPerm.Result;
                response.ListConflictPerm = taskConflict.Result;
                response.ListRoleUserTrim = sodService.GetSodSoxRoxRoleUser(response);
                Debug.WriteLine($"DONE Reading SOD SoxRox File | {DateTime.Now}");
                #endregion



                Debug.WriteLine($"Processing SOD SoxRox Analysis | {DateTime.Now}");
                #region Processing SOD SoxRox Analysis
                var sodSoxRoxRaw2 = Task.Run(() => sodService.ProcessSoxRoxDataRaw2_3(response));
                var sodSoxRoxRaw3 = Task.Run(() => sodService.ProcessSoxRoxDataRaw3_3(response));
                var taskCreateDescPerm = Task.Run(() => sodService.ProcessSoxRoxDescriptionOutput(response));
                Task.WhenAll(sodSoxRoxRaw2, sodSoxRoxRaw3, taskCreateDescPerm).Wait();
                Debug.WriteLine($"DONE Processing SOD SoxRox Analysis | {DateTime.Now}");
                #endregion



                Debug.WriteLine($"Creating SOD SoxRox Report | {DateTime.Now}");
                #region Creating SOD SoxRox Report
                //create excel
                var taskCreateDescriptionOutput = Task.Run(() => sodService.CreateSoxRoxDescriptionFile(taskCreateDescPerm.Result, clientName));
                var taskCreatereport = Task.Run(() => sodService.CreateSodSoxRoxFile(sodSoxRoxRaw2.Result, sodSoxRoxRaw3.Result, clientName));
                Task.WhenAll(taskCreateDescriptionOutput, taskCreatereport).Wait();

                SodSoxRoxOutputFile sodSoxRoxOutputFile = new SodSoxRoxOutputFile();
                sodSoxRoxOutputFile.ReportFileName = taskCreatereport.Result;
                sodSoxRoxOutputFile.DescriptionFileName = taskCreateDescriptionOutput.Result;
                Debug.WriteLine($"DONE Creating SOD SoxRox Report | {DateTime.Now}");
                #endregion


                Debug.WriteLine($"Compress files | {DateTime.Now}");
                #region Compress files
                string startupPath = Directory.GetCurrentDirectory();
                string strSourceDownload = Path.Combine(startupPath, "include", "sod");
                string strOutput1 = Path.Combine(strSourceDownload, sodSoxRoxOutputFile.ReportFileName);
                string strOutput2 = Path.Combine(strSourceDownload, sodSoxRoxOutputFile.DescriptionFileName);
                List<string> listSodReport = new List<string>();
                listSodReport.Add(sodSoxRoxOutputFile.ReportFileName);
                listSodReport.Add(sodSoxRoxOutputFile.DescriptionFileName);
                var taskCompressReportFile = Task.Run(() => adminService.CompressItem(strSourceDownload, $"{clientName}_SODReport", listSodReport));
                Task.WhenAll(taskCompressReportFile).Wait();
                string zipFileName = taskCompressReportFile.Result;
                Debug.WriteLine($"DONE Compress files | {DateTime.Now}");
                #endregion



                Debug.WriteLine($"Upload to Sharefile | {DateTime.Now}");
                #region Upload to Sharefile

                if (zipFileName != string.Empty)
                {
                    SharefileItem sfItem = new SharefileItem();
                    sfItem.FileName = zipFileName;
                    sfItem.FilePath = Path.Combine(strSourceDownload, zipFileName);
                    sfItem.Directory = "SoxSodFolder";

                    ShareFileService sfService = new ShareFileService(_config);
                    var taskUploadToSF = Task.Run(() => sfService.UploadWithUrlReturn(sfItem));
                    Task.WhenAll(taskUploadToSF).Wait();
                    string url = taskUploadToSF.Result;
                    Debug.WriteLine($"DONE Upload to Sharefile {url}| {DateTime.Now}");


                    Debug.WriteLine($"Send email to requestor | {DateTime.Now}");
                    #region Send email to requestor
                    string body = adminService.SodSoxRoxEmailBody(url, zipFileName).Result;
                    adminService.SendEmail("SOD SoxRox Report", body, requestedBy, emailCc);
                    Debug.WriteLine($"DONE Send email to requestor | {DateTime.Now}");
                    #endregion

                }

                #endregion




            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write($"Error SodSoxRoxTask {ex}", "ErrorSodSoxRoxTask");

                adminService.SendAlert(true, true, ex.ToString(), "UploadFileForSodSoxRoxAsync");
                Debug.WriteLine($"Send email to requestor | {DateTime.Now}");
                adminService.SendEmail("SOD SoxRox Error", ex.ToString(), requestedBy, emailCc);
                Debug.WriteLine($"DONE Send email to requestor | {DateTime.Now}");
            }

            return Task.FromResult(0);
        
        }
    }
}
