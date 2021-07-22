using A2B_App.Server.Data;
using A2B_App.Server.Services;
using A2B_App.Shared.Sox;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ShareFile.Api.Client;
using ShareFile.Api.Client.Extensions;
using ShareFile.Api.Client.Logging;
using ShareFile.Api.Client.Models;
using ShareFile.Api.Client.Security.Authentication.OAuth2;
using ShareFile.Api.Client.Transfers;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace A2B_App.Server.Controllers
{
    [Authorize]
    [Route("api/[controller]")]
    [ApiController]
    public class ShareFileController : ControllerBase
    {
        private readonly IConfiguration _config;
        private readonly ILogger<ShareFileController> _logger;
        private readonly IWebHostEnvironment _environment;
        private readonly SoxContext _soxContext;

        public ShareFileController(IConfiguration config,
            ILogger<ShareFileController> logger,
            IWebHostEnvironment environment,
            SoxContext soxContext)
        {
            _config = config;
            _logger = logger;
            _environment = environment;
            _soxContext = soxContext;
        }


        //[AllowAnonymous]
        [HttpPost("upload")]
        public async Task<IActionResult> SharefileUploadAsync([FromBody] SharefileItem sharefileItem)
        {
            try
            {
                Debug.WriteLine("Authentication started");
                Session session = null;
                ShareFileClient sfClient = null;
                SharefileUser user = new SharefileUser
                {
                    ControlPlane = _config.GetSection("SharefileApi").GetSection("ControlPane").Value,
                    Username = _config.GetSection("SharefileApi").GetSection("Username").Value,
                    Password = _config.GetSection("SharefileApi").GetSection("Password").Value,
                    Subdomain = _config.GetSection("SharefileApi").GetSection("SubDomain").Value
                };


                string oauthClientId = _config.GetSection("SharefileApi").GetSection("ClientId").Value;
                string oauthClientSecret = _config.GetSection("SharefileApi").GetSection("ClientSecret").Value;

                // Authenticate with username/password
                sfClient = await PasswordAuthentication(user, oauthClientId, oauthClientSecret);

                //Start session
                session = await sfClient.Sessions.Login().Expand("Principal").ExecuteAsync();

                string startupPath = Directory.GetCurrentDirectory();
                string path = string.Empty;

                switch (sharefileItem.Directory)
                {
                    case "SoxRcmFolder":
                        path = Path.Combine(startupPath, "include", "upload", "rcm", sharefileItem.FileName);
                        break;
                    case "SoxTrackerFolder":
                        path = Path.Combine(startupPath, "include", "upload", "soxtracker", sharefileItem.FileName);
                        break;
                    case "SoxQuestionnaireFolder":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxSampleSelectionFolder":
                        path = Path.Combine(startupPath, "include", "upload", "sampleselection", sharefileItem.FileName);
                        break;

                    case "SoxKeyReportLeadSheetFolder":
                        path = Path.Combine(startupPath, "include", "upload", "keyreport","leadsheet", sharefileItem.FileName);
                        break;

                    case "SoxKeyReportLeadSheetFolderPWC"://062921
                        path = Path.Combine(startupPath, "include", "upload", "keyreport", "pwc", sharefileItem.FileName);
                        break;

                    case "SoxKeyReportTrackerFolder":
                        path = Path.Combine(startupPath, "include", "upload", "keyreport", "kr tracker", sharefileItem.FileName);
                        break;
                    //Workpaper SF eri
                    case "SoxWP_ERI_ELC":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_ERI_HRP":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_ERI_FCR":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_ERI_IA":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_ERI_INV":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_ERI_ITGC":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_ERI_OTC":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_ERI_PTP":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_ERI_TAX":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_ERI_TCM":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_ERI_FA":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_ERI_ESA":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;

                    //Workpaper SF ViewRay
                    case "SoxWP_VR_ESA":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_VR_TCM":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_VR_TAX":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_VR_FCR":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_VR_ELC":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_VR_PTP":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_VR_HRP":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_VR_FA":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_VR_INV":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_VR_ITGC":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_VR_OTC":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;

                    //Workpaper SF Kindred
                    case "SoxWP_KB_TCM":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_KB_TAX":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_KB_REV":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_KB_PTP":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_KB_MEC":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_KR_ITGC":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_KB_INV":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_KB_HRP":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_KB_FRE":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_KB_FA":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_KB_ELC":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_KB_ESA":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_KB_CTA":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    //Workpaper SF ChargePoint
                    case "SoxWP_CP_PTP":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_CP_FA":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_CP_ELC":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_CP_ITGC":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_CP_HRP":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_CP_ESA":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_CP_FCR":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_CP_OTC":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_CP_INV":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_CP_TAX":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_CP_TCM":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    //Workpaper SF Techpoint
                    case "SoxWP_TP_INV":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_TP_FCR":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_TP_HRP":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_TP_ELC":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_TP_FA":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_TP_TAX":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_TP_ESA":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_TP_TCM":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_TP_PTP":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_TP_ITGC":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_TP_OTC":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    //Workpaper SF McGrath
                    case "SoxWP_MG_ELC":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_MG_INV":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_MG_FSCP":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_MG_OTC":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_MG_FA":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_MG_ESA":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_MG_PTP":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_MG_TCM":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_MG_TAX":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    //Workpaper SF Fastly
                    case "SoxWP_FAS_MA":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_FAS_TCM":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_FAS_TAX":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_FAS_PTP":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_FAS_OTC":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_FAS_HRP":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_FAS_FCR":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_FAS_FA":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_FAS_ESA":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_FAS_ELC":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    //Workpaper SF Cortexyme
                    case "SoxWP_COR_CMC":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_COR_TCM":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_COR_TAX":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_COR_PTP":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_COR_ITGC":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_COR_HRP":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_COR_FCR":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_COR_FA":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_COR_ESA":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_COR_ELC":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_COR_CTA":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    //Workpaper SF Ambarella
                    case "SoxWP_AMB_QTC":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_AMB_PTP":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_AMB_ITGC":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                   
                    case "SoxWP_AMB_INV":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_AMB_HRP":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_AMB_FCR":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;
                    case "SoxWP_AMB_ELC":
                        path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", sharefileItem.FileName);
                        break;

                   

                    default:
                        break;
                }      

                string filePath = path;
                //string filename = filePath.Split(Path.DirectorySeparatorChar).Last();
                string filename = sharefileItem.FileName;
                var fileExtension = filename.Split('.').Last();
                var fileNameOnly = filename.Split('.').First();
                

                string sfDirectory = _config.GetSection("SharefileApi").GetSection(sharefileItem.Directory).GetSection("Path").Value;
                string sfLink = _config.GetSection("SharefileApi").GetSection(sharefileItem.Directory).GetSection("Link").Value;
                //await ConsoleOut("System", "Uploading file - " + filename, colorBlack, consolas);
                Folder folder = (Folder)await sfClient.Items.ByPath(sfDirectory).ExecuteAsync();

                //var uploadedFileId = await UploadFile(sfClient, folder, filePath, fileExtension, null, fileNameOnly);
                var uploadedFileId = await UploadFile(sfClient, folder, filePath, null, null, filename);//032921
                await UploadFile(sfClient, folder, filePath, null, null, filename); // 0326
                var itemUri = sfClient.Items.GetAlias(uploadedFileId);
                var uploadedFile = await sfClient.Items.Get(itemUri).ExecuteAsync();
                Debug.WriteLine($"Successfully uploaded {uploadedFile}");

                ////Get ShareFolder URI
                //var file = await sfClient.Items.Get(uploadedFile.url).ExecuteAsync();
                //var share = await sfService.ShareViaLink(sfClient, file);
                //await ConsoleOut("System", "Share URI - " + share.Uri.ToString(), colorBlack, consolas);

                //InvFileList.InvFileList[i].ShareLink = share.Uri.ToString();

                return Ok(sfLink);
            }
            catch (Exception ex)
            {
                FileLog.Write($"Error SharefileUploadAsync {ex}", "ErrorSharefileUploadAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SharefileUploadAsync");
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
        [HttpPost("screenshot")]
        public async Task<IActionResult> SharefileDownloadLeadsheetScreenshotAsync([FromBody] KeyReportFilter filter)
        {
            try
            {
                KeyReportScreenshot keyrepScreenshot = new KeyReportScreenshot();
                keyrepScreenshot.Filter = filter;
                List<string> listScreenshotName = new List<string>();
                Debug.WriteLine("Authentication started");
                Session session = null;
                ShareFileClient sfClient = null;
                SharefileUser user = new SharefileUser
                {
                    ControlPlane = _config.GetSection("SharefileApi").GetSection("ControlPane").Value,
                    Username = _config.GetSection("SharefileApi").GetSection("Username").Value,
                    Password = _config.GetSection("SharefileApi").GetSection("Password").Value,
                    Subdomain = _config.GetSection("SharefileApi").GetSection("SubDomain").Value
                };

                string oauthClientId = _config.GetSection("SharefileApi").GetSection("ClientId").Value;
                string oauthClientSecret = _config.GetSection("SharefileApi").GetSection("ClientSecret").Value;

                // Authenticate with username/password
                sfClient = await PasswordAuthentication(user, oauthClientId, oauthClientSecret);

                //Start session
                session = await sfClient.Sessions.Login().Expand("Principal").ExecuteAsync();

                if(session.IsAuthenticated)
                {
                    //Folder folder = (Folder)await sfClient.Items.ByPath(sfDirectory).ExecuteAsync();
                    //Uri uri = new Uri("https://a2q2.sharefile.com/home/shared/fo787484-9f3d-4ac7-ab68-0db7fa45e06c");
                    //var sfItems = await sfClient.Items.Get(sfClient.Items.GetEntityUriFromId("fo787484-9f3d-4ac7-ab68-0db7fa45e06c"), false).Expand("Children").ExecuteAsync();

                    //get folder id from sox context ClientSs
                    var clientSharefile = _soxContext.ClientSs.Where(x => x.Name.ToLower().Equals(filter.ClientName.ToLower())).Select(x => new { x.SharefileScreenshotId, x.SharefileReportId }).FirstOrDefault();
                    var folderId = clientSharefile.SharefileScreenshotId;
                    var folderReportId = clientSharefile.SharefileReportId;
                    //string folderId = "fo787484-9f3d-4ac7-ab68-0db7fa45e06c";
                    if (folderId != null)
                    {
                        #region Screenshot
                        var sfFolders = await sfClient.Items.GetChildren(sfClient.Items.GetEntityUriFromId(folderId), false).ExecuteAsync();
                        if (sfFolders != null)
                        {
                            //check if folder name is equal to key report name
                            var checkFolder = sfFolders.Feed.Where(x => x.FileName.ToLower().Equals(filter.KeyReportName.ToLower())).FirstOrDefault();
                            if (checkFolder != null)
                            {
                                Debug.WriteLine($"Folder Id: {checkFolder.Id}");

                                //get all items for this folder
                                var sfItems = await sfClient.Items.GetChildren(sfClient.Items.GetEntityUriFromId(checkFolder.Id), false).ExecuteAsync();
                                if (sfItems != null)
                                {
                                    foreach (var item in sfItems.Feed)
                                    {
                                        Debug.WriteLine($"Item Id: {item.FileName}");
                                        //check file if exists                                         
                                        if(!CheckFileIfExists(filter, item.FileName, item.FileSizeBytes.Value))
                                        {
                                            //if not exists then we download the file, this will return true on success
                                            if(await Download(sfClient, item, filter))
                                            {
                                                //add to list screenshot
                                                listScreenshotName.Add(item.FileName);
                                            }
                                        }
                                        else
                                        {
                                            //add to list screenshot
                                            listScreenshotName.Add(item.FileName);
                                        }
                                    }

                                    //set value to list keyreport screenshot
                                    keyrepScreenshot.ListScreenshotName = listScreenshotName.OrderBy(x => x).ToList();
                                }
  
                            }

                        }
                        #endregion

                        #region Report
                        var sfReportFolders = await sfClient.Items.GetChildren(sfClient.Items.GetEntityUriFromId(folderReportId), false).ExecuteAsync();
                        if (sfReportFolders != null)
                        {

                            foreach (var item in sfReportFolders.Feed)
                            {
                                Debug.WriteLine($"Report Filename: {item.FileName}");
                                Debug.WriteLine($"Report Name: {item.Name}");
                                
                                string extension = item.FileName.Split(".").Last();
                                string filename = item.FileName.Replace($".{extension}", "");

                                if (filename.Equals(filter.KeyReportName) && (extension.Equals("xlsx") || extension.Equals("xls")))
                                {
                                    //check file if exists                                         
                                    if (!CheckFileIfExists(filter, item.FileName, item.FileSizeBytes.Value))
                                    {
                                        //if not exists then we download the file, this will return true on success
                                        //file location @ startupPath, "include", "sharefile", "download", $"{filter.ClientName}", $"{filter.KeyReportName}
                                        if (await Download(sfClient, item, filter))
                                        {
                                            //Add report file
                                            keyrepScreenshot.ReportFilename = item.FileName;
                                        }
                                    }
                                    else
                                    {
                                        //Add report file
                                        keyrepScreenshot.ReportFilename = item.FileName;
                                    }
                                }

                            }

                            //set value to list keyreport screenshot


                        }


                        #endregion

                    }


                }

                if(keyrepScreenshot.ListScreenshotName != null && keyrepScreenshot.ListScreenshotName.Count > 0 || keyrepScreenshot.ReportFilename != string.Empty)
                {
                    //return KeyReportScreenshot model
                    return Ok(keyrepScreenshot);
                }

                return NoContent();

            }
            catch (Exception ex)
            {
                FileLog.Write($"Error SharefileDownloadLeadsheetScreenshotAsync {ex}", "ErrorSharefileDownloadLeadsheetScreenshotAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SharefileDownloadLeadsheetScreenshotAsync");
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

        private async Task<ShareFileClient> PasswordAuthentication(SharefileUser SFUser, string ClientID, string ClientSecret)
        {
            // Initialize ShareFileClient.
            var configuration = Configuration.Default();
            configuration.Logger = new DefaultLoggingProvider();

            var sfClient = new ShareFileClient("https://secure.sf-api.com/sf/v3/", configuration);
            var oauthService = new OAuthService(sfClient, ClientID, ClientSecret);

            // Perform a password grant request.  Will give us an OAuthToken
            var oauthToken = await oauthService.PasswordGrantAsync(SFUser.Username, SFUser.Password, SFUser.Subdomain, SFUser.ControlPlane);

            // Add credentials and update sfClient with new BaseUri
            sfClient.AddOAuthCredentials(oauthToken);
            sfClient.BaseUri = oauthToken.GetUri();
            //Console.WriteLine("Sharefile authenticated");

            return sfClient;
        }

        public static async Task<bool> Download(ShareFileClient sfClient, Item itemToDownload, KeyReportFilter filter)
        {
            try
            {
                string result = string.Empty;
                string startupPath = Directory.GetCurrentDirectory();
                string path = Path.Combine(startupPath, "include", "sharefile", "download", $"{filter.ClientName}", $"{filter.KeyReportName}");

                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                var downloader = sfClient.GetAsyncFileDownloader(itemToDownload);
                var file = System.IO.File.Open(Path.Combine(path, itemToDownload.Name), FileMode.Append);

                decimal percent = 0;
                downloader.OnTransferProgress += (sender, e) =>
                {
                    //Download progress
                    percent = (((decimal)(e.Progress.BytesTransferred) / (decimal)(e.Progress.TotalBytes)) * 100);
                    Debug.WriteLine("Downloading " + e.Progress.BytesTransferred + "/" + e.Progress.TotalBytes + " - " + percent.ToString("0.##") + "%");

                    //Download completed
                    if (e.Progress.Complete)
                    {
                        Debug.WriteLine("Download complete - " + itemToDownload.Name);
                        //result = startupPath + @"\files\" + ItemToDownload.Name;
                        file.Dispose();
                    }
                };

                await downloader.DownloadToAsync(file);
                return true;
            }
            catch (Exception ex)
            {
                FileLog.Write($"Error SharefileDownloadScreenshotAsync {ex}", "ErrorSharefileDownloadScreenshotAsync");
                return false;
                //throw;
            }
            
        }

        /// <summary>
        /// Check if file already exists in directory
        /// </summary>
        /// <param name="filter"></param>
        /// <param name="filename"></param>
        /// <param name="fileSize"></param>
        /// <returns>return true if exists, else false</returns>
        private bool CheckFileIfExists(KeyReportFilter filter, string filename, long fileSize)
        {
            try
            {
                string startupPath = Directory.GetCurrentDirectory();
                string path = Path.Combine(startupPath, "include", "sharefile", "download", $"{filter.ClientName}", $"{filter.KeyReportName}", filename);
                FileInfo file = new FileInfo(path);
                if (file != null && file.Exists)
                {
                    if (file.Length.Equals(fileSize))
                    {
                        Debug.WriteLine($"Item {filename} already exists");
                        return true;
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                return false;
                //throw;
            }
            
        }

        private async Task<string> UploadFile(
            ShareFileClient sfClient,
            Folder destinationFolder,
            string FilePath,
            string FileExtension,
            string FileDetails,
            string RecordName)
        {
            string result = string.Empty;
            var file = System.IO.File.Open(FilePath, FileMode.OpenOrCreate);
            var uploadRequest = new UploadSpecificationRequest
            {
                FileName = RecordName + @"." + FileExtension,
                FileSize = file.Length,
                Details = FileDetails,
                Parent = destinationFolder.url
            };

            var uploader = sfClient.GetAsyncFileUploader(uploadRequest, file);
            decimal percent = 0;
            var uploadResponse = await uploader.UploadAsync();

            uploader.OnTransferProgress += (sender, e) =>
            {
                //Download progress
                percent = (((decimal)(e.Progress.BytesTransferred) / (decimal)(e.Progress.TotalBytes)) * 100);
                Debug.WriteLine("Uploading " + e.Progress.BytesTransferred + "/" + e.Progress.TotalBytes + " - " + percent.ToString("0.##") + "%");

                //Download completed
                if (e.Progress.Complete)
                {
                    Debug.WriteLine("Upload complete - " + file.Name);
                    file.Dispose();
                }
            };

            result = uploadResponse.First().Id;
            return result;

        }



    }


}

                                                   