using A2B_App.Server.Data;
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
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using PodioAPI;
using PodioAPI.Models;
using PodioAPI.Utils;
using PodioAPI.Utils.ItemFields;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading.Tasks;


namespace A2B_App.Server.Controllers
{
    [Authorize]
    [Route("api/[controller]")]
    [ApiController]

    public class ClientUIProcess : Controller
    {
        private readonly IConfiguration _config;
        private readonly ILogger<ClientUIProcess> _logger;
        private readonly IWebHostEnvironment _environment;
        private readonly SoxContext _soxContext;

        public ClientUIProcess(IConfiguration config,
            ILogger<ClientUIProcess> logger,
            IWebHostEnvironment environment,
            SoxContext soxContext)
        {
            _config = config;
            _logger = logger;
            _environment = environment;
            _soxContext = soxContext;
        }

        [AllowAnonymous]
        [HttpPost("create")]
        public async Task<IActionResult> CreateIteamPodio([FromBody] ClientSs clients)
        {
            PodioApiKey PodioKey = new PodioApiKey();
            PodioAppKey PodioAppKey = new PodioAppKey();
            PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
            PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;

            PodioAppKey.AppId = _config.GetSection("A2BHQClient").GetSection("ClientAppId").Value;
            PodioAppKey.AppToken = _config.GetSection("A2BHQClient").GetSection("ClientAppToken").Value;

            var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

            using (var context = _soxContext.Database.BeginTransaction())
            {
                var check = _soxContext.ClientSs.Where(x => x.ClientName.Equals(clients.ClientName) && x.ClientCode.Equals(clients.ClientCode)).FirstOrDefault();
                if (check != null)
                {
                    //update
                    if (podio.IsAuthenticated())
                    {

                        Item MyItem1 = new Item();
                        MyItem1.ItemId = (int)check.ClientItemId;

                        var ClientName = MyItem1.Field<TextItemField>("cliend-reference-code");
                        ClientName.Value = clients.ClientName;

                        var ClientCode = MyItem1.Field<TextItemField>("unique-identifier-code");
                        ClientCode.Value = clients.ClientCode;

                        var ClientCodeReferrence = await podio.ItemService.UpdateItem(MyItem1, null, null, false, true);

                        PodioAppKey.AppId = _config.GetSection("SampleSelectionPodioApp").GetSection("ClientAppId").Value;
                        PodioAppKey.AppToken = _config.GetSection("SampleSelectionPodioApp").GetSection("ClientAppToken").Value;

                        var podio2 = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                        ServicePointManager.Expect100Continue = true;
                        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                        await podio2.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                        if (podio2.IsAuthenticated())
                        {

                            var data = clients;
                            Item MyItem = new Item();

                            MyItem.ItemId = check.PodioItemId;

                            var Title = MyItem.Field<TextItemField>("title");
                            Title.Value = data.ClientName;

                            int refId = (int)check.ClientItemId;

                            var ReferrenceId = MyItem.Field<AppItemField>("client-relationship");
                            ReferrenceId.ItemIds = new List<int> { refId };

                            var External = MyItem.Field<TextItemField>("external-auditor");
                            External.Value = data.ExternalAuditor;

                            var Round1Percent = MyItem.Field<NumericItemField>("round-1-percent");
                            Round1Percent.Value = data.Percent;

                            var Round2Percent = MyItem.Field<NumericItemField>("round-2-percent");
                            Round2Percent.Value = data.PercentRound2;

                            var SharefileIdSave = MyItem.Field<TextItemField>("sharefile-id-save-file");
                            SharefileIdSave.Value = data.SharefileId;

                            var KeyReportID = MyItem.Field<TextItemField>("sf-key-report-id");
                            KeyReportID.Value = data.SharefileScreenshotId;

                            var KeyReportReportID = MyItem.Field<TextItemField>("sf-key-report-id-report");
                            KeyReportReportID.Value = data.SharefileReportId;

                            var itemId = await podio2.ItemService.UpdateItem(MyItem, null, null, false, true);

                            data.Name = data.ClientName;
                            data.ClientItemId = check.ClientItemId;
                            data.ItemId = check.PodioItemId;
                            data.PodioItemId = check.PodioItemId;
                            data.PodioLink = data.PodioLink;

                            clients.Id = check.Id;
                            _soxContext.Entry(check).CurrentValues.SetValues(data);
                            await _soxContext.SaveChangesAsync();
                            context.Commit();

                            return Ok("updated");
                        }// a2bhq test app client app end bracket podio authentication
                    }// a2bha client app end bracket podio authentication

                   
                }
                else
                { 
                    //insert

                    if (podio.IsAuthenticated())
                    {

                        Item MyItem1 = new Item();

                        var ClientName = MyItem1.Field<TextItemField>("cliend-reference-code");
                        ClientName.Value = clients.ClientName;

                        var ClientCode = MyItem1.Field<TextItemField>("unique-identifier-code");
                        ClientCode.Value = clients.ClientCode;

                        var ClientCodeReferrence = await podio.ItemService.AddNewItem(Int32.Parse(PodioAppKey.AppId), MyItem1);

                        PodioAppKey.AppId = _config.GetSection("SampleSelectionPodioApp").GetSection("ClientAppId").Value;
                        PodioAppKey.AppToken = _config.GetSection("SampleSelectionPodioApp").GetSection("ClientAppToken").Value;

                        var podio2 = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                        ServicePointManager.Expect100Continue = true;
                        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                        await podio2.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                        if (podio2.IsAuthenticated())
                        {

                            var data = clients;
                            Item MyItem = new Item();

                            var Title = MyItem.Field<TextItemField>("title");
                            Title.Value = data.ClientName;

                            var ReferrenceId = MyItem.Field<AppItemField>("client-relationship");
                            ReferrenceId.ItemIds = new List<int> { ClientCodeReferrence };

                            var External = MyItem.Field<TextItemField>("external-auditor");
                            External.Value = data.ExternalAuditor;

                            var Round1Percent = MyItem.Field<NumericItemField>("round-1-percent");
                            Round1Percent.Value = data.Percent;

                            var Round2Percent = MyItem.Field<NumericItemField>("round-2-percent");
                            Round2Percent.Value = data.PercentRound2;

                            var SharefileIdSave = MyItem.Field<TextItemField>("sharefile-id-save-file");
                            SharefileIdSave.Value = data.SharefileId;

                            var KeyReportID = MyItem.Field<TextItemField>("sf-key-report-id");
                            KeyReportID.Value = data.SharefileScreenshotId;

                            var KeyReportReportID = MyItem.Field<TextItemField>("sf-key-report-id-report");
                            KeyReportReportID.Value = data.SharefileReportId;

                            int itemId = await podio2.ItemService.AddNewItem(Int32.Parse(PodioAppKey.AppId), MyItem);

                            data.Name = data.ClientName;
                            data.ClientItemId = ClientCodeReferrence;
                            data.ItemId = itemId;
                            data.PodioItemId = itemId;
                            data.PodioLink = data.PodioLink;

                            _soxContext.Add(data);
                            await _soxContext.SaveChangesAsync();
                            context.Commit();
                            return Ok(clients);

                        }// a2bhq test app client app end bracket podio authentication
                    }// a2bha client app end bracket podio authentication
                }

            }

           


            return Ok("test");
        }


        [AllowAnonymous]
        [HttpGet("clients")]
        public IActionResult GetClientAsync()
        {
            List<string> _clients = new List<string>();
            try
            {
                using (var context = _soxContext.Database.BeginTransaction())
                {
                    var listClient = _soxContext.ClientSs.OrderBy(x => x.Name).Select(x => x.Name);

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

        
        [AllowAnonymous]
        [HttpGet("clientscode")]
        public IActionResult GetClientCodeAsync()
        {
           List<string> _clients = new List<string>();
            try
            {
                using (var context = _soxContext.Database.BeginTransaction())
                {
                    var listClient = _soxContext.ClientSs.OrderBy(x => x.ClientCode).Select(x => x.ClientCode);

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

        [AllowAnonymous]
        [HttpPost("search")]
        public IActionResult Search([FromBody]ClientSs filter)
        {
            try
            {
                using (var context = _soxContext.Database.BeginTransaction())
                {
                    var listClient = _soxContext.ClientSs.Where(
                        x => x.ClientCode.Equals(filter.ClientCode) && 
                        x.ClientName.Equals(filter.ClientName)
                        )
                        .AsNoTracking().FirstOrDefault();


                        if (listClient != null)
                        {
                            return Ok(listClient);
                        }
                        else
                        {
                            return NoContent();
                        }
                }
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetListClientAsync {ex}", "ErrorGetListClientAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetListClientAsync");
                return BadRequest();
            }

        }


    }


}


    