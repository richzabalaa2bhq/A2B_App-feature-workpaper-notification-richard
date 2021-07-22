using A2B_App.Server.Data;
//using A2B_App.Server.Log;
using A2B_App.Server.Services;
using A2B_App.Shared.Podio;
using A2B_App.Shared.Sox;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using PodioAPI;
using System;
using System.Diagnostics;
using System.Net;
using System.Threading.Tasks;

namespace A2B_App.Server.Controllers
{

    [Route("api/[controller]")]
    [ApiController]
    public class WebhookController : ControllerBase
    {

        private readonly IConfiguration _config;
        private readonly ILogger<WebhookController> _logger;
        private readonly IWebHostEnvironment _environment;
        private readonly SoxContext _soxContext;
        private readonly TimeContext _timeContext;

        public WebhookController(
            IConfiguration config,
            ILogger<WebhookController> logger,
            IWebHostEnvironment environment,
            SoxContext soxContext,
            TimeContext timeContext
        )
        {
            _config = config;
            _logger = logger;
            _environment = environment;
            _soxContext = soxContext;
            _timeContext = timeContext;
        }

        #region Sox Webhook
        [AllowAnonymous]
        [HttpPost("podio/sampleselection/388d54eb-85a6-4da3-b7bc-18061fe4f181/{Id}/{Token}")]
        public async Task<IActionResult> HookTestAsync(PodioHook podioHook, string Id, string Token)
        {
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

            PodioApiKey PodioKey = new PodioApiKey();
            PodioAppKey PodioAppKey = new PodioAppKey();
            PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
            PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
            PodioAppKey.AppId = Id;
            PodioAppKey.AppToken = Token;

            var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
            await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);
            string rawRequest = string.Empty;


            if (podio.IsAuthenticated())
            {
                //var request = context.Request;
                rawRequest = Newtonsoft.Json.JsonConvert.SerializeObject(podioHook);
                podioHook.AppId = PodioAppKey.AppId.ToString();

                switch (podioHook.type)
                {
                    case "hook.verify":
                        await podio.HookService.ValidateHookVerification(podioHook.hook_id, podioHook.code);
                        FileLog.Write($"hook.verify raw request - {rawRequest}", "HookTestAsync");
                        break;
                    // An item was created
                    case "item.create":
                        // For item events you will get "item_id", "item_revision_id" and "external_id". in post params
                        long itemIdOfCreatedItem = podioHook.item_id;
                        // Fetch the item and do what ever you want
                        FileLog.Write($"item.create request - {rawRequest}", "HookTestAsync");
                        break;

                    // An item was updated
                    case "item.update":
                        // For item events you will get "item_id", "item_revision_id" and "external_id". in post params
                        long itemIdOfUpdatedItem = podioHook.item_id;
                        // Fetch the item and do what ever you want
                        FileLog.Write($"item.update request - {rawRequest}", "HookTestAsync");
                        break;

                    // An item was deleted    
                    case "item.delete":
                        // For item events you will get "item_id", "item_revision_id" and "external_id". in post params
                        long deletedItemId = podioHook.item_id;
                        FileLog.Write($"item.delete request - {rawRequest}", "HookTestAsync");
                        break;
                }
            }

            //Podio require to return status 200 
            return Ok();

        }

        [AllowAnonymous]
        [HttpPost("podio/sampleselection2/eda2d403-95f6-4b49-b44a-f452b8e39e19")]
        public async Task<IActionResult> HookSampleSelectionAsync(PodioHook podioHook)
        {
            //Console.WriteLine("Triggered process");
            try
            {
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("SampleSelectionPodioApp").GetSection("AppId").Value;
                PodioAppKey.AppToken = _config.GetSection("SampleSelectionPodioApp").GetSection("AppToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);
                bool isHookSave, isItemSave;
                SampleSelectionService PodioService;

                if (podio.IsAuthenticated())
                {
                    //var request = context.Request;
                    var rawRequest = Newtonsoft.Json.JsonConvert.SerializeObject(podioHook);
                    podioHook.AppId = PodioAppKey.AppId.ToString();

                    switch (podioHook.type)
                    {
                        case "hook.verify":
                            await podio.HookService.ValidateHookVerification(podioHook.hook_id, podioHook.code);
                            FileLog.Write($"hook.verify raw request - {rawRequest}", "HookSampleSelectionAsync");
                            break;
                        // An item was created
                        case "item.create":
                            // For item events you will get "item_id", "item_revision_id" and "external_id". in post params
                            long itemIdOfCreatedItem = podioHook.item_id;
                            // Fetch the item and do what ever you want
                            FileLog.Write($"item.create request - {rawRequest}", "HookSampleSelectionAsync");
                            PodioService = new SampleSelectionService(_soxContext, _config);
                            isHookSave = PodioService.SaveWebhookAsync(podioHook);
                            isItemSave = await PodioService.SavePodioSampleSelectionAsync((int)podioHook.item_id);
                            if (!isHookSave || !isItemSave)
                                FileLog.Write($"Error saving webhook request - {rawRequest}", "HookSampleSelectionAsync");
                            break;

                        // An item was updated
                        case "item.update":
                            // For item events you will get "item_id", "item_revision_id" and "external_id". in post params
                            long itemIdOfUpdatedItem = podioHook.item_id;
                            // Fetch the item and do what ever you want
                            FileLog.Write($"item.update request - {rawRequest}", "HookSampleSelectionAsync");
                            PodioService = new SampleSelectionService(_soxContext, _config);
                            isHookSave = PodioService.SaveWebhookAsync(podioHook);
                            isItemSave = await PodioService.SavePodioSampleSelectionAsync((int)podioHook.item_id);
                            if (!isHookSave || !isItemSave)
                                FileLog.Write($"Error saving webhook request - {rawRequest}", "ErrorHookSampleSelectionAsync");
                            break;

                        // An item was deleted    
                        case "item.delete":
                            // For item events you will get "item_id", "item_revision_id" and "external_id". in post params
                            long deletedItemId = podioHook.item_id;
                            FileLog.Write($"item.delete request - {rawRequest}", "HookSampleSelectionAsync");
                            PodioService = new SampleSelectionService(_soxContext, _config);
                            isHookSave = PodioService.SaveWebhookAsync(podioHook);
                            isItemSave = await PodioService.DeleteSampleSelectionAsync((int)podioHook.item_id);
                            if (!isHookSave || !isItemSave)
                                FileLog.Write($"Error saving webhook request - {rawRequest}", "ErrorHookSampleSelectionAsync");
                            break;
                    }

                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorHookSampleSelectionAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "HookSampleSelectionAsync");

            }

            //Podio require to return status 200 
            return Ok();

        }

        [AllowAnonymous]
        [HttpPost("podio/sampleselectiontestround/10bba44c-fd75-4fba-b3c9-fb5131bc9120")]
        public async Task<IActionResult> HookTestRoundAsync(PodioHook podioHook)
        {
            //Console.WriteLine("Triggered process");
            try
            {
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("SampleSelectionPodioApp").GetSection("RoundAppId").Value;
                PodioAppKey.AppToken = _config.GetSection("SampleSelectionPodioApp").GetSection("RoundAppToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);
                bool isHookSave, isItemSave;
                SampleSelectionService PodioService;

                if (podio.IsAuthenticated())
                {
                    //var request = context.Request;
                    var rawRequest = Newtonsoft.Json.JsonConvert.SerializeObject(podioHook);
                    podioHook.AppId = PodioAppKey.AppId.ToString();

                    switch (podioHook.type)
                    {
                        case "hook.verify":
                            await podio.HookService.ValidateHookVerification(podioHook.hook_id, podioHook.code);
                            FileLog.Write($"hook.verify raw request - {rawRequest}", "HookTestRoundAsync");
                            break;
                        // An item was created
                        case "item.create":
                            // For item events you will get "item_id", "item_revision_id" and "external_id". in post params
                            long itemIdOfCreatedItem = podioHook.item_id;
                            // Fetch the item and do what ever you want
                            FileLog.Write($"item.create request - {rawRequest}", "HookTestRoundAsync");
                            PodioService = new SampleSelectionService(_soxContext, _config);
                            isHookSave = PodioService.SaveWebhookAsync(podioHook);
                            isItemSave = await PodioService.SavePodioTestRoundAsync((int)podioHook.item_id);
                            if (!isHookSave || !isItemSave)
                                FileLog.Write($"Error saving webhook request - {rawRequest}", "ErrorHookTestRoundAsync");
                            break;

                        // An item was updated
                        case "item.update":
                            // For item events you will get "item_id", "item_revision_id" and "external_id". in post params
                            long itemIdOfUpdatedItem = podioHook.item_id;
                            // Fetch the item and do what ever you want
                            FileLog.Write($"item.update request - {rawRequest}", "HookTestRoundAsync");
                            PodioService = new SampleSelectionService(_soxContext, _config);
                            isHookSave = PodioService.SaveWebhookAsync(podioHook);
                            isItemSave = await PodioService.SavePodioTestRoundAsync((int)podioHook.item_id);
                            if (!isHookSave || !isItemSave)
                                FileLog.Write($"Error saving webhook request - {rawRequest}", "ErrorHookTestRoundAsync");
                            break;



                        // An item was deleted    
                        case "item.delete":
                            // For item events you will get "item_id", "item_revision_id" and "external_id". in post params
                            long deletedItemId = podioHook.item_id;
                            FileLog.Write($"item.delete request - {rawRequest}", "HookTestRoundAsync");
                            PodioService = new SampleSelectionService(_soxContext, _config);
                            isHookSave = PodioService.SaveWebhookAsync(podioHook);
                            isItemSave = await PodioService.DeleteSampleSelectionAsync((int)podioHook.item_id);
                            if (!isHookSave || !isItemSave)
                                FileLog.Write($"Error saving webhook request - {rawRequest}", "ErrorHookTestRoundAsync");
                            break;
                    }

                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorHookTestRoundAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "HookTestRoundAsync");

            }

            //Podio require to return status 200 
            return Ok();

        }

        [AllowAnonymous]
        [HttpPost("podio/sampleselectionclient/d8f64d1e-5c83-45b6-9b42-427585e20996")]
        public async Task<IActionResult> HookSampleSelectionClientAsync(PodioHook podioHook)
        {
            //Console.WriteLine("Triggered process");
            try
            {
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("SampleSelectionPodioApp").GetSection("ClientAppId").Value;
                PodioAppKey.AppToken = _config.GetSection("SampleSelectionPodioApp").GetSection("ClientAppToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);
                bool isHookSave, isItemSave;
                SampleSelectionService SampleSelectionService;

                if (podio.IsAuthenticated())
                {
                    //var request = context.Request;
                    var rawRequest = Newtonsoft.Json.JsonConvert.SerializeObject(podioHook);
                    podioHook.AppId = PodioAppKey.AppId.ToString();

                    switch (podioHook.type)
                    {
                        case "hook.verify":
                            await podio.HookService.ValidateHookVerification(podioHook.hook_id, podioHook.code);
                            FileLog.Write($"hook.verify raw request - {rawRequest}", "HookSampleSelectionClientAsync");
                            break;
                        // An item was created
                        case "item.create":
                            // For item events you will get "item_id", "item_revision_id" and "external_id". in post params
                            long itemIdOfCreatedItem = podioHook.item_id;
                            // Fetch the item and do what ever you want
                            FileLog.Write($"item.create request - {rawRequest}", "HookSampleSelectionClientAsync");
                            SampleSelectionService = new SampleSelectionService(_soxContext, _config);
                            isHookSave = SampleSelectionService.SaveWebhookAsync(podioHook);
                            isItemSave = await SampleSelectionService.SavePodioClientSSAsync((int)podioHook.item_id);
                            if (!isItemSave)
                                FileLog.Write($"item.create request - {rawRequest}", "HookSampleSelectionClientAsync");
                            break;

                        // An item was updated
                        case "item.update":
                            // For item events you will get "item_id", "item_revision_id" and "external_id". in post params
                            long itemIdOfUpdatedItem = podioHook.item_id;
                            // Fetch the item and do what ever you want
                            FileLog.Write($"item.update request - {rawRequest}", "HookSampleSelectionClientAsync");
                            SampleSelectionService = new SampleSelectionService(_soxContext, _config);
                            isHookSave = SampleSelectionService.SaveWebhookAsync(podioHook);
                            isItemSave = await SampleSelectionService.SavePodioClientSSAsync((int)podioHook.item_id);
                            if (!isItemSave)
                                FileLog.Write($"item.update request - {rawRequest}", "HookSampleSelectionClientAsync");
                            break;

                        // An item was deleted    
                        case "item.delete":
                            // For item events you will get "item_id", "item_revision_id" and "external_id". in post params
                            long deletedItemId = podioHook.item_id;
                            FileLog.Write($"item.delete request - {rawRequest}", "HookSampleSelectionClientAsync");
                            SampleSelectionService = new SampleSelectionService(_soxContext, _config);
                            isHookSave = SampleSelectionService.SaveWebhookAsync(podioHook);
                            isItemSave = await SampleSelectionService.DeleteClientSSAsync((int)podioHook.item_id);
                            if (!isItemSave)
                                FileLog.Write($"item.delete request - {rawRequest}", "HookSampleSelectionClientAsync");
                            break;
                    }

                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorHookSampleSelectionClientAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "HookSampleSelectionClientAsync");

            }

            //Podio require to return status 200 
            return Ok();

        }

        [AllowAnonymous]
        [HttpPost("podio/sampleselectionmatrix/89308192-1327-4be3-a996-9372df8f5266")]
        public async Task<IActionResult> HookSampleSelectionMatrixAsync(PodioHook podioHook)
        {
            //Console.WriteLine("Triggered process");
            try
            {
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("SampleSelectionPodioApp").GetSection("MatrixAppId").Value;
                PodioAppKey.AppToken = _config.GetSection("SampleSelectionPodioApp").GetSection("MatrixAppToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);
                bool isHookSave, isItemSave;
                SampleSelectionService SampleSelectionService;

                if (podio.IsAuthenticated())
                {
                    //var request = context.Request;
                    var rawRequest = Newtonsoft.Json.JsonConvert.SerializeObject(podioHook);
                    podioHook.AppId = PodioAppKey.AppId.ToString();

                    switch (podioHook.type)
                    {
                        case "hook.verify":
                            await podio.HookService.ValidateHookVerification(podioHook.hook_id, podioHook.code);
                            FileLog.Write($"hook.verify raw request - {rawRequest}", "HookSampleSelectionMatrixAsync");
                            break;
                        // An item was created
                        case "item.create":
                            // For item events you will get "item_id", "item_revision_id" and "external_id". in post params
                            long itemIdOfCreatedItem = podioHook.item_id;
                            // Fetch the item and do what ever you want
                            FileLog.Write($"item.create request - {rawRequest}", "HookSampleSelectionMatrixAsync");
                            SampleSelectionService = new SampleSelectionService(_soxContext, _config);
                            isHookSave = SampleSelectionService.SaveWebhookAsync(podioHook);
                            isItemSave = await SampleSelectionService.SavePodioMatrixAsync((int)podioHook.item_id);
                            if (!isItemSave)
                                FileLog.Write($"item.create request - {rawRequest}", "HookSampleSelectionMatrixAsync");
                            break;

                        // An item was updated
                        case "item.update":
                            // For item events you will get "item_id", "item_revision_id" and "external_id". in post params
                            long itemIdOfUpdatedItem = podioHook.item_id;
                            FileLog.Write($"item.update request - {rawRequest}", "HookSampleSelectionMatrixAsync");
                            // Fetch the item and do what ever you want
                            SampleSelectionService = new SampleSelectionService(_soxContext, _config);
                            isHookSave = SampleSelectionService.SaveWebhookAsync(podioHook);
                            isItemSave = await SampleSelectionService.SavePodioMatrixAsync((int)podioHook.item_id);
                            if (!isItemSave)
                                FileLog.Write($"item.update request - {rawRequest}", "HookSampleSelectionMatrixAsync");
                            break;
                        // An item was deleted    
                        case "item.delete":
                            // For item events you will get "item_id", "item_revision_id" and "external_id". in post params
                            long deletedItemId = podioHook.item_id;
                            FileLog.Write($"item.delete request - {rawRequest}", "HookSampleSelectionMatrixAsync");
                            SampleSelectionService = new SampleSelectionService(_soxContext, _config);
                            isHookSave = SampleSelectionService.SaveWebhookAsync(podioHook);
                            isItemSave = await SampleSelectionService.DeleteMatrixAsync((int)podioHook.item_id);
                            if (!isItemSave)
                                FileLog.Write($"item.delete request - {rawRequest}", "HookSampleSelectionMatrixAsync");
                            break;
                    }

                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorHookSampleSelectionMatrixAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "HookSampleSelectionMatrixAsync");
            }

            //Podio require to return status 200 
            return Ok();

        }

        [AllowAnonymous]
        [HttpPost("podio/rcmcta/12d25d9f-4aef-4879-8e28-7dd573b3805c")]
        public async Task<IActionResult> HookRcmCtaAsync(PodioHook podioHook)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("RcmPodioApp").GetSection("CtaId").Value;
                PodioAppKey.AppToken = _config.GetSection("RcmPodioApp").GetSection("CtaToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);
                bool isHookSave, isItemSave;
                RcmService RcmService;

                if (podio.IsAuthenticated())
                {
                    //var request = context.Request;
                    var rawRequest = Newtonsoft.Json.JsonConvert.SerializeObject(podioHook);
                    podioHook.AppId = PodioAppKey.AppId.ToString();

                    switch (podioHook.type)
                    {
                        case "hook.verify":
                            await podio.HookService.ValidateHookVerification(podioHook.hook_id, podioHook.code);
                            FileLog.Write($"hook.verify raw request - {rawRequest}", "HookRcmCtaAsync");
                            break;
                        // An item was created
                        case "item.create":
                            // For item events you will get "item_id", "item_revision_id" and "external_id". in post params
                            long itemIdOfCreatedItem = podioHook.item_id;
                            FileLog.Write($"hook.verify raw request - {rawRequest}", "HookRcmCtaAsync");
                            RcmService = new RcmService(_soxContext, _config);
                            isHookSave = RcmService.SaveWebhookAsync(podioHook);
                            isItemSave = await RcmService.SavePodioRcmCtaAsync((int)podioHook.item_id);
                            if (!isItemSave)
                                FileLog.Write($"Error saving webhook request - {rawRequest}", "ErrorHookRcmCtaAsync");
                            break;

                        // An item was updated
                        case "item.update":
                            // For item events you will get "item_id", "item_revision_id" and "external_id". in post params
                            long itemIdOfUpdatedItem = podioHook.item_id;
                            FileLog.Write($"hook.verify raw request - {rawRequest}", "HookRcmCtaAsync");
                            RcmService = new RcmService(_soxContext, _config);
                            isHookSave = RcmService.SaveWebhookAsync(podioHook);
                            isItemSave = await RcmService.SavePodioRcmCtaAsync((int)podioHook.item_id);
                            if (!isItemSave)
                                FileLog.Write($"Error saving webhook request - {rawRequest}", "ErrorHookRcmCtaAsync");
                            break;
                        // An item was deleted    
                        case "item.delete":
                            // For item events you will get "item_id", "item_revision_id" and "external_id". in post params
                            long deletedItemId = podioHook.item_id;
                            FileLog.Write($"hook.verify raw request - {rawRequest}", "HookRcmCtaAsync");
                            RcmService = new RcmService(_soxContext, _config);
                            isHookSave = RcmService.SaveWebhookAsync(podioHook);
                            isItemSave = await RcmService.DeleteRcmCtaAsync((int)podioHook.item_id);
                            if (!isItemSave)
                                FileLog.Write($"Error saving webhook request - {rawRequest}", "ErrorHookRcmCtaAsync");
                            break;
                    }

                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorHookRcmCtaAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "HookRcmCtaAsync");
            }

            //Podio require to return status 200 
            return Ok();

        }

        [AllowAnonymous]
        [HttpPost("podio/rcm/6db681cd-4b97-48ff-a889-3ce366707758")]
        public async Task<IActionResult> HookRcmAsync(PodioHook podioHook)
        {
            //Console.WriteLine("Triggered process");
            try
            {
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("RcmPodioApp").GetSection("RcmAppId").Value;
                PodioAppKey.AppToken = _config.GetSection("RcmPodioApp").GetSection("RcmAppToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);
                bool isHookSave, isItemSave;
                RcmService RcmService;

                if (podio.IsAuthenticated())
                {
                    //var request = context.Request;
                    var rawRequest = Newtonsoft.Json.JsonConvert.SerializeObject(podioHook);
                    podioHook.AppId = PodioAppKey.AppId.ToString();

                    switch (podioHook.type)
                    {
                        case "hook.verify":
                            await podio.HookService.ValidateHookVerification(podioHook.hook_id, podioHook.code);
                            FileLog.Write($"hook.verify raw request - {rawRequest}", "HookRcmAsync");
                            break;
                        // An item was created
                        case "item.create":
                            // For item events you will get "item_id", "item_revision_id" and "external_id". in post params
                            long itemIdOfCreatedItem = podioHook.item_id;
                            FileLog.Write($"hook.verify raw request - {rawRequest}", "HookRcmAsync");
                            RcmService = new RcmService(_soxContext, _config);
                            isHookSave = RcmService.SaveWebhookAsync(podioHook);
                            isItemSave = await RcmService.SavePodioRcmAsync((int)podioHook.item_id);
                            if (!isItemSave)
                                FileLog.Write($"Error saving webhook request - {rawRequest}", "ErrorHookRcmAsync");
                            break;

                        // An item was updated
                        case "item.update":
                            // For item events you will get "item_id", "item_revision_id" and "external_id". in post params
                            long itemIdOfUpdatedItem = podioHook.item_id;
                            FileLog.Write($"hook.verify raw request - {rawRequest}", "HookRcmAsync");
                            RcmService = new RcmService(_soxContext, _config);
                            isHookSave = RcmService.SaveWebhookAsync(podioHook);
                            isItemSave = await RcmService.SavePodioRcmAsync((int)podioHook.item_id);
                            if (!isItemSave)
                                FileLog.Write($"Error saving webhook request - {rawRequest}", "ErrorHookRcmAsync");
                            break;
                        // An item was deleted    
                        case "item.delete":
                            // For item events you will get "item_id", "item_revision_id" and "external_id". in post params
                            long deletedItemId = podioHook.item_id;
                            FileLog.Write($"hook.verify raw request - {rawRequest}", "HookRcmAsync");
                            RcmService = new RcmService(_soxContext, _config);
                            isHookSave = RcmService.SaveWebhookAsync(podioHook);
                            isItemSave = await RcmService.DeleteRcmAsync((int)podioHook.item_id);
                            if (!isItemSave)
                                FileLog.Write($"Error saving webhook request - {rawRequest}", "ErrorHookRcmAsync");
                            break;
                    }

                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorHookRcmAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "HookRcmAsync");
            }

            //Podio require to return status 200 
            return Ok();

        }

        [AllowAnonymous]
        [HttpPost("podio/rcm")]
        public async Task<IActionResult> HookRcm2Async(PodioHook podioHook)
        {
            //Console.WriteLine("Triggered process");
            try
            {
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("RcmPodioApp").GetSection("RcmAppId").Value;
                PodioAppKey.AppToken = _config.GetSection("RcmPodioApp").GetSection("RcmAppToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);
                bool isHookSave, isItemSave;
                RcmService RcmService;

                if (podio.IsAuthenticated())
                {
                    //var request = context.Request;
                    var rawRequest = Newtonsoft.Json.JsonConvert.SerializeObject(podioHook);
                    podioHook.AppId = PodioAppKey.AppId.ToString();

                    switch (podioHook.type)
                    {
                        case "hook.verify":
                            await podio.HookService.ValidateHookVerification(podioHook.hook_id, podioHook.code);
                            FileLog.Write($"hook.verify raw request - {rawRequest}", "HookRcmAsync");
                            break;
                        // An item was created
                        case "item.create":
                            // For item events you will get "item_id", "item_revision_id" and "external_id". in post params
                            long itemIdOfCreatedItem = podioHook.item_id;
                            FileLog.Write($"hook.verify raw request - {rawRequest}", "HookRcmAsync");
                            RcmService = new RcmService(_soxContext, _config);
                            isHookSave = RcmService.SaveWebhookAsync(podioHook);
                            isItemSave = await RcmService.SavePodioRcmAsync((int)podioHook.item_id);
                            if (!isItemSave)
                                FileLog.Write($"Error saving webhook request - {rawRequest}", "ErrorHookRcmAsync");
                            break;

                        // An item was updated
                        case "item.update":
                            // For item events you will get "item_id", "item_revision_id" and "external_id". in post params
                            long itemIdOfUpdatedItem = podioHook.item_id;
                            FileLog.Write($"hook.verify raw request - {rawRequest}", "HookRcmAsync");
                            RcmService = new RcmService(_soxContext, _config);
                            isHookSave = RcmService.SaveWebhookAsync(podioHook);
                            isItemSave = await RcmService.SavePodioRcmAsync((int)podioHook.item_id);
                            if (!isItemSave)
                                FileLog.Write($"Error saving webhook request - {rawRequest}", "ErrorHookRcmAsync");
                            break;
                        // An item was deleted    
                        case "item.delete":
                            // For item events you will get "item_id", "item_revision_id" and "external_id". in post params
                            long deletedItemId = podioHook.item_id;
                            FileLog.Write($"hook.verify raw request - {rawRequest}", "HookRcmAsync");
                            RcmService = new RcmService(_soxContext, _config);
                            isHookSave = RcmService.SaveWebhookAsync(podioHook);
                            isItemSave = await RcmService.DeleteRcmAsync((int)podioHook.item_id);
                            if (!isItemSave)
                                FileLog.Write($"Error saving webhook request - {rawRequest}", "ErrorHookRcmAsync");
                            break;
                    }

                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorHookRcmAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "HookRcmAsync");
            }

            //Podio require to return status 200 
            return Ok(podioHook);

        }


        #endregion

        #region Master Time Webhook
        [AllowAnonymous]
        //Webhook for A2Q2 > Time > Time Code
        [HttpPost("podio/time/timecode/b3ef0cfe-b0d1-4415-b896-fc4afd713365")]
        public async Task<IActionResult> HookTimeCodeAsync(PodioHook podioHook)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("TimePodioApp").GetSection("TimeCodeAppId").Value;
                PodioAppKey.AppToken = _config.GetSection("TimePodioApp").GetSection("TimeCodeAppToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);
                bool isHookSave, isItemSave;
                TimeService TimeService;

                if (podio.IsAuthenticated())
                {
                    //var request = context.Request;
                    var rawRequest = Newtonsoft.Json.JsonConvert.SerializeObject(podioHook);
                    podioHook.AppId = PodioAppKey.AppId.ToString();

                    switch (podioHook.type)
                    {
                        case "hook.verify":
                            await podio.HookService.ValidateHookVerification(podioHook.hook_id, podioHook.code);
                            FileLog.Write($"hook.verify raw request - {rawRequest}", "HookTimeCodeAsync");
                            break;
                        // An item was created
                        case "item.create":
                            // For item events you will get "item_id", "item_revision_id" and "external_id". in post params
                            long itemIdOfCreatedItem = podioHook.item_id;
                            // Fetch the item and do what ever you want
                            FileLog.Write($"item.create request - {rawRequest}", "HookTimeCodeAsync");
                            TimeService = new TimeService(_timeContext, _config);
                            isHookSave = TimeService.SaveWebhookAsync(podioHook);
                            isItemSave = await TimeService.SavePodioTimeCodeAsync((int)podioHook.item_id);
                            if (!isItemSave)
                                FileLog.Write($"Error saving webhook request - {rawRequest}", "ErrorHookTimeCodeAsync");
                            break;

                        // An item was updated
                        case "item.update":
                            // For item events you will get "item_id", "item_revision_id" and "external_id". in post params
                            long itemIdOfUpdatedItem = podioHook.item_id;
                            // Fetch the item and do what ever you want
                            FileLog.Write($"item.update request - {rawRequest}", "HookTimeCodeAsync");
                            TimeService = new TimeService(_timeContext, _config);
                            isHookSave = TimeService.SaveWebhookAsync(podioHook);
                            isItemSave = await TimeService.SavePodioTimeCodeAsync((int)podioHook.item_id);
                            if (!isItemSave)
                                FileLog.Write($"Error saving webhook request - {rawRequest}", "ErrorHookTimeCodeAsync");
                            break;

                        // An item was deleted    
                        case "item.delete":
                            // For item events you will get "item_id", "item_revision_id" and "external_id". in post params
                            long deletedItemId = podioHook.item_id;
                            FileLog.Write($"item.delete request - {rawRequest}", "HookTimeCodeAsync");
                            TimeService = new TimeService(_timeContext, _config);
                            isHookSave = TimeService.SaveWebhookAsync(podioHook);
                            isItemSave = await TimeService.DeleteTimeCodeAsync((int)podioHook.item_id);
                            if (!isItemSave)
                                FileLog.Write($"Error saving webhook request - {rawRequest}", "ErrorHookTimeCodeAsync");
                            break;
                    }

                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorHookTimeCodeAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "HookTimeCodeAsync");
            }

            //Podio require to return status 200 
            return Ok();

        }

        [AllowAnonymous]
        //Webhook for A2Q2 > Time > Clients
        [HttpPost("podio/time/clientref/44335cf9-bd67-4b4a-a702-022a3602ce3a")]
        public async Task<IActionResult> HookClientReferenceAsync(PodioHook podioHook)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("TimePodioApp").GetSection("ClientAppId").Value;
                PodioAppKey.AppToken = _config.GetSection("TimePodioApp").GetSection("ClientAppToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);
                bool isHookSave, isItemSave;
                TimeService TimeService;

                if (podio.IsAuthenticated())
                {
                    //var request = context.Request;
                    var rawRequest = Newtonsoft.Json.JsonConvert.SerializeObject(podioHook);
                    podioHook.AppId = PodioAppKey.AppId.ToString();

                    switch (podioHook.type)
                    {
                        case "hook.verify":
                            await podio.HookService.ValidateHookVerification(podioHook.hook_id, podioHook.code);
                            FileLog.Write($"hook.verify raw request - {rawRequest}", "HookClientReferenceAsync");
                            break;
                        // An item was created
                        case "item.create":
                            // For item events you will get "item_id", "item_revision_id" and "external_id". in post params
                            long itemIdOfCreatedItem = podioHook.item_id;
                            // Fetch the item and do what ever you want
                            FileLog.Write($"item.create request - {rawRequest}", "HookClientReferenceAsync");
                            TimeService = new TimeService(_timeContext, _config);
                            isHookSave = TimeService.SaveWebhookAsync(podioHook);
                            isItemSave = await TimeService.SavePodioClientRefAsync((int)podioHook.item_id);
                            if (!isItemSave)
                                FileLog.Write($"Error saving webhook request - {rawRequest}", "ErrorHookClientReferenceAsync");
                            break;

                        // An item was updated
                        case "item.update":
                            // For item events you will get "item_id", "item_revision_id" and "external_id". in post params
                            long itemIdOfUpdatedItem = podioHook.item_id;
                            // Fetch the item and do what ever you want
                            FileLog.Write($"item.create request - {rawRequest}", "HookClientReferenceAsync");
                            TimeService = new TimeService(_timeContext, _config);
                            isHookSave = TimeService.SaveWebhookAsync(podioHook);
                            isItemSave = await TimeService.SavePodioClientRefAsync((int)podioHook.item_id);
                            if (!isItemSave)
                                FileLog.Write($"Error saving webhook request - {rawRequest}", "ErrorHookClientReferenceAsync");
                            break;

                        // An item was deleted    
                        case "item.delete":
                            // For item events you will get "item_id", "item_revision_id" and "external_id". in post params
                            long deletedItemId = podioHook.item_id;
                            FileLog.Write($"item.delete request - {rawRequest}", "HookClientReferenceAsync");
                            TimeService = new TimeService(_timeContext, _config);
                            isHookSave = TimeService.SaveWebhookAsync(podioHook);
                            isItemSave = await TimeService.DeleteClientRefAsync((int)podioHook.item_id);
                            if (!isItemSave)
                                FileLog.Write($"Error saving webhook request - {rawRequest}", "ErrorHookClientReferenceAsync");
                            break;
                    }

                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorHookClientReferenceAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "HookClientReferenceAsync");

            }

            //Podio require to return status 200 
            return Ok();

        }

        [AllowAnonymous]
        //Webhook for A2Q2 > Time > Project Reference
        [HttpPost("podio/time/projectref/aa2c1f95-bab3-4f3e-bcfd-dfd1e4a6c7bd")]
        public async Task<IActionResult> HookProjectReferenceAsync(PodioHook podioHook)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("TimePodioApp").GetSection("ProjectAppId").Value;
                PodioAppKey.AppToken = _config.GetSection("TimePodioApp").GetSection("ProjectAppToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);
                bool isHookSave, isItemSave;
                TimeService TimeService;

                if (podio.IsAuthenticated())
                {
                    //var request = context.Request;
                    var rawRequest = Newtonsoft.Json.JsonConvert.SerializeObject(podioHook);
                    podioHook.AppId = PodioAppKey.AppId.ToString();

                    switch (podioHook.type)
                    {
                        case "hook.verify":
                            await podio.HookService.ValidateHookVerification(podioHook.hook_id, podioHook.code);
                            FileLog.Write($"hook.verify raw request - {rawRequest}", "HookProjectReferenceAsync");
                            break;
                        // An item was created
                        case "item.create":
                            // For item events you will get "item_id", "item_revision_id" and "external_id". in post params
                            long itemIdOfCreatedItem = podioHook.item_id;
                            // Fetch the item and do what ever you want
                            FileLog.Write($"item.create request - {rawRequest}", "HookProjectReferenceAsync");
                            TimeService = new TimeService(_timeContext, _config);
                            isHookSave = TimeService.SaveWebhookAsync(podioHook);
                            isItemSave = await TimeService.SavePodioProjectRefAsync((int)podioHook.item_id);
                            if (!isItemSave)
                                FileLog.Write($"Error saving webhook request - {rawRequest}", "ErrorHookProjectReferenceAsync");
                            break;

                        // An item was updated
                        case "item.update":
                            // For item events you will get "item_id", "item_revision_id" and "external_id". in post params
                            long itemIdOfUpdatedItem = podioHook.item_id;
                            // Fetch the item and do what ever you want
                            FileLog.Write($"item.update request - {rawRequest}", "HookProjectReferenceAsync");
                            TimeService = new TimeService(_timeContext, _config);
                            isHookSave = TimeService.SaveWebhookAsync(podioHook);
                            isItemSave = await TimeService.SavePodioProjectRefAsync((int)podioHook.item_id);
                            if (!isItemSave)
                                FileLog.Write($"Error saving webhook request - {rawRequest}", "ErrorHookProjectReferenceAsync");
                            break;

                        // An item was deleted    
                        case "item.delete":
                            // For item events you will get "item_id", "item_revision_id" and "external_id". in post params
                            long deletedItemId = podioHook.item_id;
                            FileLog.Write($"item.delete request - {rawRequest}", "HookProjectReferenceAsync");
                            TimeService = new TimeService(_timeContext, _config);
                            isHookSave = TimeService.SaveWebhookAsync(podioHook);
                            isItemSave = await TimeService.DeleteProjectRefAsync((int)podioHook.item_id);
                            if (!isItemSave)
                                FileLog.Write($"Error saving webhook request - {rawRequest}", "ErrorHookProjectReferenceAsync");
                            break;
                    }

                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorHookProjectReferenceAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "HookProjectReferenceAsync");
            }

            //Podio require to return status 200 
            return Ok();

        }

        [AllowAnonymous]
        //Webhook for A2Q2 > Time > Task Reference
        [HttpPost("podio/time/taskref/337a3e72-b580-46b4-abe0-71b4cef4c9f0")]
        public async Task<IActionResult> HookTaskReferenceAsync(PodioHook podioHook)
        {
            //Console.WriteLine("Triggered process");
            try
            {
                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("TimePodioApp").GetSection("TaskAppId").Value;
                PodioAppKey.AppToken = _config.GetSection("TimePodioApp").GetSection("TaskAppToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);
                bool isHookSave, isItemSave;
                TimeService TimeService;

                if (podio.IsAuthenticated())
                {
                    //var request = context.Request;
                    var rawRequest = Newtonsoft.Json.JsonConvert.SerializeObject(podioHook);
                    podioHook.AppId = PodioAppKey.AppId.ToString();

                    switch (podioHook.type)
                    {
                        case "hook.verify":
                            await podio.HookService.ValidateHookVerification(podioHook.hook_id, podioHook.code);
                            FileLog.Write($"hook.verify raw request - {rawRequest}", "HookTaskReferenceAsync");
                            break;
                        // An item was created
                        case "item.create":
                            // For item events you will get "item_id", "item_revision_id" and "external_id". in post params
                            long itemIdOfCreatedItem = podioHook.item_id;
                            // Fetch the item and do what ever you want
                            FileLog.Write($"item.create request - {rawRequest}", "HookTaskReferenceAsync");
                            TimeService = new TimeService(_timeContext, _config);
                            isHookSave = TimeService.SaveWebhookAsync(podioHook);
                            isItemSave = await TimeService.SavePodioTaskRefAsync((int)podioHook.item_id);
                            if (!isItemSave)
                                FileLog.Write($"Error saving webhook request - {rawRequest}", "ErrorHookTaskReferenceAsync");
                            break;

                        // An item was updated
                        case "item.update":
                            // For item events you will get "item_id", "item_revision_id" and "external_id". in post params
                            long itemIdOfUpdatedItem = podioHook.item_id;
                            // Fetch the item and do what ever you want
                            FileLog.Write($"item.update request - {rawRequest}", "HookTaskReferenceAsync");
                            TimeService = new TimeService(_timeContext, _config);
                            isHookSave = TimeService.SaveWebhookAsync(podioHook);
                            isItemSave = await TimeService.SavePodioTaskRefAsync((int)podioHook.item_id);
                            if (!isItemSave)
                                FileLog.Write($"Error saving webhook request - {rawRequest}", "ErrorHookTaskReferenceAsync");
                            break;

                        // An item was deleted    
                        case "item.delete":
                            // For item events you will get "item_id", "item_revision_id" and "external_id". in post params
                            long deletedItemId = podioHook.item_id;
                            FileLog.Write($"item.delete request - {rawRequest}", "HookTaskReferenceAsync");
                            TimeService = new TimeService(_timeContext, _config);
                            isHookSave = TimeService.SaveWebhookAsync(podioHook);
                            isItemSave = await TimeService.DeleteTaskRefAsync((int)podioHook.item_id);
                            if (!isItemSave)
                                FileLog.Write($"Error saving webhook request - {rawRequest}", "ErrorHookTaskReferenceAsync");
                            break;
                    }

                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorHookTaskReferenceAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "HookTaskReferenceAsync");
            }

            //Podio require to return status 200 
            return Ok();

        }




        #endregion

        //---------------------------------------------------------------------------------------------
        //Will update all webhook api static route to custom api route that fetch from appsettings.json
        //---------------------------------------------------------------------------------------------
        [AllowAnonymous]
        [HttpPost("{controlName}")]
        public IActionResult TestDynamicRoute(string controlName)
        {
            return Ok(controlName);
        }


        [AllowAnonymous]
        [HttpPost("test")]
        public IActionResult Test([FromBody] object requestBody)
        {
            try
            {
                FileLog.Write($"Received request - {requestBody}", "Test");
            }
            catch (Exception ex)
            {
                FileLog.Write($"Error request - {ex}", "Test");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "Test");
                //throw;
            }
            
            return Ok(requestBody);
        }


    }
}
