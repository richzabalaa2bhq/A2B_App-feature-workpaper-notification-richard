using A2B_App.Server.Data;
//using A2B_App.Server.Log;
using A2B_App.Server.Services;
using A2B_App.Shared.Sms;
using A2B_App.Shared.Sox;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Diagnostics;
using System.Threading.Tasks;

namespace A2B_App.Server.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class SmsController : ControllerBase
    {

        private readonly IConfiguration _config;
        private readonly ILogger<SmsController> _logger;
        private readonly IWebHostEnvironment _environment;
        private readonly SmsContext _smsContext;

        public SmsController(IConfiguration config, ILogger<SmsController> logger, IWebHostEnvironment environment, SmsContext smsContext)
        {
            _config = config;
            _smsContext = smsContext;
            _environment = environment;
            _logger = logger;
        }

        [HttpGet("globe/sms")]
        public async Task<IActionResult> GetGlobeSmsAsync([FromQuery] GlobeSms globeSms)
        {

            //[FromQuery]string access_token, [FromQuery]string subscriber_number

            try
            {
                SmsServices SmsService = new SmsServices(_smsContext, _config);
                FileLog.Write($"Receive GetGlobeSmsAsync: {Newtonsoft.Json.JsonConvert.SerializeObject(globeSms)}", "GetGlobeSmsAsync");
                if (await SmsService.SaveGlobeSmsAsync(globeSms))
                {

                }
                else
                {
                    FileLog.Write($"Error: {Newtonsoft.Json.JsonConvert.SerializeObject(globeSms)}", "ErrorGetGlobeSmsAsync");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write($"Error: {Newtonsoft.Json.JsonConvert.SerializeObject(globeSms)}", "ErrorGetGlobeSmsAsync");
                AdminService adminService = new AdminService(_config);
               // adminService.SendAlert(true, true, ex.ToString(), "GetGlobeSmsAsync");
                //throw;
            }

            return Ok();
        }

        [HttpPost("globe/sms")]
        public async Task<IActionResult> PostGlobeSmsAsync([FromBody] GlobeSms globeSms)
        {
            try
            {
                SmsServices SmsService = new SmsServices(_smsContext, _config);
                FileLog.Write($"Receive PostGlobeSmsAsync: {Newtonsoft.Json.JsonConvert.SerializeObject(globeSms)}", "PostGlobeSmsAsync");
                if (await SmsService.SaveGlobeSmsAsync(globeSms))
                {

                }
                else
                {
                    FileLog.Write($"Error: {Newtonsoft.Json.JsonConvert.SerializeObject(globeSms)}", "ErrorPostGlobeSmsAsync");
                }
            }
            catch (Exception ex)
            {
                FileLog.Write($"Error: {Newtonsoft.Json.JsonConvert.SerializeObject(globeSms)}", "ErrorPostGlobeSmsAsync");
                AdminService adminService = new AdminService(_config);
             //   adminService.SendAlert(true, true, ex.ToString(), "PostGlobeSmsAsync");
                //throw;
            }

            return Ok();
        }

        [HttpPost("globe/sms/send")]
        public IActionResult PostGlobeSendSmsAsync([FromBody] SmsSend smsSend)
        {
            try
            {
                SmsServices SmsService = new SmsServices(_smsContext, _config);
                if (SmsService.SendGlobeSmsAsync(smsSend))
                {
                    FileLog.Write($"Receive PostGlobeSendSmsAsync: {Newtonsoft.Json.JsonConvert.SerializeObject(smsSend)}", "PostGlobeSendSmsAsync");
                }
                else
                {
                    FileLog.Write($"Error: {Newtonsoft.Json.JsonConvert.SerializeObject(smsSend)}", "ErrorPostGlobeSendSmsAsync");
                }
            }
            catch (Exception ex)
            {
                FileLog.Write($"Error: {Newtonsoft.Json.JsonConvert.SerializeObject(smsSend)}", "ErrorPostGlobeSendSmsAsync");
                AdminService adminService = new AdminService(_config);
              //  adminService.SendAlert(true, true, ex.ToString(), "PostGlobeSendSmsAsync");
                //throw;
            }

            return Ok();
        }

        [HttpPost("globe/employeeref")]

      
        public async Task<IActionResult> PostGlobeEmployeeRefAsync([FromBody] EmployeeSmsReference employeeSms)
        {
            try
            {
                SmsServices SmsService = new SmsServices(_smsContext, _config);
                if (await SmsService.AddEmployeeRefAsync(employeeSms))
                {
                    FileLog.Write($"Receive PostGlobeEmployeeRefAsync: {Newtonsoft.Json.JsonConvert.SerializeObject(employeeSms)}", "PostGlobeEmployeeRefAsync");
                }
                else
                {
                    FileLog.Write($"Error: {Newtonsoft.Json.JsonConvert.SerializeObject(employeeSms)}", "ErrorPostGlobeEmployeeRefAsync");
                }
            }
            catch (Exception ex)
            {
                FileLog.Write($"Error: {Newtonsoft.Json.JsonConvert.SerializeObject(employeeSms)}", "ErrorPostGlobeEmployeeRefAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "PostGlobeEmployeeRefAsync");
                //throw;
            }

            return Ok();
        }
      

    }
}
