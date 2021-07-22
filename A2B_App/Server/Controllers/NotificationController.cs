using A2B_App.Server.Data;
using A2B_App.Server.Services;
using A2B_App.Shared.Email;
using A2B_App.Shared.Skype;
using A2B_App.Shared.Sox;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace A2B_App.Server.Controllers
{
    [Authorize]
    [Route("api/[controller]")]
    [ApiController]
    public class NotificationController : ControllerBase
    {
        private readonly IConfiguration _config;
        private readonly ILogger<NotificationController> _logger;
        private readonly IWebHostEnvironment _environment;
        private readonly NotificationContext _notificationContext;

        public NotificationController(IConfiguration config,
            ILogger<NotificationController> logger,
            IWebHostEnvironment environment,
            NotificationContext notificationContext)
        {
            _config = config;
            _logger = logger;
            _environment = environment;
            _notificationContext = notificationContext;
        }


        [AllowAnonymous]
        [HttpPost("createSkypeObj")]
        public async Task<IActionResult> CreateSkypeObjAsync([FromBody] SkypeObj skypeObj)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                

            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error CreateSkypeObjAsync");
                FileLog.Write(ex.ToString(), "ErrorCreateSkypeObjAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "CreateSkypeObjAsync");
                return BadRequest(ex.ToString());
            }

            return Ok();

        }

        [AllowAnonymous]
        [HttpPost("testEmailNotification")]
        public async Task<IActionResult> TestEmailNotificationAsync()
        {
            //Console.WriteLine("Triggered process");
            try
            {

                EmailService emailService = new EmailService(_config);
                EmailParam emailParam = new EmailParam();

                EmailNotificationBody notifBody = new EmailNotificationBody();
                notifBody.Message = $"Sox workpaper is waiting for your review";
                notifBody.Url = $"https://www.google.com";
                notifBody.UrlName = $"Sox workpaper";
                notifBody.UserName = $"Levin Tagapan";

                emailParam.Message = await emailService.EmailNotificationWorkpaperTemplate(notifBody);
                emailParam.Subject = $"Sox Workpaper for Review";
                emailParam.ListEmailTo = new List<string> { "ltagapan@a2bhq.com" };
                if(emailService.Send(emailParam))
                {
                    return Ok();
                }
                else
                    return NoContent();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error TestEmailNotificationAsync");
                FileLog.Write(ex.ToString(), "TestEmailNotificationAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "TestEmailNotificationAsync");
                return BadRequest(ex.ToString());
            }

           

        }

        [AllowAnonymous]
        [HttpPost("testSkypeNotification")]
        public async Task<IActionResult> TestSkypeNotificationAsync()
        {
            //Console.WriteLine("Triggered process");
            try
            {
                string bell = $"(bell)(bell)(bell)";
                SkypeBot skypeBot = new SkypeBot(_config);
                Skype skypeParam = new Skype();
                //skypeParam.Address = _config["Skype:ErrorGc"];
                skypeParam.Address = "19:9ce1a5bf0e244cb08d60d5c93ca7e6c5@thread.skype";
                skypeParam.Message = $"{bell}<br/><b>Sox Workpaper for Review </b><br/><br/>Client: ABC Company<br/>FY: FY21<br/>Process: ITGC<br/>Control: TAX 1.2<br/>https://soxroxapp.com<br/>{bell}";
                bool isSuccess = await skypeBot.SendSkypeNotif(skypeParam);

                if (isSuccess)
                {
                    return Ok();
                }
                else
                    return NoContent();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error TestSkypeNotificationAsync");
                FileLog.Write(ex.ToString(), "TestSkypeNotificationAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "TestSkypeNotificationAsync");
                return BadRequest(ex.ToString());
            }



        }

    }
}
