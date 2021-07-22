using A2B_App.Server.Data;
using A2B_App.Server.Services;
using A2B_App.Shared.Skype;
using A2B_App.Shared.Sox;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Threading.Tasks;

namespace A2B_App.Server.Controllers
{
    [Authorize]
    [Route("api/[controller]")]
    [ApiController]
    public class MicrosoftBotController : ControllerBase
    {
        private readonly IConfiguration _config;
        private readonly ILogger<MicrosoftBotController> _logger;
        private readonly IWebHostEnvironment _environment;
        private readonly SoxContext _soxContext;
        //private readonly IBotFrameworkHttpAdapter _adapter;
        //private readonly IBot _bot;

        public MicrosoftBotController(IConfiguration config,
            ILogger<MicrosoftBotController> logger,
            IWebHostEnvironment environment,
            SoxContext soxContext)
        {
            _config = config;
            _logger = logger;
            _environment = environment;
            _soxContext = soxContext;
        }

        //[AllowAnonymous]
        [HttpPost("send")]
        public async Task<IActionResult> SendSkypeMessageAsync([FromBody] Skype skype)
        {

            try
            {
                SkypeBot skybot = new SkypeBot(_config);
                if (await skybot.SendSkypeNotif(skype))
                    return Ok();
                else
                    return BadRequest();

            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error SendSkypeMessageAsync {ex}", "ErrorSendSkypeMessageAsync");
                return BadRequest();
            }

            

        }

        //[AllowAnonymous]
        [HttpPost("sendAlert")]
        public async Task<IActionResult> SendAlertAsync()
        {

            try
            {
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, "Test", "SendAlertAsync");
                return Ok();


            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error SendAlertAsync {ex}", "ErrorSendAlertAsync");
                return BadRequest();
            }



        }

        [HttpPost("sendBizDevTest")]
        public async Task<IActionResult> SendBizDevTestAsync()
        {

            try
            {
                FormatService format = new FormatService();
                string Attendee = "Levin jay;John Doe";
                string Regarding = "BizDev Test";
                string AddressLocation = "<p>Join Zoom Meeting<br/></p><p>https://zoom.us/j/96983619235?pwd=WjlLMVY2SXJLUi8xUmFJQm1USGNmUT09</p><p>Meeting ID: 969 8361 9235<br/>Passcode: 105175<br/>One tap mobile<br/>+16699006833,,96983619235# US (San Jose)<br/>+14086380968,,96983619235# US (San Jose)</p>";
                string MeetingDate = "2021-05-31 17:30:00";
                string message = string.Empty;

                Attendee = Attendee.Replace(";", "<br/>");

                DateTime dtNowUTC = DateTime.Now; //change to Now upon deployment
                DateTime dtNow = DateTime.SpecifyKind(dtNowUTC, DateTimeKind.Utc);
                //TimeZoneInfo pacificZone = TimeZoneInfo.FindSystemTimeZoneById("America/Los_Angeles"); //linux
                TimeZoneInfo pacificZone = TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time"); //windows
                DateTime dtStart = DateTime.Parse(MeetingDate); //datetime without DateTimeKind specify
                DateTime dtStartUTC = DateTime.SpecifyKind(dtStart, DateTimeKind.Utc); //datetime specify DateTimeKind UTC - start date is in UTC timezone

                DateTime dtStartPST = TimeZoneInfo.ConvertTimeFromUtc(dtStartUTC, pacificZone); //datetime convert to PST timezone

                string formattedDateTime = dtStartPST.ToString("MM/dd/yyyy hh:mm tt");
                message = $"{Regarding} - Call Open <br/>Date & Time: {formattedDateTime}<br/>{format.ReplaceTagHtmlParagraph2(AddressLocation, true)}<br/>Expected Attendees:<br/>{Attendee}";

                AdminService adminService = new AdminService(_config);

                var result = await adminService.SendSkypeTest(message, "19:4c7257f3319a4c08a37183170a599cf3@thread.skype");
                return Ok(result);


            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error SendBizDevTestAsync {ex}", "ErrorSendBizDevTestAsync");
                return BadRequest();
            }



        }

        [AllowAnonymous]
        [HttpPost("sendImage")]
        public async Task<IActionResult> SendSkypeImageAsync()
        {

            try
            {
               
                AdminService adminService = new AdminService(_config);
                var result = await adminService.SendSkypeImage("test.png", "19:9ce1a5bf0e244cb08d60d5c93ca7e6c5@thread.skype");
                return Ok(result);


            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error SendBizDevTestAsync {ex}", "ErrorSendBizDevTestAsync");
                return BadRequest();
            }



        }

        [HttpPost("send3PMTest")]
        public async Task<IActionResult> Send3PMTestAsync()
        {

            try
            {
                FormatService format = new FormatService();
                string Attendee = "Levin jay;John Doe";
                string Regarding = "BizDev Test";
                string AddressLocation = "<p>Join Zoom Meeting<br/></p><p>https://zoom.us/j/96983619235?pwd=WjlLMVY2SXJLUi8xUmFJQm1USGNmUT09</p><p>Meeting ID: 969 8361 9235<br/>Passcode: 105175<br/>One tap mobile<br/>+16699006833,,96983619235# US (San Jose)<br/>+14086380968,,96983619235# US (San Jose)</p>";
                string MeetingDate = "2021-05-31 17:30:00";
                string message = string.Empty;

                Attendee = Attendee.Replace(";", "<br/>");

                DateTime dtNowUTC = DateTime.Now; //change to Now upon deployment
                DateTime dtNow = DateTime.SpecifyKind(dtNowUTC, DateTimeKind.Utc);
                //TimeZoneInfo pacificZone = TimeZoneInfo.FindSystemTimeZoneById("America/Los_Angeles"); //linux
                TimeZoneInfo pacificZone = TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time"); //windows
                DateTime dtStart = DateTime.Parse(MeetingDate); //datetime without DateTimeKind specify
                DateTime dtStartUTC = DateTime.SpecifyKind(dtStart, DateTimeKind.Utc); //datetime specify DateTimeKind UTC - start date is in UTC timezone

                DateTime dtStartPST = TimeZoneInfo.ConvertTimeFromUtc(dtStartUTC, pacificZone); //datetime convert to PST timezone

                string formattedDateTime = dtStartPST.ToString("MM/dd/yyyy hh:mm tt");
                //message = $"{Regarding} - Call Open <br/>Date & Time: {formattedDateTime}<br/>{format.ReplaceTagHtmlParagraph2(AddressLocation, true)}<br/>Expected Attendees:<br/>{Attendee}";
                string bell = $"(bell)(bell)(bell)";
                message = $"{bell}<br/><br/>Call | Test 10:00 AM - 11:00 AM<br/>Levin jay, John Doe<br/><br/>{bell}";

                AdminService adminService = new AdminService(_config);

                var result = await adminService.SendSkypeTest(message, "19:4c7257f3319a4c08a37183170a599cf3@thread.skype");
                return Ok(result);


            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error SendSkypeImageAsync {ex}", "ErrorSendSkypeImageAsync");
                return BadRequest();
            }



        }

    }
}
