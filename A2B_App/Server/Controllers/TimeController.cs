using A2B_App.Server.Data;
using A2B_App.Server.Services;
using A2B_App.Shared.Sox;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Linq;

namespace A2B_App.Server.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class TimeController : ControllerBase
    {
        private readonly IConfiguration _config;
        private readonly ILogger<TimeController> _logger;
        private readonly IWebHostEnvironment _environment;
        private readonly TimeContext _timeContext;

        public TimeController(IConfiguration config, ILogger<TimeController> logger, IWebHostEnvironment environment, TimeContext timeContext)
        {
            _config = config;
            _timeContext = timeContext;
            _environment = environment;
            _logger = logger;
        }

        [AllowAnonymous]
        [HttpGet("podioUser")]
        public IActionResult GetPodioUserA2B(string Email)
        {
            try
            {
                if(Email != null)
                {
                    //var listDailyMeetings = _meetingContext.DailyMeeting.FromSqlRaw($"CALL a2bmeeting.sp_select_dailymeeting('{dtToday}', '{dtDay}', '{dtNextDay}');").AsEnumerable();
                    string sqlQuery = $"CALL `podiodb`.`sp_get_podiouser`('{Email}');";
                    var podioUser = _timeContext.A2BPodioUser.FromSqlRaw(sqlQuery).AsEnumerable().FirstOrDefault();
                    if (podioUser != null)
                    {
                        return Ok(podioUser);
                    }
                }
                
                return NoContent();
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetPodioUserA2B {ex}", "ErrorGetPodioUserA2B");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetPodioUserA2B");
                return BadRequest();
            }


        }


    }
}
