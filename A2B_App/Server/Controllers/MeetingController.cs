
using A2B_App.Server.Data;
using A2B_App.Server.Services;
using A2B_App.Shared.Meeting;
using A2B_App.Shared.Skype;
using A2B_App.Shared.Sox;
using A2B_App.Shared.Time;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace A2B_App.Server.Controllers
{


    [Authorize]
    [Route("api/[controller]")]
    [ApiController]
    public class MeetingController : ControllerBase
    {

        private readonly IConfiguration _config;
        private readonly ILogger<MeetingController> _logger;
        private readonly IWebHostEnvironment _environment;
        private readonly SoxContext _soxContext;
        private readonly MeetingContext _meetingContext;

        public MeetingController(IConfiguration config,
            ILogger<MeetingController> logger,
            IWebHostEnvironment environment,
            SoxContext soxContext,
            MeetingContext meetingContext)
        {
            _config = config;
            _logger = logger;
            _environment = environment;
            _soxContext = soxContext;
            _meetingContext = meetingContext;
        }

        /// <summary>
        /// Get list of daily meeting
        /// </summary>
        /// <returns>(list)dailymeeting</returns>
        //[AllowAnonymous]
        [HttpGet("dailyMeeting")]
        public IActionResult GetListDailyMeeting()
        {
            try
            {
                string dtToday = DateTime.Now.ToString("yyyy-MM-dd");
                string dtNextDay = DateTime.Now.AddDays(1).ToString("yyyy-MM-dd");
                string dtDay = DateTime.Now.ToString("dddd");
                //var listDailyMeetings = _meetingContext.DailyMeeting.FromSqlRaw($"CALL a2bmeeting.sp_select_dailymeeting('{dtToday}', '{dtDay}', '{dtNextDay}');").AsEnumerable();
                string sqlQuery = $"CALL `a2bmeeting`.`sp_select_dailymeeting`('{dtToday}', '{dtDay}', '{dtNextDay}', 1);";
                var listDailyMeetings = _meetingContext.DailyMeeting.FromSqlRaw(sqlQuery).AsEnumerable();
                if (listDailyMeetings != null && listDailyMeetings.Any())
                {
                    return Ok(listDailyMeetings.ToList());
                }
                return NoContent();
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetListDailyMeeting {ex}", "ErrorGetListDailyMeeting");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetListDailyMeeting");
                return BadRequest();
            }
        }

        [HttpGet("dailyMeetingBizDev")]
        public IActionResult GetListDailyMeetingBizDev()
        {
            try
            {
                string dtToday = DateTime.Now.ToString("yyyy-MM-dd");
                //var listDailyMeetings = _meetingContext.DailyMeeting.FromSqlRaw($"CALL a2bmeeting.sp_select_dailymeeting('{dtToday}', '{dtDay}', '{dtNextDay}');").AsEnumerable();
                string sqlQuery = $"CALL `a2bmeeting`.`sp_select_dailymeeting_bizdev`('{dtToday}');";
                var listDailyMeetings = _meetingContext.DailyMeetingBizDev.FromSqlRaw(sqlQuery).AsEnumerable();
                if (listDailyMeetings != null && listDailyMeetings.Any())
                {
                    return Ok(listDailyMeetings.ToList());
                }
                return NoContent();
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetListDailyMeetingBizDev {ex}", "ErrorGetListDailyMeetingBizDev");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetListDailyMeetingBizDev");
                return BadRequest();
            }
        }

        [HttpGet("recordings")]
        public IActionResult GetAllRecordingGTM()
        {   
            try
            {
                string dtToday = DateTime.Now.ToString("yyyy-MM-dd");
                string dtNextDay = DateTime.Now.AddDays(1).ToString("yyyy-MM-dd");
                string dtDay = DateTime.Now.ToString("dddd");
               
                string sqlQuery = $"SELECT * FROM recording_gtm_details;";
                //var listRecordings = _meetingContext.recording_gtm_details.FromSqlRaw(sqlQuery).AsEnumerable();

                var listRecordings = _meetingContext.recording_gtm.Where( x => x.starttime.Contains(dtToday)).AsEnumerable();

                Console.WriteLine(listRecordings);
                if (listRecordings != null && listRecordings.Any())
                {
                    return Ok(listRecordings.ToList());
                }
                else
                {
                    return NoContent();
                }
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetListDailyMeeting {ex}", "ErrorGetListDailyMeeting");
                AdminService adminService = new AdminService(_config);
                //adminService.SendAlert(true, true, ex.ToString(), "GetListDailyMeeting");
                Console.WriteLine(ex.ToString());
                return BadRequest();
            }
        }
        
        [HttpGet("weeklyrecordings")]
        public IActionResult GetWeeklyRecordingGTM()
        {
            try
            {
                string dtToday = DateTime.Now.ToString("yyyy-MM-dd");
                string dtNextDay = DateTime.Now.AddDays(1).ToString("yyyy-MM-dd");
                string dtDay = DateTime.Now.ToString("dddd");

                DateTime startAtSaturday = DateTime.Now.AddDays(DayOfWeek.Saturday - DateTime.Now.DayOfWeek);
                DateTime startAtFriday = DateTime.Now.AddDays(DayOfWeek.Friday - DateTime.Now.DayOfWeek);
                DateTime startAtThursday = DateTime.Now.AddDays(DayOfWeek.Thursday - DateTime.Now.DayOfWeek);
                DateTime startAtWednesday = DateTime.Now.AddDays(DayOfWeek.Wednesday - DateTime.Now.DayOfWeek);
                DateTime startAtTuesday = DateTime.Now.AddDays(DayOfWeek.Tuesday - DateTime.Now.DayOfWeek);
                DateTime startAtMonday = DateTime.Now.AddDays(DayOfWeek.Monday - DateTime.Now.DayOfWeek);
                DateTime startAtSunday = DateTime.Now.AddDays(DayOfWeek.Sunday - DateTime.Now.DayOfWeek);

                string tostringMonday = startAtMonday.ToString("yyyy-MM-dd");
                string tostringSunday = startAtSunday.ToString("yyyy-MM-dd");
                string tostringTuesday = startAtTuesday.ToString("yyyy-MM-dd");
                string tostringWednesday = startAtWednesday.ToString("yyyy-MM-dd");
                string tostringThursday = startAtThursday.ToString("yyyy-MM-dd");
                string tostringFridayy = startAtFriday.ToString("yyyy-MM-dd");
                string tostringSaturday = startAtSaturday.ToString("yyyy-MM-dd");


                var listRecordings = _meetingContext.recording_gtm.Where( 
                    x => x.starttime.Contains(tostringSunday)
                    || x.starttime.Contains(tostringMonday)
                    || x.starttime.Contains(tostringTuesday)
                    || x.starttime.Contains(tostringWednesday)
                    || x.starttime.Contains(tostringThursday)
                    || x.starttime.Contains(tostringFridayy)
                    || x.starttime.Contains(tostringSaturday)).AsEnumerable();

                Console.WriteLine(listRecordings.ToList());
                if (listRecordings != null && listRecordings.Any())
                {
                    return Ok(listRecordings.ToList());
                }
                else
                {
                    return NoContent();
                }
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetListDailyMeeting {ex}", "ErrorGetListDailyMeeting");
                AdminService adminService = new AdminService(_config);
                //adminService.SendAlert(true, true, ex.ToString(), "GetListDailyMeeting");
                Console.WriteLine(ex.ToString());
                return BadRequest();
            }
        }
        
        [HttpPost("downloadlink")]
        public IActionResult DownloadLink(recording_gtm item)
        {
            try
            {
                string dtToday = DateTime.Now.ToString("yyyy-MM");
                string dtNextDay = DateTime.Now.AddDays(1).ToString("yyyy-MM-dd");
                string dtDay = DateTime.Now.ToString("dddd");

                var listRecordings = _meetingContext.recording_gtm_details.Where(x => x.meeting_id.Contains(item.meeting_id)).AsEnumerable();

                Console.WriteLine(listRecordings.ToList());
                if (listRecordings != null && listRecordings.Any())
                {
                    var id = listRecordings.FirstOrDefault();
                    return Ok(id.download_url);
                }
                else
                {
                    return NoContent();
                }
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetListDailyMeeting {ex}", "ErrorGetListDailyMeeting");
                AdminService adminService = new AdminService(_config);
                //adminService.SendAlert(true, true, ex.ToString(), "GetListDailyMeeting");
                Console.WriteLine(ex.ToString());
                return BadRequest();
            }
        }
        
        [HttpGet("monthlyrecordings")]
        public IActionResult GetMonthlyRecordingGTM()
        {
            try
            {
                string dtToday = DateTime.Now.ToString("yyyy-MM");
                string dtNextDay = DateTime.Now.AddDays(1).ToString("yyyy-MM-dd");
                string dtDay = DateTime.Now.ToString("dddd");

                var listRecordings = _meetingContext.recording_gtm.Where(x => x.starttime.Contains(dtToday)).AsEnumerable();

                Console.WriteLine(listRecordings.ToList());
                if (listRecordings != null && listRecordings.Any())
                {
                    return Ok(listRecordings.ToList());
                }
                else
                {
                    return NoContent();
                }
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetListDailyMeeting {ex}", "ErrorGetListDailyMeeting");
                AdminService adminService = new AdminService(_config);
                //adminService.SendAlert(true, true, ex.ToString(), "GetListDailyMeeting");
                Console.WriteLine(ex.ToString());
                return BadRequest();
            }
        }

        [HttpPost("fetch-attendees")]
        public async Task<IActionResult> FetchAttendees([FromBody] int meetingID)
        {
            var checkSoxTracker = _meetingContext.meeting
                    .Where(x =>
                        x.sys_id.Equals(meetingID)  
                    )
                     .AsNoTracking().Distinct().ToList();
            if(checkSoxTracker != null)
            {

                return Ok(checkSoxTracker);
            }
            return NoContent();

        }

        [HttpPost("daterangerecordings")]
        public IActionResult DateRangeRecordings(DateRangeCustom range)
        {
            try
            {
                string dtDay = DateTime.Now.ToString("dddd");
                string test11 = range.Start.ToString("yyyy-MM-dd");
                string test22 = range.End.ToString("yyyy-MM-dd");
                string query2 = $"CALL `a2bmeeting`.`sp_select_dailyrecording_gtm`('{test11}', '', '{test22}');";

                string sqlQuery2 = $"SELECT * FROM `recording_gtm`  WHERE `starttime` BETWEEN '{test11}'AND'{test22}';";

                var listDailyMeetings = _meetingContext.recording_gtm.FromSqlRaw(sqlQuery2).AsEnumerable();
                if (listDailyMeetings != null && listDailyMeetings.Any())
                {
                    return Ok(listDailyMeetings.ToList());
                }
                return NoContent();
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetListDailyMeeting {ex}", "ErrorGetListDailyMeeting");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetListDailyMeeting");
                return BadRequest();
            }
        }





        [HttpGet("weeklyMeeting")]
        public IActionResult GetWeekly()
        {
            try
            {
                string dtToday = DateTime.Now.ToString("yyyy-MM-dd");
                string dtNextDay = DateTime.Now.AddDays(1).ToString("yyyy-MM-dd");
                string dtDay = DateTime.Now.ToString("dddd");
                
                DateTime startAtMonday = DateTime.Now.AddDays(DayOfWeek.Monday - DateTime.Now.DayOfWeek);
                DateTime startAtSunday = DateTime.Now.AddDays(DayOfWeek.Sunday - DateTime.Now.DayOfWeek);

                string tostringMonday = startAtMonday.ToString("yyyy-MM-dd");
                string tostringSunday = startAtSunday.ToString("yyyy-MM-dd");

                string sqlQuery = $"CALL `a2bmeeting`.`sp_select_dailymeeting`('{tostringMonday}', '', '{tostringSunday}');";
                var listDailyMeetings = _meetingContext.DailyMeeting.FromSqlRaw(sqlQuery).AsEnumerable();
                if (listDailyMeetings != null && listDailyMeetings.Any())
                {
                    return Ok(listDailyMeetings.ToList());
                }
                return NoContent();
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetListDailyMeeting {ex}", "ErrorGetListDailyMeeting");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetListDailyMeeting");
                return BadRequest();
            }
        }

        [HttpGet("monthlyMeeting")]
        public IActionResult GetMonthly()
        {
            try
            {

                string dtToday = DateTime.Now.ToString("yyyy-MM-dd");
                string dtNextDay = DateTime.Now.AddDays(1).ToString("yyyy-MM-dd");
                string dtDay = DateTime.Now.ToString("dddd");
                //var listDailyMeetings = _meetingContext.DailyMeeting.FromSqlRaw($"CALL a2bmeeting.sp_select_dailymeeting('{dtToday}', '{dtDay}', '{dtNextDay}');").AsEnumerable();
               
                DateTime now = DateTime.Now;
                var startDate = new DateTime(now.Year, now.Month, 1);
                var endDate = startDate.AddMonths(1).AddDays(-1);
                string startDatee = startDate.ToString("yyyy-MM-dd");
                string endDateee = endDate.ToString("yyyy-MM-dd");

                string sqlQuery = $"CALL `a2bmeeting`.`sp_select_dailymeeting`('{startDatee}', '', '{endDateee}');";
                var listDailyMeetings = _meetingContext.DailyMeeting.FromSqlRaw(sqlQuery).AsEnumerable();
                if (listDailyMeetings != null && listDailyMeetings.Any())
                {
                    return Ok(listDailyMeetings.ToList());
                }
                return NoContent();
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetListDailyMeeting {ex}", "ErrorGetListDailyMeeting");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetListDailyMeeting");
                return BadRequest();
            }
        }

        [HttpPost("dateRangeMeeting")]
        public IActionResult DateRange(DateRangeCustom range)
        {
            try
            {
                string dtDay = DateTime.Now.ToString("dddd");
                string test11 = range.Start.ToString("yyyy-MM-dd");
                string test22 = range.End.ToString("yyyy-MM-dd");
                string query2 = $"CALL `a2bmeeting`.`sp_select_dailymeeting`('{test11}', '', '{test22}');";

                var listDailyMeetings = _meetingContext.DailyMeeting.FromSqlRaw(query2).AsEnumerable();
                if (listDailyMeetings != null && listDailyMeetings.Any())
                {
                    return Ok(listDailyMeetings.ToList());
                }
                return NoContent();
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetListDailyMeeting {ex}", "ErrorGetListDailyMeeting");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetListDailyMeeting");
                return BadRequest();
            }
        }

        /// <summary>
        /// Get list of daily meeting that is assign to specific presenter
        /// </summary>
        /// <param name="poc">Presenter name</param>
        /// <returns>(list)dailymeeting</returns>
        //[AllowAnonymous]
        [HttpGet("dailyMeetingPoc")]
        public IActionResult GetListDailyMeetingPoc(string poc)
        {
            try
            {
                string dtToday = DateTime.Now.ToString("yyyy-MM-dd");
                string dtNextDay = DateTime.Now.AddDays(1).ToString("yyyy-MM-dd");
                string dtDay = DateTime.Now.ToString("dddd");
                var listDailyMeetings = _meetingContext.DailyMeeting.FromSqlRaw($"CALL `a2bmeeting`.`sp_select_dailymeeting_presenter`('{dtToday}', '{dtDay}', '{dtNextDay}', '{poc}');").AsEnumerable();
                if (listDailyMeetings != null && listDailyMeetings.Any())
                {
                    return Ok(listDailyMeetings.ToList());
                }
                return NoContent();
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetListDailyMeetingPoc {ex}", "ErrorGetListDailyMeetingPoc");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetListDailyMeetingPoc");
                return BadRequest();
            }


        }

        /// <summary>
        /// Get list of daily meeting including attendee
        /// </summary>
        /// <returns>(list)dailymeetingAttendee</returns>
        //[AllowAnonymous]
        [HttpGet("dailyMeetingMember")]
        public IActionResult GetListDailyMeetingParticipant()
        {
            try
            {
                string dtToday = DateTime.Now.ToString("yyyy-MM-dd");
                string dtNextDay = DateTime.Now.AddDays(1).ToString("yyyy-MM-dd");
                string dtDay = DateTime.Now.ToString("dddd");
                //var listDailyMeetings = _meetingContext.DailyMeeting.FromSqlRaw($"CALL a2bmeeting.sp_select_dailymeeting('{dtToday}', '{dtDay}', '{dtNextDay}');").AsEnumerable();
                string sqlQuery = $"CALL `a2bmeeting`.`sp_select_dailymeetingandmember`('{dtToday}', '{dtDay}', '{dtNextDay}');";
                var listDailyMeetings = _meetingContext.DailyMeetingParticipant.FromSqlRaw(sqlQuery).AsEnumerable();
                if (listDailyMeetings != null && listDailyMeetings.Any())
                {
                    return Ok(listDailyMeetings.ToList());
                }
                return NoContent();
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetListDailyMeetingParticipant {ex}", "ErrorGetListDailyMeetingParticipant");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetListDailyMeeting");
                return BadRequest();
            }


        }

        [AllowAnonymous]
        [HttpPost("userDailyMeeting")]
        public IActionResult GetListDailyMeetingUser([FromBody] A2BPodioUser User)
        {
            try
            {
                string dtToday = DateTime.Now.ToString("yyyy-MM-dd");
                string dtNextDay = DateTime.Now.AddDays(1).ToString("yyyy-MM-dd");
                string dtDay = DateTime.Now.ToString("dddd");
                //var listDailyMeetings = _meetingContext.DailyMeeting.FromSqlRaw($"CALL a2bmeeting.sp_select_dailymeeting('{dtToday}', '{dtDay}', '{dtNextDay}');").AsEnumerable();
                string sqlQuery = $"CALL `a2bmeeting`.`sp_select_dailymeetingandmember`('{dtToday}', '{dtDay}', '{dtNextDay}');";
                var listDailyMeetings = _meetingContext.DailyMeetingParticipant.FromSqlRaw(sqlQuery).AsEnumerable();
                if (listDailyMeetings != null && listDailyMeetings.Any())
                {
                    return Ok(listDailyMeetings.ToList());
                }
                return NoContent();
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetListDailyMeetingUser {ex}", "ErrorGetListDailyMeetingUser");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetListDailyMeetingUser");
                return BadRequest();
            }


        }


        //[AllowAnonymous]
        [HttpGet("allMeetingNotif")]
        public IActionResult GetAllMeetingNotificationGC()
        {
            try
            {

                var meetingGc = _meetingContext.SkypeAddress
                    .Include(x => x.SkypeObj).ThenInclude(y => y.conversation)
                    .Include(z => z.ListKeyword)
                    .Take(150)
                    .AsNoTracking()
                    .AsEnumerable();
                return Ok(meetingGc.ToList());

            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetAllMeetingNotificationGC {ex}", "ErrorGetAllMeetingNotificationGC");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetAllMeetingNotificationGC");
                return BadRequest();
            }
        }


        //[AllowAnonymous]
        [HttpPost("saveMeetingNotif")]
        public async Task<IActionResult> SaveMeetingNotificationGC([FromBody] SkypeAddress skypeAdd)
        {
            try
            {
                
                if(skypeAdd != null)
                {
                    using (var context = _meetingContext.Database.BeginTransaction())
                    {
                        var checkSkypeAdd = _meetingContext.SkypeAddress.Where(id => id.Sys_id.Equals(skypeAdd.Sys_id))
                        .Include(y=>y.SkypeObj)
                        .Include(z=>z.ListKeyword)
                        .AsNoTracking()
                        .FirstOrDefault();
                        if(checkSkypeAdd != null) //update
                        {
                        
                        
                            


                            if (checkSkypeAdd.ListKeyword != null && checkSkypeAdd.ListKeyword.Any())
                            {
                                foreach (var item in checkSkypeAdd.ListKeyword)
                                {
                                    var checkKeyword = skypeAdd.ListKeyword.Where(id => id.Sys_id.Equals(item.Sys_id)).FirstOrDefault();
                                    if (checkKeyword == null)
                                        _meetingContext.Remove(item);
                                }
                            }


                            skypeAdd.Sys_id = checkSkypeAdd.Sys_id;
                            checkSkypeAdd = skypeAdd;
                            checkSkypeAdd.SkypeObj = skypeAdd.SkypeObj;
                            checkSkypeAdd.ListKeyword = skypeAdd.ListKeyword;

                            //checkSkypeAdd.ListKeyword = skypeAdd.ListKeyword;
                            //_meetingContext.Add(checkSkypeAdd.ListKeyword);


                            //if (checkSkypeAdd.SkypeObj != null)
                            //{
                            //    skypeAdd.SkypeObj.id = checkSkypeAdd.SkypeObj.id;
                            //    checkSkypeAdd.SkypeObj = skypeAdd.SkypeObj;
                            //    _meetingContext.Update(checkSkypeAdd.SkypeObj);
                            //}

                            _meetingContext.Update(checkSkypeAdd);
                            await _meetingContext.SaveChangesAsync();
                            context.Commit();
                            return Ok(checkSkypeAdd);
                      
                        }
                        else //create
                        {
                                _meetingContext.Add(skypeAdd);
                                await _meetingContext.SaveChangesAsync();
                                context.Commit();
                                return Ok(skypeAdd);
                        }
                    }
                }


                return NoContent();
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error SaveMeetingNotificationGC {ex}", "ErrorSaveMeetingNotificationGC");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetListDailyMeetingUser");
                return BadRequest();
            }
        }

        //[AllowAnonymous]
        [HttpPost("removeMeetingNotif")]
        public async Task<IActionResult> DeleteMeetingNotificationGC([FromBody] SkypeAddress skypeAdd)
        {
            try
            {

                if (skypeAdd != null)
                {
    
                    using (var context = _meetingContext.Database.BeginTransaction())
                    {
                        //Remove skypeAdd
                        var checkSkypeAdd = _meetingContext.SkypeAddress.Where(id => id.Sys_id.Equals(skypeAdd.Sys_id))
                            .Include(y => y.SkypeObj)
                            .Include(z => z.ListKeyword)
                            .FirstOrDefault();
                        if (checkSkypeAdd != null) //create if null
                        {
                            if(checkSkypeAdd.ListKeyword != null && checkSkypeAdd.ListKeyword.Any())
                            {
                                foreach (var item in checkSkypeAdd.ListKeyword)
                                {
                                    _meetingContext.Remove(item);
                                }
                            }
                           
                            if(checkSkypeAdd.SkypeObj != null)
                            {
                                _meetingContext.Remove(checkSkypeAdd.SkypeObj);
                            }
                            
                            _meetingContext.Remove(checkSkypeAdd);
                            await _meetingContext.SaveChangesAsync();
                            context.Commit();
                            return Ok();
                        }
                    }
                }


                return NoContent();
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error DeleteMeetingNotificationGC {ex}", "ErrorDeleteMeetingNotificationGC");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "DeleteMeetingNotificationGC");
                return BadRequest();
            }
        }

        //[AllowAnonymous]
        [HttpPost("testDailyMeetingMember")]
        public IActionResult GetTestListDailyMeetingParticipant([FromBody] List<dailymeetingAttendee> listDailyMeeting)
        {
            try
            {
                DateTime dtNowPST = new DateTime();
                DateTime dtStartPST = new DateTime();

                string strDtNow = "2021-04-13 16:30:00";
                DateTime dtNowUTC = DateTime.Parse(strDtNow);
                DateTime dtNow = DateTime.SpecifyKind(dtNowUTC, DateTimeKind.Utc);

                ////DateTime dtNowUTC = DateTime.Now.ToUniversalTime(); //if timezone is not UTC, convert to UTC
                //DateTime dtNowUTC = DateTime.Now; //if timezone is in UTC, no need to convert to UTC
                //DateTime dtNow = DateTime.SpecifyKind(dtNowUTC, DateTimeKind.Utc);

                TimeZoneInfo pacificZone;

                if (listDailyMeeting != null && listDailyMeeting.Any())
                {
                    foreach (var item in listDailyMeeting)
                    {

                        DateTime dtStart = DateTime.Parse(item.startdate.ToString()); //datetime without DateTimeKind specify
                        DateTime dtStartUTC = DateTime.SpecifyKind(dtStart, DateTimeKind.Utc); //datetime specify DateTimeKind UTC - start date is in UTC timezone

                        if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                        {
                            pacificZone = TimeZoneInfo.FindSystemTimeZoneById("America/Los_Angeles");
                            dtStartPST = TimeZoneInfo.ConvertTimeFromUtc(dtStartUTC, pacificZone); //datetime convert to PST timezone
                            dtNowPST = TimeZoneInfo.ConvertTimeFromUtc(dtNow, pacificZone);
                        }
                        else if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                        {
                            pacificZone = TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time");
                            dtStartPST = TimeZoneInfo.ConvertTimeFromUtc(dtStartUTC, pacificZone); //datetime convert to PST timezone
                            dtNowPST = TimeZoneInfo.ConvertTimeFromUtc(dtNow, pacificZone);
                        }


                        //dtNowPST = TimeZoneInfo.ConvertTimeBySystemTimeZoneId(dtNow, "Pacific Standard Time");
                        //dtStartPST = TimeZoneInfo.ConvertTimeBySystemTimeZoneId(dtStartUTC, "Pacific Standard Time"); //datetime convert to PST timezone

                        DateTime dtMeetingPST = new DateTime(dtNowPST.Year, dtNowPST.Month, dtNowPST.Day, dtStartPST.Hour, dtStartPST.Minute, dtStartPST.Second);
                        Debug.WriteLine($"-------------------------------------------------------------------------------");
                        Debug.WriteLine($"TITLE: {item.title}");
                        Debug.WriteLine($"{dtMeetingPST} >= {dtNowPST.AddMinutes(28)} && {dtMeetingPST} <= {dtNowPST.AddMinutes(32)}");
                        if (dtMeetingPST >= dtNowPST.AddMinutes(28) && dtMeetingPST <= dtNowPST.AddMinutes(32))
                        {
                            Debug.WriteLine($"Init notification for {item.title}");
                            //For all GC
                            var listGcAll = _meetingContext.SkypeAddress.Where(x => x.IsAllGC.Equals(true) && x.IsEnabled.Equals(true))
                                .Include(y => y.SkypeObj).ThenInclude(z => z.conversation)
                                .AsNoTracking();
                            if (listGcAll != null && listGcAll.Any())
                            {
                                foreach (var itemGcAll in listGcAll)
                                {
                                    Debug.WriteLine($"[ALL] Sending notification to {itemGcAll.SkypeObj.conversation.id}");
                                    //SkypeBot skypeBot = new SkypeBot(_config);
                                    //Skype skypeParam = new Skype();
                                    ////skypeParam.Address = _config["Skype:ErrorGc"];
                                    //skypeParam.Address = itemGcAll.SkypeObj.conversation.id;
                                    //skypeParam.Message = $"[Test] {item.title} - Call Open {item.pass_code} <br/> {item.meeting_link} <br/> Expected Attendees:<br/> {teamMember} <br/> Other Expected Attendees:<br/> {otherParticipant}";
                                    //skypeBot.SendSkypeNotif(skypeParam);
                                }
                            }

                            //For keyword base
                            var listGcKeyword = _meetingContext.SkypeAddress.Where(x => x.IsAllGC.Equals(false) && x.IsEnabled.Equals(true))
                                .Include(y => y.SkypeObj).ThenInclude(z => z.conversation)
                                .Include(w => w.ListKeyword)
                                .AsNoTracking();
                            if (listGcKeyword != null && listGcKeyword.Any())
                            {
                                foreach (var itemAddress in listGcKeyword)
                                {

                                    if (itemAddress.ListKeyword != null && itemAddress.ListKeyword.Any())
                                    {
                                        var checkIfContains = itemAddress.ListKeyword.Where(x => item.title.ToLower().Contains(x.Keyword.ToLower())).FirstOrDefault();
                                        if (checkIfContains != null)
                                        {
                                            Debug.WriteLine($"[Keyword] Sending notification to {itemAddress.SkypeObj.conversation.id}");
                                            //SkypeBot skypeBot = new SkypeBot(_config);
                                            //Skype skypeParam = new Skype();
                                            ////skypeParam.Address = _config["Skype:ErrorGc"];
                                            //skypeParam.Address = itemAddress.SkypeObj.conversation.id;
                                            //skypeParam.Message = $"[Test] {item.title} - Call Open {item.pass_code} <br/> {item.meeting_link} <br/> Expected Attendees:<br/> {teamMember} <br/> Other Expected Attendees:<br/> {otherParticipant}";
                                            //skypeBot.SendSkypeNotif(skypeParam);
                                        }
                                    }
                                }
                            }
                        }


                    }
                }

                return Ok();
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetListDailyMeetingParticipant {ex}", "ErrorGetListDailyMeetingParticipant");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetListDailyMeeting");
                return BadRequest();
            }


        }

    }
}
