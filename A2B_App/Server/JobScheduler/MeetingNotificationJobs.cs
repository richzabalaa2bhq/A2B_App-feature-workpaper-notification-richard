using A2B_App.Server.Data;
using A2B_App.Server.Services;
using A2B_App.Shared.Meeting;
using A2B_App.Shared.Skype;
using A2B_App.Shared.Sox;
using Microsoft.AspNetCore.Identity.UI.Services;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Quartz;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace A2B_App.Server.JobScheduler
{
    public class MeetingNotificationJobs : IJob
    {
        private readonly IServiceProvider _provider;
        private readonly IEmailSender _emailSender;
        private readonly IConfiguration _config;
        private readonly IServiceScopeFactory _scopeFactory;

        public MeetingNotificationJobs(IServiceProvider provider, IEmailSender emailSender, IConfiguration config, IServiceScopeFactory scopeFactory)
        {
            _provider = provider;
            _emailSender = emailSender;
            _config = config;
            _scopeFactory = scopeFactory;

        }
        public Task Execute(IJobExecutionContext context)
        {
            //Logs($"{DateTime.Now} [Reminders Service called]" + Environment.NewLine);
            Debug.WriteLine($"{DateTime.Now} [Meeting notification service called]" + Environment.NewLine);
            bool.TryParse(_config["MeetingNotification:Enabled"], out bool isLoggingEnabled);

            if(isLoggingEnabled)
                FileLog.Write($"{DateTime.Now} [Meeting notification service called]", "CronMeetingNotif");

            try
            {
                
                if (bool.TryParse(_config["MeetingNotification:Enabled"], out bool isEnabled) && isEnabled) //check if enabled in appsettings.json
                {
                    
                    using (var scope = _scopeFactory.CreateScope())
                    {
                        var _meetingContext = scope.ServiceProvider.GetRequiredService<MeetingContext>();

                        List<dailymeetingAttendee> listMeetingAttendee = new List<dailymeetingAttendee>();

                        //DateTime dtNowPST = new DateTime();
                        //DateTime dtStartPST = new DateTime();

                        ////if timezone is not UTC, convert to UTC
                        //DateTime dtNowUTC = DateTime.Now.ToUniversalTime();
                        //DateTime dtNow = DateTime.SpecifyKind(dtNowUTC, DateTimeKind.Utc);

                        //if timezone is in UTC, no need to convert to UTC
                        DateTime dtNowUTC = DateTime.Now; 
                        DateTime dtNow = DateTime.SpecifyKind(dtNowUTC, DateTimeKind.Utc);

                        TimeZoneInfo pacificZone = TimeZoneInfo.FindSystemTimeZoneById("America/Los_Angeles"); //linux
                        //TimeZoneInfo pacificZone = TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time"); //windows

                        // when we exit the using block, 
                        // the IServiceScope will dispose itself 
                        // and dispose all of the services that it resolved.
                        string dtToday = DateTime.Now.ToString("yyyy-MM-dd");
                        string dtNextDay = DateTime.Now.AddDays(1).ToString("yyyy-MM-dd");
                        string dtDay = DateTime.Now.ToString("dddd");
                        //var listDailyMeetings = _meetingContext.DailyMeetingParticipant.FromSqlRaw($"CALL `a2bmeeting`.`sp_select_dailymeetingandmember`('{dtToday}', '{dtDay}', '{dtNextDay}');").AsEnumerable();
                        string sqlQuery = $"CALL `a2bmeeting`.`sp_select_dailymeetingandmember`('{dtToday}', '{dtDay}', '{dtNextDay}');";
                        var listDailyMeetings = _meetingContext.DailyMeetingParticipant.FromSqlRaw(sqlQuery).AsEnumerable();
                        if (listDailyMeetings != null && listDailyMeetings.Any())
                        {
                            foreach (var item in listDailyMeetings)
                            {
                                
                                DateTime dtStart = DateTime.Parse(item.startdate.ToString()); //datetime without DateTimeKind specify
                                DateTime dtStartUTC = DateTime.SpecifyKind(dtStart, DateTimeKind.Utc); //datetime specify DateTimeKind UTC - start date is in UTC timezone

                                DateTime dtStartPST = TimeZoneInfo.ConvertTimeFromUtc(dtStartUTC, pacificZone); //datetime convert to PST timezone
                                DateTime dtNowPST = TimeZoneInfo.ConvertTimeFromUtc(dtNow, pacificZone);

                                //if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                                //{
                                //    //pacificZone = TimeZoneInfo.FindSystemTimeZoneById("America/Los_Angeles");
                                //    dtStartPST = TimeZoneInfo.ConvertTimeFromUtc(dtStartUTC, pacificZone); //datetime convert to PST timezone
                                //    dtNowPST = TimeZoneInfo.ConvertTimeFromUtc(dtNow, pacificZone);
                                //}
                                //else if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                                //{
                                //    //pacificZone = TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time");
                                //    dtStartPST = TimeZoneInfo.ConvertTimeFromUtc(dtStartUTC, pacificZone); //datetime convert to PST timezone
                                //    dtNowPST = TimeZoneInfo.ConvertTimeFromUtc(dtNow, pacificZone);
                                //}


                                //dtNowPST = TimeZoneInfo.ConvertTimeBySystemTimeZoneId(dtNow, "Pacific Standard Time");
                                //dtStartPST = TimeZoneInfo.ConvertTimeBySystemTimeZoneId(dtStartUTC, "Pacific Standard Time"); //datetime convert to PST timezone

                                DateTime dtMeetingPST = new DateTime(dtNowPST.Year, dtNowPST.Month, dtNowPST.Day, dtStartPST.Hour, dtStartPST.Minute, dtStartPST.Second);

                                if (isLoggingEnabled)
                                {
                                    FileLog.Write($"{DateTime.Now} [Validate meeting title] - {item.title}", "CronMeetingValidate");
                                    FileLog.Write($"{DateTime.Now} [Validate condition] - {dtMeetingPST} >= {dtNowPST.AddMinutes(28)} && {dtMeetingPST} <= {dtNowPST.AddMinutes(32)}", "CronMeetingValidate");
                                }
                                if (dtMeetingPST >= dtNowPST.AddMinutes(28) && dtMeetingPST <= dtNowPST.AddMinutes(32))
                                {

                                    if (isLoggingEnabled)
                                    {
                                        FileLog.Write($"{DateTime.Now} [Meeting notification found] - {item.title}", "CronMeetingNotif");
                                        FileLog.Write($"{DateTime.Now} [Meeting notification condition] - {dtMeetingPST} >= {dtNowPST.AddMinutes(28)} && {dtMeetingPST} <= {dtNowPST.AddMinutes(32)}", "CronMeetingNotif");
                                    }

                                    listMeetingAttendee.Add(item); //save in list for later use

                                }


                            }
                        }


                        if(listMeetingAttendee != null && listMeetingAttendee.Any())
                        {
                            foreach (var item in listMeetingAttendee)
                            {
                                //string emailCc = _config.GetSection("Email").GetSection("SODSoxRox").GetSection("EmailCc").Value;
                                string teamMember = string.Empty;
                                string otherParticipant = string.Empty;

                                if (item.team_member != null && item.team_member.ToString() != string.Empty)
                                    teamMember = item.team_member.Replace(";", "<br/>");

                                if (item.other_participant != null && item.other_participant.ToString() != string.Empty)
                                    otherParticipant = item.other_participant.Replace(",", "<br/>");


                                #region Send to Skype

                                //For all GC
                                var listGcAll = _meetingContext.SkypeAddress.Where(x => 
                                        x.IsAllGC.Equals(true) && 
                                        x.IsEnabled.Equals(true) &&
                                        x.IsBizDev.Equals(false)
                                    )
                                        .Include(y => y.SkypeObj).ThenInclude(z => z.conversation)
                                        .AsNoTracking();
                                if (listGcAll != null && listGcAll.Any())
                                {
                                    foreach (var itemGcAll in listGcAll)
                                    {
                                        if (isLoggingEnabled)
                                            FileLog.Write($"{DateTime.Now} [Sending Notification] - [All] {item.title}", "CronMeetingNotif");

                                        SkypeBot skypeBot = new SkypeBot(_config);
                                        Skype skypeParam = new Skype();
                                        //skypeParam.Address = _config["Skype:ErrorGc"];
                                        skypeParam.Address = itemGcAll.SkypeObj.conversation.id;
                                        skypeParam.Message = $"{item.title} - Call Open [{item.pass_code}] <br/> {item.meeting_link} <br/> Expected Attendees:<br/> {teamMember} <br/> Other Expected Attendees:<br/> {otherParticipant}";
                                        skypeBot.SendSkypeNotif(skypeParam);
                                    }
                                }

                                //For keyword base
                                var listGcKeyword = _meetingContext.SkypeAddress.Where(x => 
                                    x.IsAllGC.Equals(false) && 
                                    x.IsEnabled.Equals(true) &&
                                    x.IsAllGC.Equals(false)
                                )
                                    .Include(y => y.SkypeObj).ThenInclude(z => z.conversation)
                                    .Include(w => w.ListKeyword)
                                    .AsNoTracking();
                                if (listGcKeyword != null && listGcKeyword.Any())
                                {
                                    foreach (var itemAddress in listGcKeyword)
                                    {

                                        if (itemAddress.ListKeyword != null && itemAddress.ListKeyword.Any())
                                        {
                                            //Check if keyword string contains in title
                                            var checkIfContains = itemAddress.ListKeyword.Where(x => item.title.ToLower().Contains(x.Keyword.ToLower())).FirstOrDefault();
                                            if (checkIfContains != null)
                                            {
                                                if (isLoggingEnabled)
                                                    FileLog.Write($"{DateTime.Now} [Sending Notification] - [Keyword] {item.title}", "CronMeetingNotif");

                                                SkypeBot skypeBot = new SkypeBot(_config);
                                                Skype skypeParam = new Skype();
                                                //skypeParam.Address = _config["Skype:ErrorGc"];
                                                skypeParam.Address = itemAddress.SkypeObj.conversation.id;
                                                skypeParam.Message = $"{item.title} - Call Open [{item.pass_code}] <br/>{item.meeting_link}<br/>Expected Attendees:<br/> {teamMember}<br/>Other Expected Attendees:<br/> {otherParticipant}";
                                                skypeBot.SendSkypeNotif(skypeParam);
                                            }

                                            ////Extract words in meeting title and compare to keyword
                                            //List<string> words = GetWords(item.title).ToList();
                                            //bool hasMatch = itemAddress.ListKeyword.Select(x => x.Keyword)
                                            //     .Intersect(words)
                                            //     .Any();
                                            ////var check = keyword.Except(words).ToList();
                                            //if (hasMatch)
                                            //{
                                            //    if (isLoggingEnabled)
                                            //        FileLog.Write($"{DateTime.Now} [Sending Notification] - [Keyword] {item.title}", "CronMeetingNotif");

                                            //    SkypeBot skypeBot = new SkypeBot(_config);
                                            //    Skype skypeParam = new Skype();
                                            //    //skypeParam.Address = _config["Skype:ErrorGc"];
                                            //    skypeParam.Address = itemAddress.SkypeObj.conversation.id;
                                            //    skypeParam.Message = $"{item.title} - Call Open [{item.pass_code}] <br/> {item.meeting_link} <br/> Expected Attendees:<br/> {teamMember} <br/> Other Expected Attendees:<br/> {otherParticipant}";
                                            //    skypeBot.SendSkypeNotif(skypeParam);
                                            //}

                                        }



                                    }
                                }

                                #endregion


                            }
                        }


                    }

                }



            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write($"Error MeetingNotificationTask {ex}", "ErrorMeetingNotificationTask");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "MeetingNotificationTask");
            }


            return Task.CompletedTask;
        }

        /// <summary>
        /// Get words from string
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        static string[] GetWords(string input)
        {
            MatchCollection matches = Regex.Matches(input, @"\b[\w']*\b");

            var words = from m in matches.Cast<Match>()
                        where !string.IsNullOrEmpty(m.Value)
                        select TrimSuffix(m.Value);

            return words.ToArray();
        }

        static string TrimSuffix(string word)
        {
            int apostropheLocation = word.IndexOf('\'');
            if (apostropheLocation != -1)
            {
                word = word.Substring(0, apostropheLocation);
            }

            return word;
        }

    }
}
