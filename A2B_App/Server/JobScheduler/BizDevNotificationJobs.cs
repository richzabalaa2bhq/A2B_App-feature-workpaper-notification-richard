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
using A2B_App.Server.Services;

namespace A2B_App.Server.JobScheduler
{
    public class BizDevNotificationJobs : IJob
    {
        private readonly IServiceProvider _provider;
        private readonly IEmailSender _emailSender;
        private readonly IConfiguration _config;
        private readonly IServiceScopeFactory _scopeFactory;

        public BizDevNotificationJobs(IServiceProvider provider, IEmailSender emailSender, IConfiguration config, IServiceScopeFactory scopeFactory)
        {
            _provider = provider;
            _emailSender = emailSender;
            _config = config;
            _scopeFactory = scopeFactory;

        }

        public Task Execute(IJobExecutionContext context)
        {
            //Logs($"{DateTime.Now} [Reminders Service called]" + Environment.NewLine);
            Debug.WriteLine($"{DateTime.Now} [BizDev notification service called]" + Environment.NewLine);
            bool.TryParse(_config["MeetingNotification:Enabled"], out bool isLoggingEnabled);

            if (isLoggingEnabled)
                FileLog.Write($"{DateTime.Now} [BizDev notification service called]", "CronBizDevNotif");

            try
            {

                if (bool.TryParse(_config["MeetingNotification:Enabled"], out bool isEnabled) && isEnabled) //check if enabled in appsettings.json
                {

                    using (var scope = _scopeFactory.CreateScope())
                    {
                        FormatService format = new FormatService();
                        var _meetingContext = scope.ServiceProvider.GetRequiredService<MeetingContext>();

                        List<dailymeetingBizDev> listBizDevMeeting = new List<dailymeetingBizDev>();

                        //if timezone is in UTC, no need to convert to UTC
                        DateTime dtNowUTC = DateTime.Now; //change to Now upon deployment
                        DateTime dtNow = DateTime.SpecifyKind(dtNowUTC, DateTimeKind.Utc);

                        TimeZoneInfo pacificZone = TimeZoneInfo.FindSystemTimeZoneById("America/Los_Angeles"); //linux
                        //TimeZoneInfo pacificZone = TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time"); //windows

                        // when we exit the using block, 
                        // the IServiceScope will dispose itself 
                        // and dispose all of the services that it resolved.
                        string dtToday = DateTime.Now.ToString("yyyy-MM-dd");
                        //var listDailyMeetings = _meetingContext.DailyMeetingParticipant.FromSqlRaw($"CALL `a2bmeeting`.`sp_select_dailymeetingandmember`('{dtToday}', '{dtDay}', '{dtNextDay}');").AsEnumerable();
                        string sqlQuery = $"CALL `a2bmeeting`.`sp_select_dailymeeting_bizdev`('{dtToday}');";
                        var listDailyMeetings = _meetingContext.DailyMeetingBizDev.FromSqlRaw(sqlQuery).AsEnumerable();
                        if (listDailyMeetings != null && listDailyMeetings.Any())
                        {
                            foreach (var item in listDailyMeetings)
                            {

                                DateTime dtStart = DateTime.Parse(item.DateTimeofMeeting.ToString()); //datetime without DateTimeKind specify
                                DateTime dtStartUTC = DateTime.SpecifyKind(dtStart, DateTimeKind.Utc); //datetime specify DateTimeKind UTC - start date is in UTC timezone

                                DateTime dtStartPST = TimeZoneInfo.ConvertTimeFromUtc(dtStartUTC, pacificZone); //datetime convert to PST timezone
                                DateTime dtNowPST = TimeZoneInfo.ConvertTimeFromUtc(dtNow, pacificZone);
 
                                DateTime dtMeetingPST = new DateTime(dtNowPST.Year, dtNowPST.Month, dtNowPST.Day, dtStartPST.Hour, dtStartPST.Minute, dtStartPST.Second);
                                string formattedDateTime = dtStartPST.ToString("MM/dd/yyyy hh:mm tt");

                                if (isLoggingEnabled)
                                {
                                    FileLog.Write($"{DateTime.Now} [Validate meeting title] - {item.Regarding}", "CronBizDevValidate");
                                    FileLog.Write($"{DateTime.Now} [Validate condition] - {dtMeetingPST} >= {dtNowPST.AddMinutes(28)} && {dtMeetingPST} <= {dtNowPST.AddMinutes(32)}", "CronBizDevValidate");
                                }
                                if (dtMeetingPST >= dtNowPST.AddMinutes(28) && dtMeetingPST <= dtNowPST.AddMinutes(32))
                                {

                                    if (isLoggingEnabled)
                                    {
                                        FileLog.Write($"{DateTime.Now} [Meeting notification found] - {item.Regarding}", "CronBizDevNotif");
                                        FileLog.Write($"{DateTime.Now} [Meeting notification condition] - {dtMeetingPST} >= {dtNowPST.AddMinutes(28)} && {dtMeetingPST} <= {dtNowPST.AddMinutes(32)}", "CronMeetingNotif");
                                    }

                                    item.DateTimeofMeeting = formattedDateTime; //update date and time with custom format
                                    listBizDevMeeting.Add(item); //save in list for later use

                                }


                            }
                        }


                        if (listBizDevMeeting != null && listBizDevMeeting.Any())
                        {
                            foreach (var item in listBizDevMeeting)
                            {
                                //string emailCc = _config.GetSection("Email").GetSection("SODSoxRox").GetSection("EmailCc").Value;
                                string teamMember = string.Empty;
                                string otherParticipant = string.Empty;

                                if (item.MeetingWith != null && item.MeetingWith.ToString() != string.Empty)
                                    teamMember = item.MeetingWith.Replace(";", "<br/>");

                                //if (item.A2Q2Attendees != null && item.A2Q2Attendees.ToString() != string.Empty)
                                //    otherParticipant = item.A2Q2Attendees.Replace(",", "<br/>");


                                #region Send to Skype

                                //For all BizDev GC
                                var listGcAll = _meetingContext.SkypeAddress.Where(x => 
                                    x.IsAllGC.Equals(true) && 
                                    x.IsEnabled.Equals(true) &&
                                    x.IsBizDev.Equals(true)
                                )
                                    .Include(y => y.SkypeObj).ThenInclude(z => z.conversation)
                                    .AsNoTracking();
                                if (listGcAll != null && listGcAll.Any())
                                {
                                    foreach (var itemGcAll in listGcAll)
                                    {
                                        if (isLoggingEnabled)
                                            FileLog.Write($"{DateTime.Now} [Sending Notification] - [All] {item.Regarding}", "CronBizDevNotif");

                                        SkypeBot skypeBot = new SkypeBot(_config);
                                        Skype skypeParam = new Skype();
                                        //skypeParam.Address = _config["Skype:ErrorGc"];
                                        skypeParam.Address = itemGcAll.SkypeObj.conversation.id;
                                        skypeParam.Message = $"{item.Regarding} - Call Open <br/>Date & Time: {item.DateTimeofMeeting}<br/>{format.ReplaceTagHtmlParagraph2(item.MeetingLocationAddress, true)}<br/>Expected Attendees:<br/>{teamMember}";
                                        skypeBot.SendSkypeNotif(skypeParam);
                                    }
                                }

                                //For keyword base
                                var listGcKeyword = _meetingContext.SkypeAddress.Where(x => 
                                    x.IsAllGC.Equals(false) && 
                                    x.IsEnabled.Equals(true) && 
                                    x.IsBizDev
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
                                            var checkIfContains = itemAddress.ListKeyword.Where(x => item.Regarding.ToLower().Contains(x.Keyword.ToLower())).FirstOrDefault();
                                            if (checkIfContains != null)
                                            {
                                                if (isLoggingEnabled)
                                                    FileLog.Write($"{DateTime.Now} [Sending Notification] - [Keyword] {item.Regarding}", "CronBizDevNotif");

                                                SkypeBot skypeBot = new SkypeBot(_config);
                                                Skype skypeParam = new Skype();
                                                //skypeParam.Address = _config["Skype:ErrorGc"];
                                                skypeParam.Address = itemAddress.SkypeObj.conversation.id;
                                                skypeParam.Message = $"{item.Regarding} - Call Open <br/>Date & Time: {item.DateTimeofMeeting}<br/>{format.ReplaceTagHtmlParagraph2(item.MeetingLocationAddress, true)}<br/>Expected Attendees:<br/>{teamMember}";
                                                skypeBot.SendSkypeNotif(skypeParam);
                                            }


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
                FileLog.Write($"Error BizDevNotificationTask {ex}", "ErrorBizDevNotificationTask");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "BizDevNotificationTask");
            }


            return Task.CompletedTask;
        }


    }
}
