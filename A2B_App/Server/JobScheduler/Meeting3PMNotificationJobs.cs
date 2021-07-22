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
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace A2B_App.Server.JobScheduler
{
    public class Meeting3PMNotificationJobs : IJob
    {
        private readonly IServiceProvider _provider;
        private readonly IEmailSender _emailSender;
        private readonly IConfiguration _config;
        private readonly IServiceScopeFactory _scopeFactory;

        public Meeting3PMNotificationJobs(IServiceProvider provider, IEmailSender emailSender, IConfiguration config, IServiceScopeFactory scopeFactory)
        {
            _provider = provider;
            _emailSender = emailSender;
            _config = config;
            _scopeFactory = scopeFactory;

        }
        public Task Execute(IJobExecutionContext context)
        {
            //Logs($"{DateTime.Now} [Reminders Service called]" + Environment.NewLine);
            Debug.WriteLine($"{DateTime.Now} [Meeting 3PM notification service called]" + Environment.NewLine);
            bool.TryParse(_config["Meeting3PMNotification:Enabled"], out bool isLoggingEnabled);

            if(isLoggingEnabled)
                FileLog.Write($"{DateTime.Now} [Meeting 3PM notification service called]", "CronMeeting3PMNotif");

            try
            {
                
                if (bool.TryParse(_config["Meeting3PMNotification:Enabled"], out bool isEnabled) && isEnabled) //check if enabled in appsettings.json
                {
                    
                    using (var scope = _scopeFactory.CreateScope())
                    {
                        var _meetingContext = scope.ServiceProvider.GetRequiredService<MeetingContext>();

                        List<dailymeetingAttendee> listMeetingAttendee = new List<dailymeetingAttendee>();
                        List<t_employee> listEmployee = new List<t_employee>();

                        //DateTime dtNowPST = new DateTime();
                        //DateTime dtStartPST = new DateTime();

                        ////if timezone is not UTC, convert to UTC
                        //DateTime dtNowUTC = DateTime.Now.ToUniversalTime();
                        //DateTime dtNow = DateTime.SpecifyKind(dtNowUTC, DateTimeKind.Utc);

                        //if timezone is in UTC, no need to convert to UTC
                        DateTime dtNowUTC = DateTime.Now; 
                        DateTime dtNow = DateTime.SpecifyKind(dtNowUTC, DateTimeKind.Utc);

                       TimeZoneInfo pacificZone = TimeZoneInfo.FindSystemTimeZoneById("America/Los_Angeles"); //linux
                        // TimeZoneInfo pacificZone = TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time"); //windows

                        // when we exit the using block, 
                        // the IServiceScope will dispose itself 
                        // and dispose all of the services that it resolved.
                        string dtToday = DateTime.Now.AddDays(1).ToString("yyyy-MM-dd");
                        string dtNextDay = DateTime.Now.AddDays(2).ToString("yyyy-MM-dd");
                        string dtDay = DateTime.Now.AddDays(1).ToString("dddd");
                        //var listDailyMeetings = _meetingContext.DailyMeetingParticipant.FromSqlRaw($"CALL `a2bmeeting`.`sp_select_dailymeetingandmember`('{dtToday}', '{dtDay}', '{dtNextDay}');").AsEnumerable();
                        string sqlQuery = $"CALL `a2bmeeting`.`sp_select_dailymeetingandmember`('{dtToday}', '{dtDay}', '{dtNextDay}');";
                        var listDailyMeetings = _meetingContext.DailyMeetingParticipant.FromSqlRaw(sqlQuery).AsEnumerable();

                        string sqlQueryEmployee = $"CALL `a2bmeeting`.`sp_get_employee`();";
                        listEmployee = _meetingContext.ClientParticipant.FromSqlRaw(sqlQueryEmployee).ToList();

                        if (listDailyMeetings != null && listDailyMeetings.Any())
                        {
                            foreach (var item in listDailyMeetings)
                            { 
                                listMeetingAttendee.Add(item); //save in list for later use
                            }
                        }


                        if(listMeetingAttendee != null && listMeetingAttendee.Any())
                        {
                            string bell = $"(bell)(bell)(bell)";

                            //#region Send to Skype ALL GC
                            ////For all GC
                            var listGcAll = _meetingContext.SkypeAddress.Where(x =>
                                    x.IsAllGC.Equals(true) &&
                                    x.IsEnabled.Equals(true) &&
                                    x.IsBizDev.Equals(false) &&
                                    x.Is3PMNotification.Equals(true)
                                )
                                    .Include(y => y.SkypeObj).ThenInclude(z => z.conversation)
                                    .AsNoTracking();
                            if (listGcAll != null && listGcAll.Any())
                           {
                                foreach (var itemGcAll in listGcAll)
                                {

                                    StringBuilder sb = new StringBuilder();
                                    sb.Append(bell);
                                    sb.Append("<br/><br/>");
                                    foreach (var item in listMeetingAttendee)
                                    {
                                        //string emailCc = _config.GetSection("Email").GetSection("SODSoxRox").GetSection("EmailCc").Value;
                                        string teamMember = string.Empty;
                                        string otherParticipant = string.Empty;
                                        StringBuilder sbParticipant = new StringBuilder();

                                        if (item.team_member != null && item.team_member.ToString() != string.Empty)
                                        {
                                            //split and check if it contains in listEmployee
                                            string[] member = item.team_member.Split(';');
                                          if(member.Any())
                                            {
                                                foreach (var itemMember in member)
                                                {
                                                    if (listEmployee.Where(mem => mem.f_name.ToLower().Equals(itemMember.ToLower())).FirstOrDefault() != null)
                                                        sbParticipant.Append($"{itemMember}, ");
                                                }
                                            }
                                        }
                                           

                                        if (item.other_participant != null && item.other_participant.ToString() != string.Empty)
                                        {
                                            //split and check if it contains in listEmployee
                                            string[] member = item.other_participant.Split(',');
                                            if (member.Any())
                                            {
                                                foreach (var itemMember in member)
                                                {
                                                    if (listEmployee.Where(mem => mem.f_email.ToLower().Equals(itemMember.ToLower())).FirstOrDefault() != null)
                                                        sbParticipant.Append($"{itemMember}, ");
                                                }
                                            }
                                        }
                                            
                                        if(sbParticipant.Length > 2)
                                        {
                                            sbParticipant.Length -= 2; //Remove last 2 element
                                        }


                                        if (isLoggingEnabled)
                                            FileLog.Write($"{DateTime.Now} [Sending 3PM Notification] - [All] {item.title}", "CronMeeting3PMNotif");

                                        //string participant = teamMember + (teamMember != string.Empty ? $", {otherParticipant}" : string.Empty);

                                        // add if there is an A2Q2 and ATeamHQ attendee
                                        if(sbParticipant.Length != 0)
                                        {
                                            sb.Append($"<b>{item.title}</b><br/>{sbParticipant}<br/><br/>");
                                        }
                                        

                                    }
                                    //sb.Append("<br/><br/>");
                                    sb.Append(bell);

                                    
                                    SkypeBot skypeBot = new SkypeBot(_config);
                                    Skype skypeParam = new Skype();
                                    //skypeParam.Address = _config["Skype:ErrorGc"];
                                    skypeParam.Address = itemGcAll.SkypeObj.conversation.id;
                                    skypeParam.Message = $"{sb}";
                                    //message = $"{bell}<br/><br/>Call | Test 10:00 AM - 11:00 AM<br/>Levin jay, John Doe<br/><br/>{bell}";
                                    skypeBot.SendSkypeNotif(skypeParam);
                                   
                                    

                                }
                            }

                            //#endregion

                            #region Send to Skype Keyword Base
                            //For all GC
                            var listGcKeyword = _meetingContext.SkypeAddress.Where(x =>
                                    x.IsAllGC.Equals(false) &&
                                    x.IsEnabled.Equals(true) &&
                                    x.IsBizDev.Equals(false) &&
                                    x.Is3PMNotification.Equals(true)
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
                                        StringBuilder sb = new StringBuilder();
                                        sb.Append(bell);
                                        sb.Append("<br/><br/>");
                                        bool isFound = false;

                                        foreach (var item in listMeetingAttendee)
                                        {
                                            var checkIfContains = itemAddress.ListKeyword.Where(x => item.title.ToLower().Contains(x.Keyword.ToLower())).FirstOrDefault();
                                            if(checkIfContains != null)
                                            {
                                                //string emailCc = _config.GetSection("Email").GetSection("SODSoxRox").GetSection("EmailCc").Value;
                                                string teamMember = string.Empty;
                                                string otherParticipant = string.Empty;
                                                StringBuilder sbParticipant = new StringBuilder();

                                                if (item.team_member != null && item.team_member.ToString() != string.Empty)
                                                {
                                                    //split and check if it contains in listEmployee
                                                    string[] member = item.team_member.Split(';');
                                                    if (member.Any())
                                                    {
                                                        foreach (var itemMember in member)
                                                        {
                                                            if (listEmployee.Where(mem => mem.f_name.ToLower().Equals(itemMember.ToLower())).FirstOrDefault() != null)
                                                                sbParticipant.Append($"{itemMember}, ");
                                                        }
                                                    }
                                                }

                                                if (item.other_participant != null && item.other_participant.ToString() != string.Empty)
                                                {
                                                    //split and check if it contains in listEmployee
                                                    string[] member = item.other_participant.Split(',');
                                                    if (member.Any())
                                                    {
                                                        foreach (var itemMember in member)
                                                        {
                                                            if (listEmployee.Where(mem => mem.f_email.ToLower().Equals(itemMember.ToLower())).FirstOrDefault() != null)
                                                                sbParticipant.Append($"{itemMember}, ");
                                                        }
                                                    }
                                                }

                                                if (sbParticipant.Length > 2)
                                                {
                                                    sbParticipant.Length -= 2; //Remove last 2 element
                                                }

                                                if (isLoggingEnabled)
                                                    FileLog.Write($"{DateTime.Now} [Sending 3PM Notification] - [All] {item.title}", "CronMeeting3PMNotif");

                                                //string participant = teamMember + (teamMember != string.Empty ? $", {otherParticipant}" : string.Empty);

                                                // add if there is an A2Q2 and ATeamHQ attendee
                                                if (sbParticipant.Length != 0)
                                                {
                                                    sb.Append($"<b>{item.title}</b><br/>{sbParticipant}<br/><br/>");
                                                    isFound = true;
                                                }
                                            }
                                        }

                                        //sb.Append("<br/><br/>");
                                        sb.Append(bell);

                                        if(isFound)
                                        {
                                            SkypeBot skypeBot = new SkypeBot(_config);
                                            Skype skypeParam = new Skype();
                                            //skypeParam.Address = _config["Skype:ErrorGc"];
                                            skypeParam.Address = itemAddress.SkypeObj.conversation.id;
                                            skypeParam.Message = $"{sb}";
                                            //message = $"{bell}<br/><br/>Call | Test 10:00 AM - 11:00 AM<br/>Levin jay, John Doe<br/><br/>{bell}";
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
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write($"Error Meeting3PMNotificationTask {ex}", "ErrorMeeting3PMNotificationTask");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "Meeting3PMNotificationTask");
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
