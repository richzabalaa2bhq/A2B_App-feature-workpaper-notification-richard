using A2B_App.Shared.Email;
using A2B_App.Shared.Skype;
using A2B_App.Shared.Sox;
using Ionic.Zip;
using Microsoft.Bot.Schema;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace A2B_App.Server.Services
{
    public class AdminService
    {
        private readonly IConfiguration _config;
        public AdminService(IConfiguration config)
        {
            _config = config;
        }

        /// <summary>
        /// Send alert to skype and email
        /// </summary>
        /// <param name="isSkype">(bool) send to skype</param>
        /// <param name="isEmail">(bool) send to email</param>
        /// <param name="message">(string) message</param>
        /// <param name="method">(string) method that trigger the error</param>
        public async void SendAlert(bool isSkype, bool isEmail, string message, string method)
        {
            bool.TryParse(_config["Logging:SendAlert"], out bool isAlertEnabled);
            if (isAlertEnabled)
            {
                //Debug.WriteLine("Alert Triggered");
                if (isSkype)
                {
                    SkypeBot skypeBot = new SkypeBot(_config);
                    Skype skype = new Skype();
                    skype.Address = _config.GetSection("Skype").GetSection("ErrorGC").Value.ToString(); //Live
                    skype.Message = $"<b>A2B App Error</b> <br /> <br />Method: {method}<br />{message}";

                    //Microsoft.Bot.Schema.Mention mention = new Microsoft.Bot.Schema.Mention();
                    //mention.Mentioned.Id = "a2b.earocha";
                    //mention.Mentioned.Name = "Mark";
                    //skype.Mention = null;
                    await skypeBot.SendSkypeNotif(skype);
                }

                if (isEmail)
                {
                    EmailService emailService = new EmailService(_config);
                    EmailParam emailParam = new EmailParam();
                    emailParam.Subject = "A2B App Error";
                    emailParam.Message = $"<b>A2B App Error</b> <br /> <br />Method: {method}<br />{message}";
                    emailParam.ListEmailTo = new List<string>();
                    emailParam.ListEmailCc = new List<string>();
                    emailParam.ListEmailTo.Add(_config.GetSection("Email").GetSection("A2BHelpDesk").GetSection("EmailTo").Value);
                    emailParam.ListEmailCc.Add(_config.GetSection("Email").GetSection("A2BHelpDesk").GetSection("EmailCc").Value);
                    emailService.Send(emailParam);

                }
            }            
        }

        public async Task<bool> SendSkypeTest(string message, string address)
        {
            try
            {
                SkypeBot skypeBot = new SkypeBot(_config);
                Skype skype = new Skype();
                skype.Address = address; //Live
                skype.Message = $"{message}";
                var isSuccess = await skypeBot.SendSkypeNotif(skype);
                return isSuccess;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                return false;
                //throw;
            }
            
        }

        public async Task<bool> SendSkypeImage(string fileName, string address)
        {
            try
            {
                SkypeBot skypeBot = new SkypeBot(_config);
                Skype skype = new Skype();
                skype.Address = address; //Live
                
                var isSuccess = await skypeBot.SendSkypeImageNotif(skype, fileName);
                return isSuccess;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                return false;
                //throw;
            }

        }

        public void SendEmail(string subject, string body, string emailTo, string ccTo)
        {
            EmailService emailService = new EmailService(_config);
            EmailParam emailParam = new EmailParam();
            emailParam.Subject = subject;
            emailParam.Message = body;
            emailParam.ListEmailTo = new List<string>();
            emailParam.ListEmailCc = new List<string>();
            emailParam.ListEmailTo.Add(emailTo);
            var emailCcList = ccTo.Split(',').ToList();
            if(emailCcList.Any())
            {
                foreach (var item in emailCcList)
                {
                    emailParam.ListEmailCc.Add(item);
                }
            }
            emailService.Send(emailParam);
        }

        public Task<string> CompressItem(string path, string reportname, List<string> listFileName)
        {
            string zipFileName = string.Empty;
            try
            {
                List<string> fileFullPath = new List<string>();
                using (ZipFile zip = new ZipFile())
                {
                    if(listFileName.Any())
                    {
                        foreach (var item in listFileName)
                        {
                            fileFullPath.Add(Path.Combine(path, item));
                        }
                        zip.AddFiles(fileFullPath, false, "");
                        string dtNow = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                        string zipFullPath = Path.Combine(path, $"{dtNow}_{reportname}.zip");
                        zipFileName = $"{dtNow}_{reportname}.zip";
                        zip.Save(zipFullPath);
                    }
                }

            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorCompressItem");
                //SendAlert(true, true, ex.ToString(), "CompressItem");
            }


            return Task.FromResult(zipFileName);
        }

        public Task<string> SodSoxRoxEmailBody(string url, string filename)
        {
            string body = $" <tbody> " +
                "<tr> " +
                    "<td style='border-collapse:collapse'> " +
                        "<table align='center' width='600' cellspacing='0' cellpadding='0' border='0' " +
                            "style='border-collapse:collapse;font-family:Arial,sans-serif;font-weight:normal;margin:0 auto;max-width:600px;min-width:600px;width:600px'> " +
                            "<tbody> " +
                                "<tr> " +
                                    "<td> " +
                                        "<table width='600' align='center' cellspacing='0' cellpadding='0' border='0' " +
                                            "style='border-collapse:collapse;font-family:Arial,sans-serif;font-weight:normal;margin:0 auto;width:600px'> " +
                                            "<tbody> " +
                                                "<tr> " +
                                                    "<td style='border-collapse:collapse;padding:0 5px'> " +
                                                        "<table width='auto' align='left' cellspacing='0' cellpadding='0' border='0' " +
                                                            "style='border-collapse:collapse;font-family:roboto,arial;font-weight:500;width:auto'> " +
                                                            "<tbody> " +
                                                                "<tr> " +
                                                                    "<td " +
                                                                        "style='border-collapse:collapse;padding-bottom: 30px;padding-top:32px'> " +
                                                                        "<img src='https://a2bhq.com/wp-content/uploads/2017/09/newlogoa2b-300x185.png' " +
                                                                            "width='100' /> " +
                                                                    "</td> " +
                                                                "</tr> " +
                                                            "</tbody> " +
                                                        "</table> " +
                                                        "<table role='presentation' width='auto' align='right' cellspacing='0' " +
                                                            "cellpadding='0' border='0' " +
                                                            "style='border-collapse:collapse;font-family:roboto,arial;font-weight:500;width:auto'> " +
                                                            "<tbody> " +
                                                                "<tr> " +
                                                                    "<td " +
                                                                        "style='border-collapse:collapse;padding-right:0;padding-top:41px'> " +
                                                                        "<table cellspacing='0' cellpadding='0' border='0' " +
                                                                            "style='border-collapse:collapse;font-family:roboto,arial;font-weight:500;width:auto' " +
                                                                            "width='auto'> " +
                                                                            "<tbody> " +
                                                                                "<tr> " +
                                                                                "</tr> " +
                                                                            "</tbody> " +
                                                                        "</table> " +
                                                                    "</td> " +
                                                                "</tr> " +
                                                            "</tbody> " +
                                                        "</table> " +
                                                    "</td> " +
                                                "</tr> " +
                                            "</tbody> " +
                                        "</table> " +
                                    "</td> " +
                                "</tr> " +
                                "<tr> " +
                                    "<td style='border-collapse:collapse'> " +
                                        "<table width='598' align='center' cellspacing='0' cellpadding='0' border='0' " +
                                            "style='background:#ffffff;border-collapse:collapse;border-left:1px solid #e0e0e0;border-right:1px solid #e0e0e0;border-top:1px solid #e0e0e0;font-family:Arial,sans-serif;font-weight:normal;width:598px;border-radius:5px 5px 0 0'> " +
                                            "<tbody> " +
                                                "<tr> " +
                                                    "<td style='padding:30px 40px 0 40px;padding-top:35px;width:518px' width='518'> " +
                                                        "<table align='center' " +
                                                            "style='border-collapse:collapse;font-family:Arial,sans-serif;font-weight:normal;margin:auto;max-width:598px;width:518px' " +
                                                            "cellspacing='0' cellpadding='0' border='0' width='518'> " +
                                                            "<tbody> " +
                                                                "<tr> " +
                                                                    "<td " +
                                                                        "style='border-collapse:collapse;color:#455a64;font-family:Roboto,arial;font-size:14px;line-height:24px;padding-bottom:17px'> " +
                                                                        "<p>Hello A2B User,</p> " +
                                                                        "<p>Your file is available for download "+
                                                                        "<a style='color:#1a73e8;font-family:Roboto,arial;font-size:14px;text-decoration:none' href='"+ url + "' target='_blank'>"+ filename +"</a></p> " +
                                                                        "<p>For comments and concern, please email " +
                                                                            "<a style='color:#1a73e8;font-family:Roboto,arial;font-size:14px;text-decoration:none' href='mailto:help@a2bhq.com'>help@a2bhq.com</a> " +
                                                                        "</p> " +
                                                                    "</td> " +
                                                                "</tr> " +
                                                            "</tbody> " +
                                                        "</table> " +
                                                    "</td> " +
                                                "</tr> " +
                                                "<tr> " +
                                                    "<td style='padding:30px 40px 20px 40px;width:518px' width='518'> " +
                                                        "<table cellspacing='0' cellpadding='0' border='0' style='border-collapse:collapse;font-family:Arial,sans-serif;font-weight:normal;width:518px' width='518'> " +
                                                            "<tbody> " +
                                                                "<tr> " +
                                                                    "<td " +
                                                                        "style='border-collapse:collapse;color:#455a64;font-family:Roboto,arial;font-size:14px;font-weight:regular;line-height:24px;padding-bottom:0'> " +
                                                                        "Sincerely, " +
                                                                    "</td> " +
                                                                "</tr> " +
                                                                "<tr> " +
                                                                    "<td " +
                                                                        "style='border-collapse:collapse;color:#455a64;font-family:Roboto,arial;font-size:16px;line-height:24px;padding-bottom:12px;font-weight:500'> " +
                                                                        "A2B App " +
                                                                    "</td> " +
                                                                "</tr> " +
                                                            "</tbody> " +
                                                        "</table> " +
                                                    "</td> " +
                                                "</tr> " +
                                            "</tbody> " +
                                        "</table> " +
                                    "</td> " +
                                "</tr> " +
                                "<tr> " +
                                    "<td " +
                                        "style='color:#455a64;font-family:Roboto,arial;font-size:14px;line-height:24px;padding-bottom:17px'> " +
                                        "<table cellspacing='0' cellpadding='0' border='0' " +
                                            "style='background:#ffffff;border-bottom:1px solid #e0e0e0;border-collapse:collapse;border-left:1px solid #e0e0e0;border-right:1px solid #e0e0e0;border-top:2px solid #eef1f2;font-family:Arial,sans-serif;font-weight:normal;margin:0 auto;max-width:600px;min-width:600px;width:598px;border-radius:0 0 5px 5px'  " +
                                            "width='598' align='center'> " +
                                            "<tbody> " +
                                                "<tr> " +
                                                    "<td style='border-collapse:collapse;padding:25px 40px 27px 40px'> " +
                                                        "<table cellspacing='0' cellpadding='0' border='0' width='172' align='left' " +
                                                            "style='border-collapse:collapse;font-family:Arial,sans-serif;font-weight:normal'> " +
                                                        "</table> " +
                                                        "<table align='left' width='226' cellspacing='0' cellpadding='0' border='0' " +
                                                            "style='border-collapse:collapse;font-family:Arial,sans-serif;font-weight:normal'> " +
                                                        "</table> " +
                                                    "</td> " +
                                                "</tr> " +
                                            "</tbody> " +
                                        "</table> " +
                                    "</td> " +
                                "</tr> " +
                            "</tbody> " +
                        "</table> " +
                    "</td> " +
                "</tr> " +
                "<tr> " +
                    "<td> " +
                    "</td> " +
                "</tr> " +
            "</tbody> ";



            return Task.FromResult(body);
        }

        public Task<string> EmailNotificationTemplate(EmailNotificationBody emailNotif)
        {
            string body = $" <tbody> " +
                "<tr> " +
                    "<td style='border-collapse:collapse'> " +
                        "<table align='center' width='600' cellspacing='0' cellpadding='0' border='0' " +
                            "style='border-collapse:collapse;font-family:Arial,sans-serif;font-weight:normal;margin:0 auto;max-width:600px;min-width:600px;width:600px'> " +
                            "<tbody> " +
                                "<tr> " +
                                    "<td> " +
                                        "<table width='600' align='center' cellspacing='0' cellpadding='0' border='0' " +
                                            "style='border-collapse:collapse;font-family:Arial,sans-serif;font-weight:normal;margin:0 auto;width:600px'> " +
                                            "<tbody> " +
                                                "<tr> " +
                                                    "<td style='border-collapse:collapse;padding:0 5px'> " +
                                                        "<table width='auto' align='left' cellspacing='0' cellpadding='0' border='0' " +
                                                            "style='border-collapse:collapse;font-family:roboto,arial;font-weight:500;width:auto'> " +
                                                            "<tbody> " +
                                                                "<tr> " +
                                                                    "<td " +
                                                                        "style='border-collapse:collapse;padding-bottom: 30px;padding-top:32px'> " +
                                                                        "<img src='https://a2bhq.com/wp-content/uploads/2017/09/newlogoa2b-300x185.png' " +
                                                                            "width='100' /> " +
                                                                    "</td> " +
                                                                "</tr> " +
                                                            "</tbody> " +
                                                        "</table> " +
                                                        "<table role='presentation' width='auto' align='right' cellspacing='0' " +
                                                            "cellpadding='0' border='0' " +
                                                            "style='border-collapse:collapse;font-family:roboto,arial;font-weight:500;width:auto'> " +
                                                            "<tbody> " +
                                                                "<tr> " +
                                                                    "<td " +
                                                                        "style='border-collapse:collapse;padding-right:0;padding-top:41px'> " +
                                                                        "<table cellspacing='0' cellpadding='0' border='0' " +
                                                                            "style='border-collapse:collapse;font-family:roboto,arial;font-weight:500;width:auto' " +
                                                                            "width='auto'> " +
                                                                            "<tbody> " +
                                                                                "<tr> " +
                                                                                "</tr> " +
                                                                            "</tbody> " +
                                                                        "</table> " +
                                                                    "</td> " +
                                                                "</tr> " +
                                                            "</tbody> " +
                                                        "</table> " +
                                                    "</td> " +
                                                "</tr> " +
                                            "</tbody> " +
                                        "</table> " +
                                    "</td> " +
                                "</tr> " +
                                "<tr> " +
                                    "<td style='border-collapse:collapse'> " +
                                        "<table width='598' align='center' cellspacing='0' cellpadding='0' border='0' " +
                                            "style='background:#ffffff;border-collapse:collapse;border-left:1px solid #e0e0e0;border-right:1px solid #e0e0e0;border-top:1px solid #e0e0e0;font-family:Arial,sans-serif;font-weight:normal;width:598px;border-radius:5px 5px 0 0'> " +
                                            "<tbody> " +
                                                "<tr> " +
                                                    "<td style='padding:30px 40px 0 40px;padding-top:35px;width:518px' width='518'> " +
                                                        "<table align='center' " +
                                                            "style='border-collapse:collapse;font-family:Arial,sans-serif;font-weight:normal;margin:auto;max-width:598px;width:518px' " +
                                                            "cellspacing='0' cellpadding='0' border='0' width='518'> " +
                                                            "<tbody> " +
                                                                "<tr> " +
                                                                    "<td " +
                                                                        "style='border-collapse:collapse;color:#455a64;font-family:Roboto,arial;font-size:14px;line-height:24px;padding-bottom:17px'> " +
                                                                        "<p>Hello ,</p> " +
                                                                        "<p>New workpaper has been created by " +
                                                                        "<a style='color:#1a73e8;font-family:Roboto,arial;font-size:14px;text-decoration:none' href='" + emailNotif.Url + "' target='_blank'>" + emailNotif.UrlName + "</a></p> " +
                                                                        "<p>For comments and concern, please email " +
                                                                            "<a style='color:#1a73e8;font-family:Roboto,arial;font-size:14px;text-decoration:none' href='mailto:help@a2bhq.com'>help@a2bhq.com</a> " +
                                                                        "</p> " +
                                                                    "</td> " +
                                                                "</tr> " +
                                                            "</tbody> " +
                                                        "</table> " +
                                                    "</td> " +
                                                "</tr> " +
                                                "<tr> " +
                                                    "<td style='padding:30px 40px 20px 40px;width:518px' width='518'> " +
                                                        "<table cellspacing='0' cellpadding='0' border='0' style='border-collapse:collapse;font-family:Arial,sans-serif;font-weight:normal;width:518px' width='518'> " +
                                                            "<tbody> " +
                                                                "<tr> " +
                                                                    "<td " +
                                                                        "style='border-collapse:collapse;color:#455a64;font-family:Roboto,arial;font-size:14px;font-weight:regular;line-height:24px;padding-bottom:0'> " +
                                                                        "Sincerely, " +
                                                                    "</td> " +
                                                                "</tr> " +
                                                                "<tr> " +
                                                                    "<td " +
                                                                        "style='border-collapse:collapse;color:#455a64;font-family:Roboto,arial;font-size:16px;line-height:24px;padding-bottom:12px;font-weight:500'> " +
                                                                        "A2B App " +
                                                                    "</td> " +
                                                                "</tr> " +
                                                            "</tbody> " +
                                                        "</table> " +
                                                    "</td> " +
                                                "</tr> " +
                                            "</tbody> " +
                                        "</table> " +
                                    "</td> " +
                                "</tr> " +
                                "<tr> " +
                                    "<td " +
                                        "style='color:#455a64;font-family:Roboto,arial;font-size:14px;line-height:24px;padding-bottom:17px'> " +
                                        "<table cellspacing='0' cellpadding='0' border='0' " +
                                            "style='background:#ffffff;border-bottom:1px solid #e0e0e0;border-collapse:collapse;border-left:1px solid #e0e0e0;border-right:1px solid #e0e0e0;border-top:2px solid #eef1f2;font-family:Arial,sans-serif;font-weight:normal;margin:0 auto;max-width:600px;min-width:600px;width:598px;border-radius:0 0 5px 5px'  " +
                                            "width='598' align='center'> " +
                                            "<tbody> " +
                                                "<tr> " +
                                                    "<td style='border-collapse:collapse;padding:25px 40px 27px 40px'> " +
                                                        "<table cellspacing='0' cellpadding='0' border='0' width='172' align='left' " +
                                                            "style='border-collapse:collapse;font-family:Arial,sans-serif;font-weight:normal'> " +
                                                        "</table> " +
                                                        "<table align='left' width='226' cellspacing='0' cellpadding='0' border='0' " +
                                                            "style='border-collapse:collapse;font-family:Arial,sans-serif;font-weight:normal'> " +
                                                        "</table> " +
                                                    "</td> " +
                                                "</tr> " +
                                            "</tbody> " +
                                        "</table> " +
                                    "</td> " +
                                "</tr> " +
                            "</tbody> " +
                        "</table> " +
                    "</td> " +
                "</tr> " +
                "<tr> " +
                    "<td> " +
                    "</td> " +
                "</tr> " +
            "</tbody> ";



            return Task.FromResult(body);
        }

         

    }
}
