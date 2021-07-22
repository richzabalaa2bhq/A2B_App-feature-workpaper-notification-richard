using A2B_App.Shared.Email;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Threading.Tasks;

namespace A2B_App.Server.Services
{

    public class EmailService
    {
        //private PodioField _podioField;
        private readonly IConfiguration _config;
        public EmailService(IConfiguration config)
        {
            _config = config;
        }

        public bool Send(EmailParam emailParam)
        {
            try
            {
                MailMessage objeto_mail = new MailMessage();
                SmtpClient client = new SmtpClient();
                client.Port = 587;
                client.Host = _config.GetSection("Email").GetSection("Smtp").GetSection("Host").Value; //data["Email"]["Host"];
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.UseDefaultCredentials = false;
                client.EnableSsl = true;

                string MAILER_USER = _config.GetSection("Email").GetSection("Smtp").GetSection("Username").Value; //data["Email"]["Username"];
                string MAILER_PASSWORD = _config.GetSection("Email").GetSection("Smtp").GetSection("Password").Value; //data["Email"]["Password"];
                string MAILER_EMAIL = _config.GetSection("Email").GetSection("Smtp").GetSection("Email").Value; //data["Email"]["Email"];
                string MAILER_NAME = _config.GetSection("Email").GetSection("Smtp").GetSection("EmailName").Value; //data["Email"]["EmailName"];

                client.Credentials = new System.Net.NetworkCredential(MAILER_USER, MAILER_PASSWORD);
                objeto_mail.From = new MailAddress(MAILER_EMAIL, MAILER_NAME);

                bool isValid = false;
                if (emailParam.ListEmailTo != null && emailParam.ListEmailTo.Any())
                {
                    isValid = true;
                    foreach (var item in emailParam.ListEmailTo)
                    {
                        objeto_mail.To.Add(new MailAddress(item));
                    }
                }

                if (emailParam.ListEmailCc != null && emailParam.ListEmailCc.Any())
                {
                    isValid = true;
                    foreach (var item in emailParam.ListEmailCc)
                    {
                        objeto_mail.To.Add(new MailAddress(item));
                    }
                }


                objeto_mail.ReplyToList.Add(new MailAddress("help@a2bhq.com", "Help | A2BHQ"));
                objeto_mail.IsBodyHtml = true;
                objeto_mail.Subject = $"{emailParam.Subject}";
                objeto_mail.Body = $"{emailParam.Message}";

                //objeto_mail.Attachments.Add(new Attachment(strOutputA2Q2));
                //objeto_mail.Attachments.Add(new Attachment(strOutputA2BHQ));
                if(isValid)
                {
                    client.Send(objeto_mail);
                    return true;
                }                 
                else
                    return false;
            }
            catch (Exception ex)
            {
                Shared.Sox.FileLog.Write(ex.ToString(), "ErrorEmailSend");
                return false;
            }
        }

        public Task<string> EmailNotificationWorkpaperTemplate(EmailNotificationBody emailNotif)
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
                                                                        $"<p>Hello {emailNotif.UserName},</p> " +
                                                                        $"<p> {emailNotif.Message} " +
                                                                        $"<a style='color:#1a73e8;font-family:Roboto,arial;font-size:14px;text-decoration:none' href='{emailNotif.Url}' target='_blank'>{emailNotif.UrlName}</a></p> " +
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
