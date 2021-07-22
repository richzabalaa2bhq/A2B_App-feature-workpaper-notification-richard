using A2B_App.Server.Data;
//using A2B_App.Server.Log;
using A2B_App.Shared.Sms;
using A2B_App.Shared.Sox;
using Microsoft.Extensions.Configuration;
using RestSharp;
using System;
using System.Linq;
using System.Threading.Tasks;

namespace A2B_App.Server.Services
{
    public class SmsServices
    {
        private SmsContext _smsContext;
        private readonly IConfiguration _config;
        public SmsServices(SmsContext smsContext, IConfiguration config)
        {
            _smsContext = smsContext;
            _config = config;
        }

        public async Task<bool> SaveGlobeSmsAsync(GlobeSms globeSms)
        {

            bool isSuccess = false;
            try
            {
                using (var context = _smsContext.Database.BeginTransaction())
                {
                    Subscribe subscribe = new Subscribe();
                    //Subscribe 
                    if (globeSms.subscriber_number != null)
                    {
                        subscribe = _smsContext.Subscribe.FirstOrDefault(id => id.SubScriberNum == globeSms.subscriber_number);
                        if (subscribe == null)
                        {
                            subscribe = new Subscribe();
                            subscribe.SubScriberNum = globeSms.subscriber_number;
                            subscribe.Status = "Subscribe";
                            subscribe.Token = globeSms.access_token;
                            subscribe.DateCreated = DateTime.Now.ToUniversalTime();
                            subscribe.DateUpdated = DateTime.Now.ToUniversalTime();
                            _smsContext.Add(subscribe);
                        }
                        else
                        {
                            subscribe.Status = "Subscribe";
                            subscribe.Token = globeSms.access_token;
                            subscribe.DateUpdated = DateTime.Now.ToUniversalTime();
                            _smsContext.Update(subscribe);
                        }
                    }

                    //Unsubscribe
                    else if (globeSms.unsubscribed != null)
                    {
                        subscribe = _smsContext.Subscribe.FirstOrDefault(id => id.SubScriberNum == globeSms.unsubscribed.subscriber_number);
                        if (subscribe != null)
                        {
                            subscribe.Status = "Unsubscribe";
                            subscribe.DateUpdated = DateTime.Now.ToUniversalTime();
                            subscribe.Token = string.Empty;
                            _smsContext.Update(subscribe);
                        }
                    }


                    await _smsContext.SaveChangesAsync();
                    context.Commit();
                    isSuccess = true;
                }
            }
            catch (Exception ex)
            {
                FileLog.Write($"Error SaveGlobeSmsAsync {ex}", "ErrorSaveGlobeSmsAsync");


            }

            return isSuccess;
        }

        public bool SendGlobeSmsAsync(SmsSend smsSend)
        {
            bool isSuccess = false;
            try
            {
                var subscribe = _smsContext.Subscribe.FirstOrDefault(id => id.SubScriberNum == smsSend.Address);
                if (subscribe != null)
                {
                    var shortcode = _config.GetSection("GlobeSms").GetSection("ShortCode").Value;
                    var access_token = subscribe.Token;
                    var defaulturl = _config.GetSection("GlobeSms").GetSection("DefaultUrl").Value;
                    var address = smsSend.Address;
                    var clientCorrelator = _config.GetSection("GlobeSms").GetSection("ClientCorrelator").Value;
                    var message = smsSend.Message;
                    var url = defaulturl + shortcode + "/requests" + "?access_token=" + access_token;
                    var client = new RestClient(url);
                    var request = new RestRequest();
                    request.Method = Method.POST;
                    request.RequestFormat = DataFormat.Json;
                    request.Parameters.Clear();
                    request.AddJsonBody(new
                    {
                        outboundSMSMessageRequest = new
                        {
                            clientCorrelator = clientCorrelator,
                            senderAddress = shortcode,
                            outboundSMSTextMessage = new { message = message },
                            address = address
                        }
                    });
                    IRestResponse response = client.Execute(request);
                    var content = response.Content;
                    isSuccess = true;
                }

            }
            catch (Exception ex)
            {
                FileLog.Write($"Error SendGlobeSmsAsync {ex}", "ErrorSendGlobeSmsAsync");

            }

            return isSuccess;
        }

        public async Task<bool> AddEmployeeRefAsync(EmployeeSmsReference employeeSms)
        {
            bool isSuccess = false;
            using (var context = _smsContext.Database.BeginTransaction())
            {
                try
                {

                    var employeeRef = _smsContext.EmployeeSmsReference.FirstOrDefault(id => id.EmployeeId == employeeSms.EmployeeId);
                    if (employeeRef != null)
                    {
                        employeeRef.SubscribeId = employeeSms.SubscribeId;
                        _smsContext.Update(employeeRef);
                    }
                    else
                    {
                        _smsContext.Add(employeeSms);
                    }

                    await _smsContext.SaveChangesAsync();
                    context.Commit();
                    isSuccess = true;

                }
                catch (Exception ex)
                {
                    context.Rollback();
                    FileLog.Write($"Error AddEmployeeRefAsync {ex}", "ErrorAddEmployeeRefAsync");

                }

            }

            return isSuccess;
        }
    }
}
