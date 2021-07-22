using A2B_App.Server.Data;
//using A2B_App.Server.Log;
using A2B_App.Server.Services;
using A2B_App.Shared.Sms;
using A2B_App.Shared.Sox;
using A2B_App.Shared.Verification;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Linq;
using System.Threading.Tasks;

namespace A2B_App.Server.Controllers
{
    [Authorize]
    [ApiController]
    [Route("api/[controller]")]
    public class VerificationController : ControllerBase
    {
        private readonly IConfiguration _config;
        private readonly ILogger<VerificationController> _logger;
        private readonly IWebHostEnvironment _environment;
        private readonly VerificationContext _verificationContext;
        private readonly SmsContext _smsContext;

        public VerificationController(IConfiguration config,
            ILogger<VerificationController> logger,
            IWebHostEnvironment environment,
            VerificationContext verificationContext,
            SmsContext smsContext)
        {
            _config = config;
            _logger = logger;
            _environment = environment;
            _verificationContext = verificationContext;
            _smsContext = smsContext;
        }

        [HttpPost("create")]
        public async Task<IActionResult> PostCreateVerificationAsync([FromBody] RequestVerification requestVerification)
        {

            Random rand = new Random();
            string RefId = rand.Next(0, 999999).ToString("D6");
            string VerificationNum = rand.Next(0, 999999).ToString("D6");
            bool isCreated = false;
            bool isSend = false;
            using (var context = _verificationContext.Database.BeginTransaction())
            {
                try
                {

                    int minuteExpiry = int.Parse(_config.GetSection("Verification").GetSection("MinutesExpiry").Value);
                    Verification verify = new Verification();
                    verify.RefId = RefId;
                    verify.VerificationNum = VerificationNum;
                    verify.RequestedFromUserId = requestVerification.RequestedFromUserId;
                    verify.AppUse = requestVerification.AppUse;
                    verify.DateCreated = DateTime.Now.ToUniversalTime();
                    verify.DateUpdated = DateTime.Now.ToUniversalTime();
                    verify.ExpiryDate = DateTime.Now.AddMinutes(minuteExpiry).ToUniversalTime();
                    _verificationContext.Add(verify);

                    await _verificationContext.SaveChangesAsync();
                    context.Commit();
                    isCreated = true;
                }
                catch (Exception ex)
                {
                    //Console.WriteLine(ex.ToString());
                    //ErrorLog.Write(ex);
                    await context.RollbackAsync();
                    FileLog.Write($"Error CreatePostCreateVerificationAsync {ex}", "ErrorCreatePostCreateVerificationAsync");
                    AdminService adminService = new AdminService(_config);
                    adminService.SendAlert(true, true, ex.ToString(), "CreatePostCreateVerificationAsync");
                }
            }

            if (isCreated)
            {
                using (var context = _smsContext.Database.BeginTransaction())
                {
                    try
                    {
                        var checkEmployeeReference = _smsContext.EmployeeSmsReference.FirstOrDefault(x => x.EmployeeId == requestVerification.RequestedFromUserId);
                        if (checkEmployeeReference != null)
                        {

                            var checkNumber = _smsContext.Subscribe.Where(x => x.Id == checkEmployeeReference.SubscribeId && x.Status == "Subscribe").FirstOrDefault();
                            if (checkNumber != null)
                            {
                                SmsServices SmsService = new SmsServices(_smsContext, _config);
                                SmsSend smsSend = new SmsSend
                                {
                                    Address = checkNumber.SubScriberNum,
                                    Message = $"Your One-Time PIN is {VerificationNum}. Do not share this with anyone. Ref No. {RefId}"
                                };

                                if (SmsService.SendGlobeSmsAsync(smsSend))
                                {
                                    isSend = true;
                                }
                            }

                        }

                    }
                    catch (Exception ex)
                    {
                        FileLog.Write($"Error CreatePostCreateVerificationAsync {ex}", "ErrorCreatePostCreateVerificationAsync");
                    }
                }
            }



            if (isSend)
            {
                return Ok(new { RefNum = RefId });
            }
            else
            {
                return BadRequest("Bad request...");
            }


        }


        [HttpPost("verify")]
        public async Task<IActionResult> PostVerifyVerificationAsync([FromBody] RequestVerification requestVerification)
        {
            bool isVerified = false;
            using (var context = _verificationContext.Database.BeginTransaction())
            {
                try
                {
                    DateTime dtNow = DateTime.Now.ToUniversalTime();
                    var checkVerification = _verificationContext.Verification
                        .Where(x =>
                            x.RequestedFromUserId.Equals(requestVerification.RequestedFromUserId) &&
                            x.RefId.Equals(requestVerification.RefId) &&
                            x.VerificationNum.Equals(requestVerification.VerificationNum) &&
                            x.AppUse.Equals(requestVerification.AppUse) &&
                            x.Status.Equals(null) &&
                            x.ExpiryDate >= dtNow
                            )
                        .FirstOrDefault();

                    //x.ExpiryDate <= DateTime.Now.ToUniversalTime()
                    if (checkVerification != null)
                    {
                        isVerified = true;
                        checkVerification.Status = "Verified";
                        checkVerification.AppUse = requestVerification.AppUse;
                        checkVerification.Attempt++;
                        _verificationContext.Update(checkVerification);
                        await _verificationContext.SaveChangesAsync();
                        context.Commit();
                    }
                }
                catch (Exception ex)
                {
                    //Console.WriteLine(ex.ToString());
                    await context.RollbackAsync();
                    FileLog.Write($"Error PostVerifyVerificationAsync {ex}", "ErrorPostVerifyVerificationAsync");
                    AdminService adminService = new AdminService(_config);
                    adminService.SendAlert(true, true, ex.ToString(), "PostVerifyVerificationAsync");
                }
            }

            if (isVerified)
            {
                return Ok(new { status = "verified" });
            }
            else
            {
                return BadRequest("Bad request...");
            }


        }

    }
}
