using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Threading.Tasks;
using A2B_App.Server.Data;
using A2B_App.Server.Services;
using A2B_App.Shared.Email;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

namespace A2B_App.Server.Controllers
{
    [Authorize]
    [Route("api/[controller]")]
    [ApiController]
    public class UtilityController : ControllerBase
    {
        private readonly IConfiguration _config;
        private readonly ILogger<UtilityController> _logger;
        private readonly IWebHostEnvironment _environment;
        private readonly SoxContext _soxContext;

        public UtilityController(IConfiguration config,
            ILogger<UtilityController> logger,
            IWebHostEnvironment environment,
            SoxContext soxContext)
        {
            _config = config;
            _logger = logger;
            _environment = environment;
            _soxContext = soxContext;
        }

        /*        [AllowAnonymous]
                 [HttpPost("email")]
               public  IActionResult SendEmail([FromBody] EmailParam emailParam)
                {
                    try
                    {

                        EmailService emailService = new EmailService(_config);
                        if (emailService.Send(emailParam))
                            return Ok();
                        else
                            return BadRequest();
                    }
                    catch (Exception ex)
                    {
                        if(_environment.IsDevelopment())
                            return BadRequest(ex);
                        else
                            return BadRequest();

                    }

                }
        */

    }
}
