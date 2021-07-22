using A2B_App.Server.Services;
using A2B_App.Shared.Sox;
using A2B_App.Shared.Utilities;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;

namespace A2B_App.Server.Controllers
{
    [Authorize]
    [Route("api/[controller]")]
    [ApiController]
    public class ConfigController : ControllerBase
    {
        private readonly IConfiguration _config;

        public ConfigController(IConfiguration config)
        {
            _config = config;
        }

        [AllowAnonymous]
        [HttpGet("externalRef")]
        public IActionResult GetExternalReference()
        {

            try
            {
                List<ExternalLink> listExternalLink = new List<ExternalLink>();

                ExternalLink exLink = new ExternalLink();
                exLink.Name = "StaffReport";
                exLink.Link = _config.GetSection("ExternalLink").GetSection("StaffReport").Value;
                listExternalLink.Add(exLink);


                return Ok(listExternalLink);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorGetExternalReference");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetExternalReference");
                return BadRequest();
            }



        }



    }
}
