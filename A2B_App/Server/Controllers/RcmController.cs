using A2B_App.Server.Data;
using A2B_App.Server.Services;
using A2B_App.Shared.Podio;
//using A2B_App.Server.Log;
using A2B_App.Shared.Sox;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using PodioAPI;
using PodioAPI.Models;
using PodioAPI.Utils.ItemFields;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace A2B_App.Server.Controllers
{
    [Authorize]
    [Route("api/[controller]")]
    [ApiController]
    public class RcmController : ControllerBase
    {

        private readonly IConfiguration _config;
        private readonly ILogger<RcmController> _logger;
        private readonly IWebHostEnvironment _environment;
        private readonly SoxContext _soxContext;

        public RcmController(IConfiguration config,
            ILogger<RcmController> logger,
            IWebHostEnvironment environment,
            SoxContext soxContext)
        {
            _config = config;
            _logger = logger;
            _environment = environment;
            _soxContext = soxContext;
        }

        [AllowAnonymous]
        [HttpGet("rcm/get/{itemId:int}")]
        public IActionResult GetRcmAsync(int itemId)
       {
            Rcm _rcm = new Rcm();
            bool status = false;
            try
            {
                using (var context = _soxContext.Database.BeginTransaction())
                {

                    bool isExists = false;
                    var rcmCheck = _soxContext.Rcm.FirstOrDefault(id => id.PodioItemId == itemId);
                    if (rcmCheck != null)
                    {
                        _rcm = rcmCheck;
                        isExists = true;
                    }

                    if (isExists)
                    {
                        status = true;
                    }

                }
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetRcmAsync {ex}", "ErrorGetRcmAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetRcmAsync");
            }

            if (status)
            {
                return Ok(_rcm);
            }
            else
            {
                return NoContent();
            }

        }

        [AllowAnonymous]
        [HttpGet("rcm/clients/")]
        public IActionResult GetListClientAsync()
       {
            List<string> _clients = new List<string>();
            try
            {
                using (var context = _soxContext.Database.BeginTransaction())
                {
                    var listClient = _soxContext.ClientSs.OrderBy(x => x.Name).Select(x => x.Name);

                    if (listClient != null)
                    {
                        _clients = listClient.ToList();
                    }
                }
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetListClientAsync {ex}", "ErrorGetListClientAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetListClientAsync");
            }

            if (_clients != null)
            {
                return Ok(_clients.ToArray());
            }
            else
            {
                return NoContent();
            }

        }


        //[AllowAnonymous]
        [HttpGet("rcm/controlname/")]
        public IActionResult GetListControlNameAsync()
       {



            List<string> _controlName = new List<string>();
            try
            {
                var controlName = _config.GetSection("ControlName").AsEnumerable().Where(x => x.Value != null).Select(x => x.Value);
                _controlName = controlName.ToList();
                _controlName.Sort();
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetListControlNameAsync {ex}", "ErrorGetListControlNameAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetListControlNameAsync");
            }

            if (_controlName != null)
            {
                return Ok(_controlName.ToArray());
            }
            else
            {
                return NoContent();
            }

        }

        [AllowAnonymous]
        [HttpPost("rcm/get/")]
        public IActionResult GetListRcmAsync([FromBody] RcmItemFilter rcmFilter)
        {

            List<Rcm> _rcm = new List<Rcm>();
            bool status = false;
            try
            {
                using (var context = _soxContext.Database.BeginTransaction())
                {

                    bool isExists = false;
                    IQueryable<Rcm> rcmCheck = null;
                    switch (rcmFilter.WorkpaperVersion)
                    {
                        case "sox":
                            if (rcmFilter.ClientName != "All" && rcmFilter.ControlName != "All" && rcmFilter.StartDate != null && rcmFilter.EndDate != null)
                            {

                                rcmCheck = _soxContext.Rcm
                                    .Where(x =>
                                        x.ClientName.Contains(rcmFilter.ClientName) &&
                                        x.ControlId.Contains(rcmFilter.ControlName) &&
                                        !x.ControlId.ToLower().StartsWith("it") &&
                                        x.CreatedOn >= rcmFilter.StartDate &&
                                        x.CreatedOn <= rcmFilter.EndDate &&
                                        x.Status.ToLower() != "closed" &&
                                        x.Status.ToLower() != "completed"
                                     )
                                    .OrderByDescending(x => x.CreatedOn)
                                    .AsNoTracking()
                                    .Skip(rcmFilter.Offset)
                                    .Take(rcmFilter.Limit);
                            }
                            else if (rcmFilter.ClientName != "All" && rcmFilter.ControlName != "All" && rcmFilter.StartDate == null && rcmFilter.EndDate == null)
                            {
                                rcmCheck = _soxContext.Rcm
                                    .Where(x =>
                                        x.ClientName.Contains(rcmFilter.ClientName) &&
                                        x.ControlId.Contains(rcmFilter.ControlName) &&
                                        !x.ControlId.ToLower().StartsWith("it") &&
                                        x.Status.ToLower() != "closed" &&
                                        x.Status.ToLower() != "completed"
                                     )
                                    .OrderByDescending(x => x.CreatedOn)
                                    .AsNoTracking()
                                    .Skip(rcmFilter.Offset)
                                    .Take(rcmFilter.Limit);
                            }
                            else if (rcmFilter.ClientName != "All" && rcmFilter.ControlName == "All" && rcmFilter.StartDate != null && rcmFilter.EndDate != null)
                            {
                                rcmCheck = _soxContext.Rcm
                                    .Where(x =>
                                        x.ClientName.Contains(rcmFilter.ClientName) &&
                                        !x.ControlId.ToLower().StartsWith("it") &&
                                        x.CreatedOn >= rcmFilter.StartDate &&
                                        x.CreatedOn <= rcmFilter.EndDate &&
                                        x.Status.ToLower() != "closed" &&
                                        x.Status.ToLower() != "completed"
                                     )
                                    .OrderByDescending(x => x.CreatedOn)
                                    .AsNoTracking()
                                    .Skip(rcmFilter.Offset)
                                    .Take(rcmFilter.Limit);
                            }
                            else if (rcmFilter.ClientName != "All" && rcmFilter.ControlName == "All" && rcmFilter.StartDate == null && rcmFilter.EndDate == null)
                            {
                                rcmCheck = _soxContext.Rcm
                                    .Where(x =>
                                        x.ClientName.Contains(rcmFilter.ClientName) &&
                                        !x.ControlId.ToLower().StartsWith("it") &&
                                        (x.Status != "Closed" || x.Status != "Completed")
                                     )
                                    .OrderByDescending(x => x.CreatedOn)
                                    .AsNoTracking()
                                    .Skip(rcmFilter.Offset)
                                    .Take(rcmFilter.Limit);
                            }
                            else if (rcmFilter.ClientName == "All" && rcmFilter.ControlName != "All" && rcmFilter.StartDate != null && rcmFilter.EndDate != null)
                            {
                                rcmCheck = _soxContext.Rcm
                                    .Where(x =>
                                        x.ControlId.Contains(rcmFilter.ControlName)
                                        && !x.ControlId.ToLower().StartsWith("it")
                                        && x.CreatedOn >= rcmFilter.StartDate
                                        && x.CreatedOn <= rcmFilter.EndDate
                                        && x.Status.ToLower() != "closed"
                                        && x.Status.ToLower() != "completed"
                                     )
                                    .OrderByDescending(x => x.CreatedOn)
                                    .AsNoTracking()
                                    .Skip(rcmFilter.Offset)
                                    .Take(rcmFilter.Limit);
                            }
                            else if (rcmFilter.ClientName == "All" && rcmFilter.ControlName != "All" && rcmFilter.StartDate == null && rcmFilter.EndDate == null)
                            {
                                rcmCheck = _soxContext.Rcm
                                    .Where(x =>
                                        x.ControlId.Contains(rcmFilter.ControlName)
                                        && !x.ControlId.ToLower().StartsWith("it")
                                        && x.Status.ToLower() != "closed" &&
                                        x.Status.ToLower() != "completed"
                                     )
                                    .OrderByDescending(x => x.CreatedOn)
                                    .AsNoTracking()
                                    .Skip(rcmFilter.Offset)
                                    .Take(rcmFilter.Limit);
                            }
                            else
                            {
                                rcmCheck = _soxContext.Rcm.Where(x =>
                                   !x.ControlId.ToLower().StartsWith("it") &&
                                    x.Status.ToLower() != "closed" &&
                                    x.Status.ToLower() != "completed")
                                    .OrderByDescending(x => x.CreatedOn)
                                    .AsNoTracking()
                                    .Skip(rcmFilter.Offset).Take(rcmFilter.Limit);
                            }

                            break;
                        case "itgc":

                            if (rcmFilter.ClientName != "All" && rcmFilter.ControlName != "All" && rcmFilter.StartDate != null && rcmFilter.EndDate != null)
                            {

                                rcmCheck = _soxContext.Rcm
                                    .Where(x =>
                                        x.ClientName.Contains(rcmFilter.ClientName) &&
                                        x.ControlId.Contains(rcmFilter.ControlName) &&
                                        x.ControlId.ToLower().StartsWith("it") &&
                                        x.CreatedOn >= rcmFilter.StartDate &&
                                        x.CreatedOn <= rcmFilter.EndDate &&
                                        x.Status.ToLower() != "closed" &&
                                        x.Status.ToLower() != "completed"
                                     )
                                    .OrderByDescending(x => x.CreatedOn)
                                    .AsNoTracking()
                                    .Skip(rcmFilter.Offset)
                                    .Take(rcmFilter.Limit);
                            }
                            else if (rcmFilter.ClientName != "All" && rcmFilter.ControlName != "All" && rcmFilter.StartDate == null && rcmFilter.EndDate == null)
                            {
                                rcmCheck = _soxContext.Rcm
                                    .Where(x =>
                                        x.ClientName.Contains(rcmFilter.ClientName) &&
                                        x.ControlId.Contains(rcmFilter.ControlName) &&
                                        x.ControlId.ToLower().StartsWith("it") &&
                                        x.Status.ToLower() != "closed" &&
                                        x.Status.ToLower() != "completed"
                                     )
                                    .OrderByDescending(x => x.CreatedOn)
                                    .AsNoTracking()
                                    .Skip(rcmFilter.Offset)
                                    .Take(rcmFilter.Limit);
                            }
                            else if (rcmFilter.ClientName != "All" && rcmFilter.ControlName == "All" && rcmFilter.StartDate != null && rcmFilter.EndDate != null)
                            {
                                rcmCheck = _soxContext.Rcm
                                    .Where(x =>
                                        x.ClientName.Contains(rcmFilter.ClientName) &&
                                        x.ControlId.ToLower().StartsWith("it") &&
                                        x.CreatedOn >= rcmFilter.StartDate &&
                                        x.CreatedOn <= rcmFilter.EndDate &&
                                        x.Status.ToLower() != "closed" &&
                                        x.Status.ToLower() != "completed"
                                     )
                                    .OrderByDescending(x => x.CreatedOn)
                                    .AsNoTracking()
                                    .Skip(rcmFilter.Offset)
                                    .Take(rcmFilter.Limit);
                            }
                            else if (rcmFilter.ClientName != "All" && rcmFilter.ControlName == "All" && rcmFilter.StartDate == null && rcmFilter.EndDate == null)
                            {
                                rcmCheck = _soxContext.Rcm
                                    .Where(x =>
                                        x.ClientName.Contains(rcmFilter.ClientName) &&
                                        x.ControlId.ToLower().StartsWith("it") &&
                                        (x.Status != "Closed" || x.Status != "Completed")
                                     )
                                    .OrderByDescending(x => x.CreatedOn)
                                    .AsNoTracking()
                                    .Skip(rcmFilter.Offset)
                                    .Take(rcmFilter.Limit);
                            }
                            else if (rcmFilter.ClientName == "All" && rcmFilter.ControlName != "All" && rcmFilter.StartDate != null && rcmFilter.EndDate != null)
                            {
                                rcmCheck = _soxContext.Rcm
                                    .Where(x =>
                                        x.ControlId.Contains(rcmFilter.ControlName)
                                        && x.ControlId.ToLower().StartsWith("it")
                                        && x.CreatedOn >= rcmFilter.StartDate
                                        && x.CreatedOn <= rcmFilter.EndDate
                                        && x.Status.ToLower() != "closed"
                                        && x.Status.ToLower() != "completed"
                                     )
                                    .OrderByDescending(x => x.CreatedOn)
                                    .AsNoTracking()
                                    .Skip(rcmFilter.Offset)
                                    .Take(rcmFilter.Limit);
                            }
                            else if (rcmFilter.ClientName == "All" && rcmFilter.ControlName != "All" && rcmFilter.StartDate == null && rcmFilter.EndDate == null)
                            {
                                rcmCheck = _soxContext.Rcm
                                    .Where(x =>
                                        x.ControlId.Contains(rcmFilter.ControlName)
                                        && x.ControlId.ToLower().StartsWith("it")
                                        && x.Status.ToLower() != "closed" &&
                                        x.Status.ToLower() != "completed"
                                     )
                                    .OrderByDescending(x => x.CreatedOn)
                                    .AsNoTracking()
                                    .Skip(rcmFilter.Offset)
                                    .Take(rcmFilter.Limit);
                            }
                            else
                            {
                                rcmCheck = _soxContext.Rcm.Where(x =>
                                   x.ControlId.ToLower().StartsWith("it") &&
                                    x.Status.ToLower() != "closed" &&
                                    x.Status.ToLower() != "completed")
                                    .OrderByDescending(x => x.CreatedOn)
                                    .AsNoTracking()
                                    .Skip(rcmFilter.Offset).Take(rcmFilter.Limit);
                            }

                            break;
                        default:
                            break;
                    }
                    

                    if (rcmCheck != null)
                    {               
                        _rcm = rcmCheck.ToList();
                        isExists = true;
                    }

                    if (isExists)
                    {
                        status = true;
                    }

                }
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetListRcmAsync {ex}", "ErrorGetListRcmAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetListRcmAsync");
            }

            if (status)
            {
                if (_rcm.Any())
                {
                    foreach (var item in _rcm)
                    {
                        var workpaperStatus = _soxContext.QuestionnaireTesterSet
                            .Where(y => y.Rcm.Id.Equals(item.Id))
                            .Include(stat => stat.WorkpaperStatus)
                            .OrderByDescending(c => c.Position)
                            .Take(1)
                            .AsNoTracking()
                            .FirstOrDefault();
                        if (workpaperStatus != null)
                        {
                            item.WorkpaperStatus = workpaperStatus.WorkpaperStatus.StatusName;
                        }
                    }
                }
                return Ok(_rcm.ToArray());
            }
            else
            {
                return NoContent();
            }

        }

        [AllowAnonymous]
        [HttpPost("rcm/create/")]
        public async Task<IActionResult> SaveRcmAsync([FromBody] RcmCta rcm)
        {

            bool status = false;
            try
            {
                using (var context = _soxContext.Database.BeginTransaction())
                {
                    RcmCta _rcm = new RcmCta();
                    _rcm = rcm;

                    bool isExists = false;
                    var rcmCheck = _soxContext.RcmCta.FirstOrDefault(id => id.PodioItemId == rcm.PodioItemId);
                    if (rcmCheck != null)
                    {
                        isExists = true;
                    }

                    if (isExists)
                    {
                        //if exists we update database row
                        //_soxContext.Update(_rcm);
                        rcm.Id = rcmCheck.Id;
                        _soxContext.Entry(rcmCheck).CurrentValues.SetValues(rcm);
                    }
                    else
                    {
                        //create new item
                        _soxContext.Add(_rcm);
                    }

                    await _soxContext.SaveChangesAsync();
                    context.Commit();
                    status = true;
                }
            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex.ToString());
                //ErrorLog.Write(ex);
                FileLog.Write($"Error SaveRcmAsync {ex}", "ErrorSaveRcmAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SaveRcmAsync");
            }
            if (status)
            {
                return Ok();
            }
            else
            {
                return NoContent();
            }


        }

        [AllowAnonymous]
        [HttpPut("rcm/update/")]
        public async Task<IActionResult> UpdateRcmAsync([FromBody] Rcm rcm)
        {
            bool status = false;
            try
            {
                using (var context = _soxContext.Database.BeginTransaction())
                {

                    var rcmCheck = _soxContext.Rcm.Where(id => id.PodioItemId.Equals(rcm.PodioItemId)).FirstOrDefault();

                    if (rcmCheck != null)
                    {
                        rcm.Id = rcmCheck.Id;
                        _soxContext.Entry(rcmCheck).CurrentValues.SetValues(rcm);
                        await _soxContext.SaveChangesAsync();
                        context.Commit();
                        status = true;
                    }

                }
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error UpdateRcmAsync {ex}", "ErrorUpdateRcmAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "UpdateRcmAsync");
            }

            if (status)
            {
                return Ok();
            }
            else
            {
                return NoContent();
            }

        }


        /// <summary>
        /// Api Route for RCM Question 1
        /// </summary>
        /// <returns>List<string></returns>
        [AllowAnonymous]
        [HttpGet("fy")]
        public IActionResult GetListRcmFYAsync()
        {
            List<string> _listFY = new List<string>();
            try
            {
                using (var context = _soxContext.Database.BeginTransaction())
                {
                    //Get all active FY
                    var listFy = _soxContext.Rcm
                        .Where(x => !x.Status.Equals("Completed") && !x.Status.Equals("Closed"))
                        .Select(x => x.FY)
                        .AsNoTracking();
                        
                    //check if value is not null
                    if (listFy != null)
                    {
                        _listFY = listFy.Distinct().OrderBy(x => x).ToList();

                    }
                }
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetListRcmFYAsync {ex}", "ErrorGetListRcmFYAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetListRcmFYAsync");
            }

            if (_listFY != null)
            {
                _listFY.Sort(); //sort list 
                return Ok(_listFY.ToArray());
            }
            else
            {
                return NoContent();
            }

        }

        /// <summary>
        /// Api Route for RCM Question 2
        /// </summary>
        /// <param name="filter"></param>
        /// <returns>List<string></returns>
        //[AllowAnonymous]
        [HttpPost("clientbyyear")]
        public IActionResult GetListClientAsync([FromBody] RcmQuestionnaireFilter filter)
        {
            List<RcmOutputFile> _listClient = new List<RcmOutputFile>();
            try
            {
                string rcmSharefile = _config.GetSection("SharefileApi").GetSection("SoxRcmFolder").GetSection("Link").Value;
                using (var context = _soxContext.Database.BeginTransaction())
                {
                    var listRcm = _soxContext.Rcm
                        .Where(x => 
                            !x.Status.Equals("Completed") && 
                            !x.Status.Equals("Closed") && 
                            x.FY.Equals(filter.FY) &&
                            x.ClientName != string.Empty &&
                            x.ClientName != null
                        )
                        .Select(x => x.ClientName)
                        .AsNoTracking();

                    if (listRcm != null)
                    {
                        var listClient = listRcm.Distinct().OrderBy(x => x).ToList();
                        foreach (var item in listClient)
                        {
                            _listClient.Add(new RcmOutputFile
                            {
                                ClientName = item,
                                LoadingStatus = string.Empty,
                                SharefileLink = rcmSharefile,
                                FileName = string.Empty
                            });
                        }
                        //_listClient = listRcm.Distinct().OrderBy(x => x).ToList();
                    }
                }
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetListClientAsync {ex}", "ErrorGetListClientAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetListClientAsync");
            }

            if (_listClient != null)
            {
                return Ok(_listClient.ToArray());
            }
            else
            {
                return NoContent();
            }

        }


        /// <summary>
        /// Api Route for RCM Question 1
        /// </summary>
        /// <returns>List<string></returns>
        [AllowAnonymous]
        [HttpGet("q1fy")]
        public IActionResult GetListQ1FYAsync()
        {
            List<string> _listFY = new List<string>();
            try
            {
                using (var context = _soxContext.Database.BeginTransaction())
                {
                    //Get all active FY
                    var listFy = _soxContext.RcmQuestionnaire
                        .Where(x => x.Status.ToLower().Equals("active"))
                        .Include(x => x.ListQ1Year)
                        .Select(x => x.ListQ1Year)
                        .AsNoTracking();
                        
                        
                    //check if value is not null
                    if (listFy != null)
                    {
                        //loop through all RcmFY items
                        foreach (var item in listFy)
                        {
                            var tempListFY = item.Select(x => x.FY);
                            foreach (var itemFy in tempListFY)
                            {
                                if (item != null && !_listFY.Contains(itemFy))
                                {
                                    //Add FY items in list and return
                                    _listFY.Add(itemFy);
                                }
                            } 
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetListQ1FYAsync {ex}", "ErrorGetListQ1FYAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetListQ1FYAsync");
            }

            if (_listFY != null)
            {
                _listFY.Sort(); //sort list 
                return Ok(_listFY.ToArray());
            }
            else
            {
                return NoContent();
            }

        }


        /// <summary>
        /// Api Route for RCM Question 2
        /// </summary>
        /// <param name="filter"></param>
        /// <returns>List<string></returns>
        [AllowAnonymous]
        [HttpPost("q2client")]
        public IActionResult GetListQ2ClientAsync([FromBody] RcmQuestionnaireFilter filter)
        {
            List<string> _listClient = new List<string>();
            try
            {
                using (var context = _soxContext.Database.BeginTransaction())
                {
                    var listRcm = _soxContext.RcmQuestionnaire
                        //.Where(x => x.Status.Equals("Active") && x.ListQ1Year.Where(x => x.FY.Contains(Fy)))
                        .Where(x => x.Status.ToLower().Equals("active"))
                        .Include(x => x.ListQ1Year)
                        .AsNoTracking();

                    if (listRcm != null)
                    {
                        foreach (var item in listRcm)
                        {

                            //foreach (var itemFy in item.ListQ1Year)
                            //{
                            //    if(itemFy.FY.Equals(filter.FY) && !_listClient.Contains(item.Q2Client))
                            //    {
                            //        _listClient.Add(item.Q2Client);
                            //    }
                            //}

                            var itemFy = item.ListQ1Year.Where(x => x.FY.Equals(filter.FY)).FirstOrDefault();
                            if (itemFy != null && !_listClient.Contains(item.Q2Client))
                            {
                                _listClient.Add(item.Q2Client);
                            }

                        }  
                    }
                }
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetListQ2ClientAsync {ex}", "ErrorGetListQ2ClientAsync");
                AdminService adminService = new AdminService(_config);
              //  adminService.SendAlert(true, true, ex.ToString(), "GetListQ2ClientAsync");
            }

            if (_listClient != null)
            {
                _listClient.Sort(); //sort list 
                return Ok(_listClient.ToArray());
            }
            else
            {
                return NoContent();
            }

        }


        /// <summary>
        /// Api Route for RCM Question 3
        /// </summary>
        /// <param name="filter"></param>
        /// <returns>List<string></returns>
        [AllowAnonymous]
        [HttpPost("q3process")]
        public IActionResult GetListQ3ProcessAsync([FromBody] RcmQuestionnaireFilter filter)
        {
            List<string> _listProcess = new List<string>();
            try
            {
                using (var context = _soxContext.Database.BeginTransaction())
                {
                    var listRcm = _soxContext.RcmQuestionnaire
                        .Where(x => x.Status.ToLower().Equals("active") && x.Q2Client.Equals(filter.Client))
                        .Include(x => x.ListQ1Year)
                        .AsNoTracking();

                    if (listRcm != null)
                    {
                        foreach (var item in listRcm)
                        {

                            //foreach (var itemFy in item.ListQ1Year)
                            //{
                            //    if (itemFy.FY.Equals(filter.FY) && !_listProcess.Contains(item.Q3Process))
                            //    {
                            //        _listProcess.Add(item.Q3Process);
                            //    }
                            //}

                            var itemFy = item.ListQ1Year.Where(x => x.FY.Equals(filter.FY)).FirstOrDefault();
                            if (itemFy != null && !_listProcess.Contains(item.Q3Process))
                            {
                                _listProcess.Add(item.Q3Process);
                            }


                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetListQ3ProcessAsync {ex}", "ErrorGetListQ3ProcessAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetListQ3ProcessAsync");
            }

            if (_listProcess != null)
            {
                _listProcess.Sort(); //sort list 
                return Ok(_listProcess.ToArray());
            }
            else
            {
                return NoContent();
            }

        }


        /// <summary>
        /// Api Route for RCM Question 3
        /// </summary>
        /// <param name="filter"></param>
        /// <returns>List<string></returns>
        [AllowAnonymous]
        [HttpPost("q4subprocess")]
        public IActionResult GetListQ4SubProcessAsync([FromBody] RcmQuestionnaireFilter filter)
        {
            List<string> _listSubProcess = new List<string>();
            try
            {
                using (var context = _soxContext.Database.BeginTransaction())
                {
                    var listRcm = _soxContext.RcmQuestionnaire
                        .Where(x => x.Status.ToLower().Equals("active") 
                            && x.Q2Client.Equals(filter.Client)
                            && x.Q3Process.Equals(filter.Process))
                        .Include(x => x.ListQ1Year)
                        .AsNoTracking();

                    if (listRcm != null)
                    {
                        foreach (var item in listRcm)
                        {

                            //foreach (var itemFy in item.ListQ1Year)
                            //{
                            //    if (itemFy.FY.Equals(filter.FY) && !_listSubProcess.Contains(item.Q4SubProcess))
                            //    {
                            //        _listSubProcess.Add(item.Q4SubProcess);
                            //    }
                            //}

                            var itemFy = item.ListQ1Year.Where(x => x.FY.Equals(filter.FY)).FirstOrDefault();
                            if (itemFy != null && !_listSubProcess.Contains(item.Q4SubProcess))
                            {
                                _listSubProcess.Add(item.Q4SubProcess);
                            }

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetListQ4SubProcessAsync {ex}", "ErrorGetListQ4SubProcessAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetListQ4SubProcessAsync");
            }

            if (_listSubProcess != null)
            {
                _listSubProcess.Sort(); //sort list 
                return Ok(_listSubProcess.ToArray());
            }
            else
            {
                return NoContent();
            }

        }


        /// <summary>
        /// Api Route for RCM Question 9
        /// </summary>
        /// <param name="filter"></param>
        /// <returns>List<string></returns>
        [AllowAnonymous]
        [HttpPost("q9controlid")]
        public IActionResult GetListQ9ControlIdAsync([FromBody] RcmQuestionnaireFilter filter)
        {
            List<string> _listControlId = new List<string>();
            try
            {
                using (var context = _soxContext.Database.BeginTransaction())
                {
                    var listRcm = _soxContext.RcmQuestionnaire
                        .Where(x => x.Status.ToLower().Equals("Active")
                            && x.Q2Client.Equals(filter.Client)
                            && x.Q3Process.Equals(filter.Process)
                            && x.Q4SubProcess.Equals(filter.SubProcess))
                        .Include(x => x.ListQ1Year)
                        .Include(x => x.ListQ9ControlId)
                        .AsNoTracking();

                    if (listRcm != null)
                    {
                        foreach (var item in listRcm)
                        {

                            var itemFy = item.ListQ1Year.Where(x => x.FY.Equals(filter.FY)).FirstOrDefault();
                            if(itemFy != null)
                            {
                                foreach (var itemControlId in item.ListQ9ControlId)
                                {
                                    if (!_listControlId.Contains(itemControlId.ControlId))
                                    {
                                        _listControlId.Add(itemControlId.ControlId);
                                    }
                                }
                            }

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetListQ9ControlIdAsync {ex}", "ErrorGetListQ9ControlIdAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetListQ9ControlIdAsync");
            }

            if (_listControlId != null)
            {
                _listControlId.Sort(); //sort list 
                return Ok(_listControlId.ToArray());
            }
            else
            {
                return NoContent();
            }

        }


        /// <summary>
        /// Api Route for RCM Question Question 7,8 13 to Question 19
        /// </summary>
        /// <param name="filter"></param>
        /// <returns>RcmQ13toQ19</returns>
        [AllowAnonymous]
        [HttpPost("q13toq19")]
        public IActionResult GetListQ13toQ19([FromBody] RcmQuestionnaireFilter filter)
        {
            RcmQ13toQ19 q13toQ19 = new RcmQ13toQ19();
            List<RcmQuestionnaire> _listRcm = new List<RcmQuestionnaire>();
            try
            {
                using (var context = _soxContext.Database.BeginTransaction())
                {
                    var listRcm = _soxContext.RcmQuestionnaire
                        .Where(x => x.Status.ToLower().Equals("active")
                            && x.Q2Client.Equals(filter.Client)
                            && x.Q3Process.Equals(filter.Process)
                            && x.Q4SubProcess.Equals(filter.SubProcess))
                        .Include(x => x.ListQ1Year)
                        .AsNoTracking();

                    if (listRcm != null)
                    {
                        _listRcm = listRcm.ToList();
                    }
                }

                if(_listRcm != null)
                {
                    foreach (var item in _listRcm)
                    {
                        using (var context = _soxContext.Database.BeginTransaction())
                        {
                            var tempQ7 = _soxContext.RcmFinancialStatement.Where(x => x.RcmQuestionnaire.Equals(item)).Select(x => x.FinStatement);
                            if (tempQ7 != null)
                            {
                                List<CheckBoxItem> listCbiQ7 = new List<CheckBoxItem>();
                                foreach (var itemQ7 in tempQ7)
                                {
                                    CheckBoxItem q7Cbi = new CheckBoxItem();
                                    q7Cbi.selected = false;
                                    q7Cbi.item = itemQ7;
                                    listCbiQ7.Add(q7Cbi);
                                    
                                }
                                q13toQ19.ListQ7FinStatementElement = listCbiQ7;
                                q13toQ19.ListQ7FinStatementElement = q13toQ19.ListQ7FinStatementElement.OrderBy(x => x.item).ToList();
                            }

                            var tempQ8 = _soxContext.RcmFinancialStatementAssert.Where(x => x.RcmQuestionnaire.Equals(item)).Select(x => x.FinStatementAssert);
                            if (tempQ8 != null)
                            {
                                //q13toQ19.ListQ8FinStatementAssert = tempQ8.ToList();
                                //q13toQ19.ListQ8FinStatementAssert.Sort(); //sort list 
                                List<CheckBoxItem> listCbiQ8 = new List<CheckBoxItem>();
                                foreach (var itemQ8 in tempQ8)
                                {
                                    CheckBoxItem q8Cbi = new CheckBoxItem();
                                    q8Cbi.selected = false;
                                    q8Cbi.item = itemQ8;
                                    listCbiQ8.Add(q8Cbi);

                                }
                                q13toQ19.ListQ8FinStatementAssert = listCbiQ8;
                                q13toQ19.ListQ8FinStatementAssert = q13toQ19.ListQ8FinStatementAssert.OrderBy(x => x.item).ToList();
                            }

                            var tempQ12 = _soxContext.RcmControlOwner.Where(x => x.RcmQuestionnaire.Equals(item)).Select(x => x.ControlOwner);
                            if (tempQ12 != null)
                            {
                                //q13toQ19.ListQ12ControlOwner = tempQ12.ToList();
                                //q13toQ19.ListQ12ControlOwner.Sort(); //sort list 
                                bool isOthersFound = false;
                                List<CheckBoxItem> listCbiQ12 = new List<CheckBoxItem>();
                                foreach (var itemQ12 in tempQ12)
                                {
                                    if(!itemQ12.Equals("Others"))
                                    {
                                        CheckBoxItem q12Cbi = new CheckBoxItem();
                                        q12Cbi.selected = false;
                                        q12Cbi.item = itemQ12;
                                        listCbiQ12.Add(q12Cbi);
                                    }   
                                    else
                                    {
                                        isOthersFound = true;
                                    }
                                }
                                q13toQ19.ListQ12ControlOwner = listCbiQ12;
                                q13toQ19.ListQ12ControlOwner = q13toQ19.ListQ12ControlOwner.OrderBy(x => x.item).ToList();
                                
                                //add "others" at the bottom
                                if(isOthersFound)
                                {
                                    CheckBoxItem q12Cbi = new CheckBoxItem();
                                    q12Cbi.selected = false;
                                    q12Cbi.item = "Others";
                                    q13toQ19.ListQ12ControlOwner.Add(q12Cbi);
                                }
                            }

                            var tempQ13 = _soxContext.RcmFrequency.Where(x => x.RcmQuestionnaire.Equals(item)).Select(x => x.Frequency);
                            if (tempQ13 != null)
                            {
                                q13toQ19.ListQ13Frequency = tempQ13.ToList();
                                q13toQ19.ListQ13Frequency.Sort(); //sort list 
                            }

                            var tempQ14 = _soxContext.RcmControlKey.Where(x => x.RcmQuestionnaire.Equals(item)).Select(x => x.Option);
                            if (tempQ14 != null)
                            {
                                q13toQ19.ListQ14ControlKey = tempQ14.ToList();
                                q13toQ19.ListQ14ControlKey.Sort(); //sort list 
                            }

                            var tempQ15 = _soxContext.RcmControlType.Where(x => x.RcmQuestionnaire.Equals(item)).Select(x => x.ControlType);
                            if (tempQ15 != null)
                            {
                                q13toQ19.ListQ15ControlType = tempQ15.ToList();
                                q13toQ19.ListQ15ControlType.Sort(); //sort list 
                            }

                            var tempQ16 = _soxContext.RcmNatureProcedure.Where(x => x.RcmQuestionnaire.Equals(item)).Select(x => x.NatureProcedure);
                            if (tempQ16 != null)
                            {
                                q13toQ19.ListQ16NatureProcedure = tempQ16.ToList();
                                q13toQ19.ListQ16NatureProcedure.Sort(); //sort list 
                            }

                            var tempQ17 = _soxContext.RcmFraudControl.Where(x => x.RcmQuestionnaire.Equals(item)).Select(x => x.Option);
                            if (tempQ17 != null)
                            {
                                q13toQ19.ListQ17FraudControl = tempQ17.ToList();
                                q13toQ19.ListQ17FraudControl.Sort(); //sort list 
                            }

                            var tempQ18 = _soxContext.RcmRiskLevel.Where(x => x.RcmQuestionnaire.Equals(item)).Select(x => x.Option);
                            if (tempQ18 != null)
                            {
                                q13toQ19.ListQ18RiskLevel = tempQ18.ToList();
                                q13toQ19.ListQ18RiskLevel.Sort(); //sort list 
                            }

                            var tempQ19 = _soxContext.RcmManagementReviewControl.Where(x => x.RcmQuestionnaire.Equals(item)).Select(x => x.MgmtReviewControl);
                            if (tempQ19 != null)
                            {
                                q13toQ19.ListQ19MgmtReviewControl = tempQ19.ToList();
                                q13toQ19.ListQ19MgmtReviewControl.Sort(); //sort list 
                            }

                        }

   

                    }
                    
                }

            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetListQ9ControlIdAsync {ex}", "ErrorGetListQ9ControlIdAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetListQ9ControlIdAsync");
                if (_environment.IsDevelopment())
                {
                    return BadRequest(ex.ToString());
                }
                
            }

            if (q13toQ19 != null)
            {
                return Ok(q13toQ19);
            }
            else
            {
                return NoContent();
            }

        }


        /// <summary>
        /// Save RCM to Podio
        /// </summary>
        /// <param name="rcm"></param>
        /// <returns>rcm</returns>
        [AllowAnonymous]
        [HttpPost("podiosave")]
        public async Task<IActionResult> CreatePodioRcmAsync([FromBody] Rcm rcm)
        {
            bool status = false;
            RcmService rcmService = new RcmService(_soxContext, _config);
            try
            {
                rcm = await rcmService.SavePodioRcm(rcm); //return podio item id if success
                if (rcm.PodioItemId != 0)
                    status = true;

            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex.ToString());
                //ErrorLog.Write(ex);
                FileLog.Write(ex.ToString(), "ErrorCreatePodioRcmAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "CreatePodioRcmAsync");
                if (_environment.IsDevelopment())
                {
                    return BadRequest(ex.ToString());
                }
                else
                {
                    return BadRequest();
                }
                
            }
            if (status)
            {
                return Ok(rcm);
            }
            else
            {
                return NoContent();
            }

        }


        /// <summary>
        /// Save RCM to database
        /// </summary>
        /// <param name="rcm"></param>
        /// <returns></returns>
        [AllowAnonymous]
        [HttpPost("save")]
        public async Task<IActionResult> SaveToDatabase([FromBody] Rcm rcm)
        {
            RcmService rcmService = new RcmService(_soxContext, _config);
            SoxTrackerService soxTrackerService = new SoxTrackerService(_soxContext, _config);
            bool status = false;
            try
            {
                if(await rcmService.SaveRcmToDatabase(rcm))
                    status = true;
                    
                
                   // status = true;
            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSaveToDatabase");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SaveToDatabase");
                if (_environment.IsDevelopment())
                {
                    return BadRequest(ex.ToString());
                }
                else
                {
                    return BadRequest();
                }
                
            }
            if (status)
            {
                return Ok();
            }
            else
            {
                return NoContent();
            }



        }

        [AllowAnonymous]
        [HttpPost("getcontrol")]
        public IActionResult GetRcmControl([FromBody] RcmQuestionnaireFilter filter)
        {
            Rcm _rcm = new Rcm();
            try
            {
                using (var context = _soxContext.Database.BeginTransaction())
                {
                     _rcm = _soxContext.Rcm
                        .Where(x => !x.Status.ToLower().Equals("closed")
                            && !x.Status.ToLower().Equals("completed")
                            && x.ClientName.Equals(filter.Client)
                            && x.Process.Equals(filter.Process)
                            && x.Subprocess.Equals(filter.SubProcess)
                            && x.FY.Equals(filter.FY)
                            && x.ControlId.Equals(filter.ControlId))
                        .FirstOrDefault();

                    if (_rcm != null && _rcm.PodioItemId != 0)
                    {
                        return Ok(_rcm);
                    }
                    else
                    {
                        return NoContent();
                    }
                }
                   
                
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetRcmControl {ex}", "ErrorGetRcmControl");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetRcmControl");
                if (_environment.IsDevelopment())
                {
                    return BadRequest(ex.ToString());
                }
                else
                {
                    return BadRequest();
                }
            }

            
           

        }

        [AllowAnonymous]
        [HttpDelete("delete/{itemId:int}")]
        public async Task<IActionResult> DeleteRcmAsync(int itemId)
        {
            Rcm _rcm = new Rcm();
            try
            {
                using (var context = _soxContext.Database.BeginTransaction())
                {
                    _rcm = _soxContext.Rcm
                       .Where(x => x.PodioItemId.Equals(itemId))
                       .FirstOrDefault();

                    if(_rcm != null)
                    {

                        _soxContext.Remove(_rcm);
                        await _soxContext.SaveChangesAsync();
                        context.Commit();

                        return Ok();
                    }
                    else
                    {
                        return NoContent();
                    }
                }

                

            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error DeleteRcmAsync {ex}", "ErrorDeleteRcmAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "DeleteRcmAsync");
                if (_environment.IsDevelopment())
                {
                    return BadRequest(ex.ToString());
                }
                else
                {
                    return BadRequest();
                }
            }




        }

        [AllowAnonymous]
        [HttpPost("podioupdate")]
        public async Task<IActionResult> UpdatePodioRcm([FromBody] Rcm rcm)
        {
            bool status = false;

            try
            {
                
                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("RcmPodioApp").GetSection("RcmAppId").Value;
                PodioAppKey.AppToken = _config.GetSection("RcmPodioApp").GetSection("RcmAppToken").Value;

                if (PodioAppKey.AppId != string.Empty && PodioAppKey.AppToken != string.Empty)
                {
                    var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                    ServicePointManager.Expect100Continue = true;
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                    await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                    if (podio.IsAuthenticated() && rcm != null)
                    {

                        Item rcmItem = new Item
                        {
                            ItemId = rcm.PodioItemId
                        };

                        #region Podio Fields
                        int q1Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q1").Value);
                        int q2Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q2").Value);
                        int q3Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q3").Value);                     
                        int q4Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q4").Value);
                        int q5Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q5").Value);
                        int q6Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q6").Value);
                        int q7Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q7").Value);
                        int q8Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q8").Value);
                        int q8ExistenceField = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q8Existence").Value);
                        int q8PresentationField = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q8Presentation").Value);
                        int q8RightsField = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q8Rights").Value);
                        int q8ValuationField = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q8Valuation").Value);
                        int q9Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q9").Value);
                        int q10Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q10").Value);
                        int shortDescField = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("ShortDesc").Value);
                        int q11Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q11").Value);
                        int q12Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q12").Value);
                        int q13Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q13").Value);
                        int q14Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q14").Value);
                        int q15Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q15").Value);
                        int q16Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q16").Value);
                        int q17Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q17").Value);
                        int q18Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q18").Value);
                        int q19Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q19").Value);
                        int q20Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q20").Value);
                        int q21Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q21").Value);
                        int durationField = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Duration").Value);


                        if (rcm.FY != string.Empty && rcm.FY != null)
                        {
                            
                            var textQ1 = rcmItem.Field<TextItemField>(q1Field);
                            textQ1.Value = rcm.FY;
                        }
                        else
                        {
                            var textQ1 = rcmItem.Field<TextItemField>(q1Field);
                            textQ1.Value = " ";
                        }

                        if (rcm.ClientName != string.Empty && rcm.ClientName != null)
                        {
                            var checkClientItemId = _soxContext.ClientSs.Where(x => x.ClientName.Equals(rcm.ClientName)).Select(x => x.ItemId).FirstOrDefault();
                            if (checkClientItemId != null)
                            {
                                
                                var appReference = rcmItem.Field<AppItemField>(q2Field);
                                List<int> listRoundItem = new List<int>();
                                listRoundItem.Add(checkClientItemId.Value);
                                appReference.ItemIds = listRoundItem;
                            }
                        }

                        if (rcm.Process != string.Empty && rcm.Process != null)
                        {       
                            var textQ3 = rcmItem.Field<TextItemField>(q3Field);
                            textQ3.Value = rcm.Process;
                        }
                        else
                        {
                            var textQ3 = rcmItem.Field<TextItemField>(q3Field);
                            textQ3.Value = " ";
                        }

                        if (rcm.Subprocess != string.Empty && rcm.Subprocess != null)
                        {
                            
                            var textQ4 = rcmItem.Field<TextItemField>(q4Field);
                            textQ4.Value = rcm.Subprocess;
                        }
                        else
                        {
                            var textQ4 = rcmItem.Field<TextItemField>(q4Field);
                            textQ4.Value = " ";
                        }

                        if (rcm.ControlObjective != string.Empty && rcm.ControlObjective != null)
                        {
                            
                            var textQ5 = rcmItem.Field<TextItemField>(q5Field);
                            textQ5.Value = rcm.ControlObjective;
                        }
                        else
                        {
                            var textQ5 = rcmItem.Field<TextItemField>(q5Field);
                            textQ5.Value = " ";
                        }

                        if (rcm.SpecificRisk != string.Empty && rcm.SpecificRisk != null)
                        {
                            
                            var textQ6 = rcmItem.Field<TextItemField>(q6Field);
                            textQ6.Value = rcm.SpecificRisk;
                        }
                        else
                        {
                            var textQ6 = rcmItem.Field<TextItemField>(q6Field);
                            textQ6.Value = " ";
                        }

                        if (rcm.FinStatemenElement != string.Empty && rcm.FinStatemenElement != null)
                        {
                            var textQ7 = rcmItem.Field<TextItemField>(q7Field);
                            textQ7.Value = rcm.FinStatemenElement;
                        }
                        else
                        {
                            var textQ7 = rcmItem.Field<TextItemField>(q7Field);
                            textQ7.Value = " ";
                        }

                        //Financial Statement Assertion

                        if (rcm.FinancialStatementAssert != string.Empty && rcm.FinancialStatementAssert != null)
                        {
                            var textQ8 = rcmItem.Field<TextItemField>(q8Field);
                            if (rcm.FinancialStatementAssert.ToLower().Contains("accuracy"))
                            {
                                textQ8.Value = "Yes";
                                rcm.CompletenessAccuracy = "Yes";
                            }
                            else
                            {
                                textQ8.Value = "No";
                                rcm.CompletenessAccuracy = "No";
                            }

                        }

                        if (rcm.FinancialStatementAssert != string.Empty && rcm.FinancialStatementAssert != null)
                        {                   
                            var textQ8Existence = rcmItem.Field<TextItemField>(q8ExistenceField);
                            if (rcm.FinancialStatementAssert.ToLower().Contains("existence"))
                            {
                                textQ8Existence.Value = "Yes";
                                rcm.ExistenceDisclosure = "Yes";
                            }
                            else
                            {
                                textQ8Existence.Value = "No";
                                rcm.ExistenceDisclosure = "No";
                            }
                                
                        }

                        if (rcm.FinancialStatementAssert != string.Empty && rcm.FinancialStatementAssert != null)
                        {                          
                            var textQ8Presentation = rcmItem.Field<TextItemField>(q8PresentationField);
                            if (rcm.FinancialStatementAssert.ToLower().Contains("presentation"))
                            {
                                textQ8Presentation.Value = "Yes";
                                rcm.PresentationDisclosure = "Yes";
                            }

                            else
                            {
                                textQ8Presentation.Value = "No";
                                rcm.PresentationDisclosure = "No";
                            }
                                
                        }

                        if (rcm.FinancialStatementAssert != string.Empty && rcm.FinancialStatementAssert != null)
                        {
                            var textQ8Rights = rcmItem.Field<TextItemField>(q8RightsField);
                            if (rcm.FinancialStatementAssert.ToLower().Contains("rights"))
                            {
                                textQ8Rights.Value = "Yes";
                                rcm.RigthsObligation = "Yes";
                            }
                            else
                            {
                                textQ8Rights.Value = "No";
                                rcm.RigthsObligation = "No";
                            }
                                
                        }

                        if (rcm.FinancialStatementAssert != string.Empty && rcm.FinancialStatementAssert != null)
                        {
                            var textQ8Valuation = rcmItem.Field<TextItemField>(q8ValuationField);
                            if (rcm.FinancialStatementAssert.ToLower().Contains("valuation"))
                            {
                                textQ8Valuation.Value = "Yes";
                                rcm.ValuationAlloc = "Yes";
                            }
                            else
                            {
                                textQ8Valuation.Value = "No";
                                rcm.ValuationAlloc = "No";
                            }
                                
                        }
                        //Financial Statement Assertion

                        if (rcm.ControlId != string.Empty && rcm.ControlId != null)
                        {
                            
                            var textQ9 = rcmItem.Field<TextItemField>(q9Field);
                            textQ9.Value = rcm.ControlId;
                        }
                        else
                        {
                            var textQ9 = rcmItem.Field<TextItemField>(q9Field);
                            textQ9.Value = " ";
                        }


                        if (rcm.ControlActivity != string.Empty && rcm.ControlActivity != null)
                        {
                            
                            var textQ10 = rcmItem.Field<TextItemField>(q10Field);
                            textQ10.Value = rcm.ControlActivity;
                        }
                        else
                        {
                            var textQ10 = rcmItem.Field<TextItemField>(q10Field);
                            textQ10.Value = " ";
                        }

                        if (rcm.ShortDescription != string.Empty && rcm.ShortDescription != null)
                        {

                            var textShortDesc = rcmItem.Field<TextItemField>(shortDescField);
                            textShortDesc.Value = rcm.ShortDescription;
                        }
                        else
                        {
                            var textShortDesc = rcmItem.Field<TextItemField>(shortDescField);
                            textShortDesc.Value = " ";
                        }

                        if (rcm.ControlPlaceDate != string.Empty && rcm.ControlPlaceDate != null)
                        {
                            
                            var textQ11 = rcmItem.Field<TextItemField>(q11Field);
                            textQ11.Value = rcm.ControlPlaceDate;
                        }
                        else
                        {
                            var textQ11 = rcmItem.Field<TextItemField>(q11Field);
                            textQ11.Value = " ";
                        }


                        if (rcm.ControlOwner != string.Empty && rcm.ControlOwner != null)
                        {
                            
                            var textQ12 = rcmItem.Field<TextItemField>(q12Field);
                            textQ12.Value = rcm.ControlOwner;
                        }
                        else
                        {
                            var textQ12 = rcmItem.Field<TextItemField>(q12Field);
                            textQ12.Value = " ";
                        }


                        if (rcm.ControlFrequency != string.Empty && rcm.ControlFrequency != null)
                        {
                           
                            var textQ13 = rcmItem.Field<TextItemField>(q13Field);
                            textQ13.Value = rcm.ControlFrequency;
                        }
                        else
                        {
                            var textQ13 = rcmItem.Field<TextItemField>(q13Field);
                            textQ13.Value = " ";
                        }


                        if (rcm.Key != string.Empty && rcm.Key != null)
                        {
                            
                            var textQ14 = rcmItem.Field<TextItemField>(q14Field);
                            textQ14.Value = rcm.Key;
                        }
                        else
                        {
                            var textQ14 = rcmItem.Field<TextItemField>(q14Field);
                            textQ14.Value = " ";
                        }


                        if (rcm.ControlType != string.Empty && rcm.ControlType != null)
                        {
                            
                            var textQ15 = rcmItem.Field<TextItemField>(q15Field);
                            textQ15.Value = rcm.ControlType;
                        }
                        else
                        {
                            var textQ15 = rcmItem.Field<TextItemField>(q15Field);
                            textQ15.Value = " ";
                        }


                        if (rcm.NatureProc != string.Empty && rcm.NatureProc != null)
                        {
                            
                            var textQ16 = rcmItem.Field<TextItemField>(q16Field);
                            textQ16.Value = rcm.NatureProc;
                        }
                        else
                        {
                            var textQ16 = rcmItem.Field<TextItemField>(q16Field);
                            textQ16.Value = " ";
                        }


                        if (rcm.FraudControl != string.Empty && rcm.FraudControl != null)
                        {
                            
                            var textQ17 = rcmItem.Field<TextItemField>(q17Field);
                            textQ17.Value = rcm.FraudControl;
                        }
                        else
                        {
                            var textQ17 = rcmItem.Field<TextItemField>(q17Field);
                            textQ17.Value = " ";
                        }


                        if (rcm.RiskLvl != string.Empty && rcm.RiskLvl != null)
                        {
                            
                            var textQ18 = rcmItem.Field<TextItemField>(q18Field);
                            textQ18.Value = rcm.RiskLvl;
                        }
                        else
                        {
                            var textQ18 = rcmItem.Field<TextItemField>(q18Field);
                            textQ18.Value = " ";
                        }


                        if (rcm.ManagementRevControl != string.Empty && rcm.ManagementRevControl != null)
                        {
                            
                            var textQ19 = rcmItem.Field<TextItemField>(q19Field);
                            textQ19.Value = rcm.ManagementRevControl;
                        }
                        else
                        {
                            var textQ19 = rcmItem.Field<TextItemField>(q19Field);
                            textQ19.Value = " ";
                        }


                        if (rcm.PbcList != string.Empty && rcm.PbcList != null)
                        {
                            
                            var text20 = rcmItem.Field<TextItemField>(q20Field);
                            text20.Value = rcm.PbcList;
                        }
                        else
                        {
                            var text20 = rcmItem.Field<TextItemField>(q20Field);
                            text20.Value = " ";
                        }


                        if (rcm.TestProc != string.Empty && rcm.TestProc != null)
                        {
                            
                            var text21 = rcmItem.Field<TextItemField>(q21Field);
                            text21.Value = rcm.TestProc;
                        }
                        else
                        {
                            var text21 = rcmItem.Field<TextItemField>(q21Field);
                            text21.Value = " ";
                        }

                        //int statusField = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Status").Value);
                        //var categoryStatus = rcmItem.Field<CategoryItemField>(statusField);
                        //categoryStatus.OptionText = "Active";

                        if (rcm.Duration != null && rcm.Duration.HasValue)
                        {
                            
                            DurationItemField durationItem = rcmItem.Field<DurationItemField>(durationField);
                            durationItem.Value = rcm.Duration;
                        }



                        #endregion

                        var roundId = await podio.ItemService.UpdateItem(rcmItem, null, null, false, true);
                        if(roundId != null)
                        {
                            rcm.PodioItemId = roundId.Value;
                        }

                        status = true;
                    }
                }

            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex.ToString());
                //ErrorLog.Write(ex);
                FileLog.Write($"Error UpdatePodioRcm {ex}", "ErrorUpdatePodioRcm");
                //AdminService adminService = new AdminService(_config);
                //adminService.SendAlert(true, true, ex.ToString(), "UpdatePodioRcm");
                if (_environment.IsDevelopment())
                {
                    return BadRequest(ex.ToString());
                }
                else
                {
                    return BadRequest();
                }

            }
            if (status)
            {
                return Ok(rcm);
            }
            else
            {
                return NoContent();
            }


        }

        [AllowAnonymous]
        [HttpPost("update")]
        public async Task<IActionResult> UpdateDbRcm([FromBody] Rcm rcm)
        {
            Rcm _rcm = new Rcm();
            SoxTracker _soxtracker = new SoxTracker();
            try
            {
                using (var context = _soxContext.Database.BeginTransaction())
                {
                    _rcm = _soxContext.Rcm
                       .Where(x => x.Status.Equals(rcm.Status)
                           && x.ClientName.Equals(rcm.ClientName)
                           && x.Process.Equals(rcm.Process)
                           && x.Subprocess.Equals(rcm.Subprocess)
                           && x.FY.Equals(rcm.FY)
                           && x.ControlId.Equals(rcm.ControlId))
                       .FirstOrDefault();
                    if (_rcm != null)
                    {
                        _rcm.ControlObjective = rcm.ControlObjective;
                        _rcm.SpecificRisk = rcm.SpecificRisk;
                        _rcm.ControlActivity = rcm.ControlActivity;
                        _rcm.ControlPlaceDate = rcm.ControlPlaceDate;
                        _rcm.PbcList = rcm.PbcList;
                        _rcm.TestProc = rcm.TestProc;
                        _rcm.ControlFrequency = rcm.ControlFrequency;
                        _rcm.Key = rcm.Key;
                        _rcm.ControlType = rcm.ControlType;
                        _rcm.NatureProc = rcm.NatureProc;
                        _rcm.FraudControl = rcm.FraudControl;
                        _rcm.RiskLvl = rcm.RiskLvl;
                        _rcm.ManagementRevControl = rcm.ManagementRevControl;
                        _rcm.FinStatemenElement = rcm.FinStatemenElement;
                       // _rcm.CompletenessAccuracy = rcm.CompletenessAccuracy;
                        _rcm.ControlOwner = rcm.ControlOwner;
                        _rcm.FinancialStatementAssert = rcm.FinancialStatementAssert;
                        //added 0512
                        _rcm.CompletenessAccuracy = rcm.CompletenessAccuracy;
                        _rcm.ExistenceDisclosure = rcm.ExistenceDisclosure;
                        _rcm.PresentationDisclosure = rcm.PresentationDisclosure;
                        _rcm.RigthsObligation = rcm.RigthsObligation;
                        _rcm.ValuationAlloc = rcm.ValuationAlloc;





                        _soxContext.Update(_rcm);
                        await _soxContext.SaveChangesAsync();
                        context.Commit();
                    }
                    else
                    {
                        return NoContent();
                    }
                }
                using (var context = _soxContext.Database.BeginTransaction())
                {
      
                    _soxtracker = _soxContext.SoxTracker
                                .Where(x => x.ClientName.Equals(rcm.ClientName)
                           && x.Process.Equals(rcm.Process)
                           && x.Subprocess.Equals(rcm.Subprocess)
                           && x.FY.Equals(rcm.FY)
                           && x.ControlId.Equals(rcm.ControlId))
                       .FirstOrDefault();
                    if (_soxtracker != null)
                    {
                        _soxtracker.PBC = rcm.PbcList;
                        _soxContext.Update(_soxtracker);
                        await _soxContext.SaveChangesAsync();
                        context.Commit();
                    }
                    else
                    {
                        return NoContent();
                    }
                }
                return Ok(_rcm);


            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error UpdateDbRcm {ex}", "ErrorUpdateDbRcm");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "UpdateDbRcm");
                if (_environment.IsDevelopment())
                {
                    return BadRequest(ex.ToString());
                }
                else
                {
                    return BadRequest();
                }
            }




        }


        //[AllowAnonymous]
        [HttpPost("generate/{clientName}/{Fy}")]
        public IActionResult GenerateRcmFile(string clientName, string Fy)
        {
            //List<string> excelFilename = new List<string>();
            string excelFilename = string.Empty;

            try
            {

                //Get RCM
                List<Rcm> _rcm = new List<Rcm>();
                var checkRcm = _soxContext.Rcm
                    .Where(x => 
                        x.ClientName.Equals(clientName) 
                        && x.FY.Equals(Fy)
                        && x.Status.ToLower() != "closed" 
                        && x.Status.ToLower() != "completed"
                    )
                    .OrderBy(x => x.Process)
                    .ThenBy(x => x.ControlId);

                if(checkRcm != null)
                {
                    _rcm = checkRcm.ToList();

                    #region Create Excel
                    //Creating excel 
                    //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    using (ExcelPackage xls = new ExcelPackage())
                    {
                        ExcelService xlsService = new ExcelService();
                        FormatService txtFormat = new FormatService();
                        //set sheet name
                        var ws = xls.Workbook.Worksheets.Add("RCM");


                        ws.View.FreezePanes(7, 12);

                        ws.View.ZoomScale = 80;
                        //set column width
                        ws.Column(1).Width = 11;
                        ws.Column(2).Width = 15;
                        ws.Column(3).Width = 25;
                        ws.Column(4).Width = 25;
                        ws.Column(5).Width = 25;
                        ws.Column(6).Width = 6;
                        ws.Column(7).Width = 6;
                        ws.Column(8).Width = 6;
                        ws.Column(9).Width = 6;
                        ws.Column(10).Width = 6;
                        ws.Column(11).Width = 10;
                        ws.Column(12).Width = 62;
                        ws.Column(13).Width = 62;
                        ws.Column(14).Width = 22;
                        ws.Column(15).Width = 22;
                        ws.Column(16).Width = 22;
                        ws.Column(17).Width = 10;
                        ws.Column(18).Width = 10;
                        ws.Column(19).Width = 10;
                        ws.Column(20).Width = 10;
                        ws.Column(21).Width = 10;
                        ws.Column(22).Width = 10;
                        ws.Column(23).Width = 60;
                        ws.Column(24).Width = 90;

                        //disable grid
                        ws.View.ShowGridLines = false;

                        //set row
                        int row = 1;

                        //set title
                        ws.Cells[row, 1].Value = clientName;
                        row++;
                        ws.Cells[row, 1].Value = "Risk Control Matrix (RCM)";
                        row++;
                        ws.Cells[row, 1].Value = Fy;
                        xlsService.ExcelSetFontBold(ws, 1, 1, row, 1); //(workspace, from row, from column, to row, to column)

                        ws.Cells[5, 6].Value = "Financial Assertions";
                        ws.Cells["F5:J5"].Merge = true;
                        
                        xlsService.ExcelSetBackgroundColorGray(ws, 5, 6, 5, 10);
                        xlsService.ExcelSetBorder(ws, 5, 6, 5, 10);
                        xlsService.ExcelSetHorizontalAlignCenter(ws, 5, 6, 5, 10); 
                        xlsService.ExcelSetVerticalAlignCenter(ws, 5, 6, 5, 10); 
                        row += 3;


                        ws.Row(row).Height = 85;
                        //set table header
                        ws.Cells[row, 1].Value = "Process";
                        ws.Cells[row, 2].Value = "Sub-process";
                        ws.Cells[row, 3].Value = "Control Objective";
                        ws.Cells[row, 4].Value = "Specific Risk";
                        ws.Cells[row, 5].Value = "Financial Statement Elements";
                        ws.Cells[row, 6].Value = "Completeness & Accuracy";
                        ws.Cells[row, 7].Value = "Existence & Occurrence";
                        ws.Cells[row, 8].Value = "Presentation & Disclosure";
                        ws.Cells[row, 9].Value = "Rights & Obligations";
                        ws.Cells[row, 10].Value = "Valuation & Allocation";
                        ws.Cells[row, 11].Value = "Control Id";
                        ws.Cells[row, 12].Value = "Control Activity FY";
                        ws.Cells[row, 13].Value = "Short Description";
                        ws.Cells[row, 14].Value = "Control In Place Date ";
                        ws.Cells[row, 15].Value = "Control Owner";
                        ws.Cells[row, 16].Value = "Control Frequency";
                        ws.Cells[row, 17].Value = "Key/Non - Key";
                        ws.Cells[row, 18].Value = "Control Type";
                        ws.Cells[row, 19].Value = "Nature of Procedure ";
                        ws.Cells[row, 20].Value = "Fraud Control";
                        ws.Cells[row, 21].Value = "Risk Level";
                        ws.Cells[row, 22].Value = "Management Review Control";
                        ws.Cells[row, 23].Value = "PBC List";
                        ws.Cells[row, 24].Value = "Test procedures";

                        //set vertical header
                        ws.Cells[row, 6].Style.TextRotation = 90;
                        ws.Cells[row, 7].Style.TextRotation = 90;
                        ws.Cells[row, 8].Style.TextRotation = 90;
                        ws.Cells[row, 9].Style.TextRotation = 90;
                        ws.Cells[row, 10].Style.TextRotation = 90;

                        //format cell
                        xlsService.ExcelWrapText(ws, row, 1, row, 24); //(workspace, from row, from column, to row, to column)
                        xlsService.ExcelSetHorizontalAlignCenter(ws, row, 1, row, 5); //(workspace, from row, from column, to row, to column)
                        xlsService.ExcelSetHorizontalAlignCenter(ws, row, 11, row, 24); //(workspace, from row, from column, to row, to column)
                        xlsService.ExcelSetVerticalAlignCenter(ws, row, 1, row, 5); //(workspace, from row, from column, to row, to column)
                        xlsService.ExcelSetVerticalAlignCenter(ws, row, 11, row, 24);  //(workspace, from row, from column, to row, to column)
                        xlsService.ExcelWrapText(ws, row, 1, row, 5); //(workspace, from row, from column, to row, to column)
                        xlsService.ExcelWrapText(ws, row, 11, row, 24); //(workspace, from row, from column, to row, to column)

                        //font color and font background
                        xlsService.ExcelSetFontColorRed(ws, row, 20, row, 24);
                        xlsService.ExcelSetBackgroundColorGray(ws, row, 1, row, 24);

                        //set border
                        xlsService.ExcelSetBorder(ws, row, 1, row, 24);

                        //set auto filter
                        ws.Cells[row, 1, row, 24].AutoFilter = true;

                        row++;

                        //loop through all RCM found
                        foreach (var item in _rcm)
                        {
                            ws.Cells[row, 1].Value = item.Process;
                            ws.Cells[row, 2].Value = item.Subprocess;
                            ws.Cells[row, 3].Value = item.ControlObjective != null ? txtFormat.FormatwithNewLine(item.ControlObjective, true) : string.Empty;
                            ws.Cells[row, 4].Value = item.SpecificRisk != null ? txtFormat.FormatwithNewLine(item.SpecificRisk, true) : string.Empty;
                            ws.Cells[row, 5].Value = item.FinStatemenElement != null ? txtFormat.FormatwithNewLine(item.FinStatemenElement, true) : string.Empty;
                            if (item.FinancialStatementAssert  != null)
                            {
                                ws.Cells[row, 6].Value = item.CompletenessAccuracy;
                                ws.Cells[row, 7].Value = item.ExistenceDisclosure;
                                ws.Cells[row, 8].Value = item.PresentationDisclosure;
                                ws.Cells[row, 9].Value = item.RigthsObligation;
                                ws.Cells[row, 10].Value = item.ValuationAlloc;

                            }
                            
                           

                            ws.Cells[row, 11].Value = item.ControlId;
                            ws.Cells[row, 12].Value = item.ControlActivityFy19 != null ? txtFormat.FormatwithNewLine(item.ControlActivityFy19, true) : string.Empty;
                            ws.Cells[row, 13].Value = item.ShortDescription != null ? txtFormat.FormatwithNewLine(item.ShortDescription, true) : string.Empty;
                            ws.Cells[row, 14].Value = item.ControlPlaceDate;
                            ws.Cells[row, 15].Value = item.ControlOwner;
                            ws.Cells[row, 16].Value = item.ControlFrequency;
                            ws.Cells[row, 17].Value = item.Key;
                            ws.Cells[row, 18].Value = item.ControlType;
                            ws.Cells[row, 19].Value = item.NatureProc;
                            ws.Cells[row, 20].Value = item.FraudControl;
                            ws.Cells[row, 21].Value = item.RiskLvl;
                            ws.Cells[row, 22].Value = item.ManagementRevControl;
                            ws.Cells[row, 23].Value = item.PbcList != null ? txtFormat.FormatwithNewLine(item.PbcList, true) : string.Empty;
                            ws.Cells[row, 24].Value = item.TestProc != null ? txtFormat.FormatwithNewLine(item.TestProc, true) : string.Empty;


                            xlsService.ExcelSetVerticalAlignCenter(ws, row, 1, row, 24);
                            xlsService.ExcelSetHorizontalAlignLeft(ws, row, 1, row, 24);
                            xlsService.ExcelSetFontSize(ws, row, 1, row, 24, 10);
                            xlsService.ExcelSetBorder(ws, row, 1, row, 24);
                            xlsService.ExcelWrapText(ws, row, 1, row, 24);
                            row++;
                        }


                        //save file
                        string startupPath = Environment.CurrentDirectory;
                        //string strSourceDownload = startupPath + "\\include\\sampleselection\\download\\";
                        string strSourceDownload = Path.Combine(startupPath, "include", "upload", "rcm");

                        if (!Directory.Exists(strSourceDownload))
                        {
                            Directory.CreateDirectory(strSourceDownload);
                        }
                        var ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                        string filename = $"Rcm-{clientName}-{ts}.xlsx";
                        string strOutput = Path.Combine(strSourceDownload, filename);

                        //Check if file not exists
                        if (System.IO.File.Exists(strOutput))
                        {
                            System.IO.File.Delete(strOutput);
                        }

                        xls.SaveAs(new FileInfo(strOutput));
                        excelFilename = filename;

                        //return status "Ok" with filename
                        return Ok(excelFilename);

                    }
                    #endregion
                
                }

                //return status "NoContent" 
                return NoContent();

            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GenerateRcmFile {ex}", "ErrorGenerateRcmFile");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GenerateRcmFile");
                if (_environment.IsDevelopment())
                {
                    return BadRequest(ex.ToString());
                }
                else
                {
                    return BadRequest();
                }
            }

            
            //return Ok(excelFilename);
        }


        //[AllowAnonymous]
        [HttpPost("keyreport1")]
        public IActionResult GetKeyReport1Control([FromBody] KeyReportFilter filter)
        {

            try
            {
                using (var context = _soxContext.Database.BeginTransaction())
                {
                    //search for RCM key control
                    var _rcm = _soxContext.Rcm
                       .Where(x => !x.Status.ToLower().Equals("closed")
                           && !x.Status.ToLower().Equals("completed")
                           && x.ClientName.Equals(filter.ClientName)
                           && x.FY.Equals(filter.FY)
                           && x.Key.ToLower().Equals("key"))
                       .FirstOrDefault();



                    if (_rcm != null && _rcm.PodioItemId != 0)
                    {
                        //RCM key control found
                        //we get workpaper or questionnaire per 

                        string consolOrigId = _config.GetSection("KeyReportApp").GetSection("ConsolOrigFormatId").Value;
                        //string allIUCId = _config.GetSection("KeyReportApp").GetSection("AllIUCId").Value;
                        //string testStatusTrackerId = _config.GetSection("KeyReportApp").GetSection("TestStatusTrackerId").Value;
                        //string exceptionId = _config.GetSection("KeyReportApp").GetSection("ExceptionId").Value;

                        var checkKeyReportQuestionnaire = _soxContext.KeyReportQuestion
                            .Where(x =>x.AppId.Equals(consolOrigId))
                            .Include(x => x.Options);

                        if(checkKeyReportQuestionnaire != null)
                        {
                            return Ok(checkKeyReportQuestionnaire.ToList());
                        }
                        else
                            return NoContent();

                    }
                    else
                    {
                        //if not found return no content
                        return NoContent();
                    }
                }


            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetKeyReport1Control {ex}", "ErrorGetKeyReport1Controll");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetKeyReport1Controll");
                if (_environment.IsDevelopment())
                {
                    return BadRequest(ex.ToString());
                }
                else
                {
                    return BadRequest();
                }
            }




        }


        //[AllowAnonymous]
        [HttpPost("keyreport2")]
        public IActionResult GetKeyReport2Control([FromBody] KeyReportFilter filter)
        {

            try
            {
                using (var context = _soxContext.Database.BeginTransaction())
                {
                    //search for RCM key control
                    var _rcm = _soxContext.Rcm
                       .Where(x => !x.Status.ToLower().Equals("closed")
                           && !x.Status.ToLower().Equals("completed")
                           && x.ClientName.Equals(filter.ClientName)
                           && x.FY.Equals(filter.FY)
                           && x.Key.ToLower().Equals("key"))
                       .FirstOrDefault();



                    if (_rcm != null && _rcm.PodioItemId != 0)
                    {
                        //RCM key control found
                        //we get workpaper or questionnaire per 

                        //string consolOrigId = _config.GetSection("KeyReportApp").GetSection("ConsolOrigFormatId").Value;
                        string allIUCId = _config.GetSection("KeyReportApp").GetSection("AllIUCId").Value;
                        //string testStatusTrackerId = _config.GetSection("KeyReportApp").GetSection("TestStatusTrackerId").Value;
                        //string exceptionId = _config.GetSection("KeyReportApp").GetSection("ExceptionId").Value;

                        var checkKeyReportQuestionnaire = _soxContext.KeyReportQuestion
                            .Where(x => x.AppId.Equals(allIUCId))
                            .Include(x => x.Options);

                        if (checkKeyReportQuestionnaire != null)
                        {
                            return Ok(checkKeyReportQuestionnaire.ToList());
                        }
                        else
                            return NoContent();

                    }
                    else
                    {
                        //if not found return no content
                        return NoContent();
                    }
                }


            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetKeyReport2Control {ex}", "ErrorGetKeyReport2Control");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetKeyReport2Control");
                if (_environment.IsDevelopment())
                {
                    return BadRequest(ex.ToString());
                }
                else
                {
                    return BadRequest();
                }
            }




        }


        //[AllowAnonymous]
        [HttpPost("keyreport3")]
        public IActionResult GetKeyReport3Control([FromBody] KeyReportFilter filter)
        {

            try
            {
                using (var context = _soxContext.Database.BeginTransaction())
                {
                    //search for RCM key control
                    var _rcm = _soxContext.Rcm
                       .Where(x => !x.Status.ToLower().Equals("closed")
                           && !x.Status.ToLower().Equals("completed")
                           && x.ClientName.Equals(filter.ClientName)
                           && x.FY.Equals(filter.FY)
                           && x.Key.ToLower().Equals("key"))
                       .FirstOrDefault();



                    if (_rcm != null && _rcm.PodioItemId != 0)
                    {
                        //RCM key control found
                        //we get workpaper or questionnaire per 

                        //string consolOrigId = _config.GetSection("KeyReportApp").GetSection("ConsolOrigFormatId").Value;
                        //string allIUCId = _config.GetSection("KeyReportApp").GetSection("AllIUCId").Value;
                        string testStatusTrackerId = _config.GetSection("KeyReportApp").GetSection("TestStatusTrackerId").Value;
                        //string exceptionId = _config.GetSection("KeyReportApp").GetSection("ExceptionId").Value;

                        var checkKeyReportQuestionnaire = _soxContext.KeyReportQuestion
                            .Where(x => x.AppId.Equals(testStatusTrackerId))
                            .Include(x => x.Options);

                        if (checkKeyReportQuestionnaire != null)
                        {
                            return Ok(checkKeyReportQuestionnaire.ToList());
                        }
                        else
                            return NoContent();

                    }
                    else
                    {
                        //if not found return no content
                        return NoContent();
                    }
                }


            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetKeyReport3Control {ex}", "ErrorGetKeyReport3Control");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetKeyReport3Control");
                if (_environment.IsDevelopment())
                {
                    return BadRequest(ex.ToString());
                }
                else
                {
                    return BadRequest();
                }
            }




        }


        //[AllowAnonymous]
        [HttpPost("keyreport4")]
        public IActionResult GetKeyReport4Control([FromBody] KeyReportFilter filter)
        {

            try
            {
                using (var context = _soxContext.Database.BeginTransaction())
                {
                    //search for RCM key control
                    var _rcm = _soxContext.Rcm
                       .Where(x => !x.Status.ToLower().Equals("closed")
                           && !x.Status.ToLower().Equals("completed")
                           && x.ClientName.Equals(filter.ClientName)
                           && x.FY.Equals(filter.FY)
                           && x.Key.ToLower().Equals("key"))
                       .FirstOrDefault();



                    if (_rcm != null && _rcm.PodioItemId != 0)
                    {
                        //RCM key control found
                        //we get workpaper or questionnaire per 

                        //string consolOrigId = _config.GetSection("KeyReportApp").GetSection("ConsolOrigFormatId").Value;
                        //string allIUCId = _config.GetSection("KeyReportApp").GetSection("AllIUCId").Value;
                        //string testStatusTrackerId = _config.GetSection("KeyReportApp").GetSection("TestStatusTrackerId").Value;
                        string exceptionId = _config.GetSection("KeyReportApp").GetSection("ExceptionId").Value;

                        var checkKeyReportQuestionnaire = _soxContext.KeyReportQuestion
                            .Where(x => x.AppId.Equals(exceptionId))
                            .Include(x => x.Options);

                        if (checkKeyReportQuestionnaire != null)
                        {
                            return Ok(checkKeyReportQuestionnaire.ToList());
                        }
                        else
                            return NoContent();

                    }
                    else
                    {
                        //if not found return no content
                        return NoContent();
                    }
                }


            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetKeyReport4Control {ex}", "ErrorGetKeyReport4Control");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetKeyReport4Control");
                if (_environment.IsDevelopment())
                {
                    return BadRequest(ex.ToString());
                }
                else
                {
                    return BadRequest();
                }
            }




        }

        //[AllowAnonymous]
        [HttpGet("view_rcm/{fy}/{ClientName}")]
        public IActionResult fetch_records_rcm(String fy, String ClientName)
        {

            //return "test";
            List<Rcm> _clients = new List<Rcm>();
            try
            {
                using (var context = _soxContext.Database.BeginTransaction())
                {
                    
                    var listClient = _soxContext.Rcm.Where(x =>
                        x.FY.Equals(fy)
                        && x.ClientName.Equals(ClientName)

                    )
                    .OrderBy(x => x.Process)
                    .ThenBy(x => x.ControlId);

                    if (listClient != null)
                    {
                        _clients = listClient.ToList();
                    }
                }
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error fetch_records_rcm {ex}", "Errorfetch_records_rcm");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "fetch_records_rcm");
            }

            if (_clients != null)
            {
                return Ok(_clients.ToArray());
            }
            else
            {
                return NoContent();
            }

        }
        
        //[AllowAnonymous]
        [HttpPost("rcm")]
        public IActionResult GetRcmDataAsync([FromBody] KeyReportFilter filter)
        {
            SoxTracker soxTracker = new SoxTracker();
            try
            {

                //Get all active FY
                var checkSoxTracker = _soxContext.SoxTracker
                   .Where(x =>
                       x.FY.ToLower().Equals(filter.FY.ToLower())
                       && x.ClientName.ToLower().Equals(filter.ClientName.ToLower())
                       && x.KeyReportName.ToLower().Equals(filter.KeyReportName.ToLower())
                       && x.ControlId.ToLower().Equals(filter.ControlId.ToLower())
                       && x.Status.ToLower().Equals("active")
                   ).AsNoTracking().FirstOrDefault();

                //Console.WriteLine(checkSoxTracker);

                if(checkSoxTracker != null)
                {
                    var checkRcm = _soxContext.Rcm
                    .Where(x =>
                        x.FY.ToLower().Equals(checkSoxTracker.FY.ToLower())
                        && x.ClientName.ToLower().Equals(checkSoxTracker.ClientName.ToLower())
                        && x.ControlId.ToLower().Equals(checkSoxTracker.ControlId.ToLower())
                        && x.Status.ToLower().Equals("active")
                    ).AsNoTracking()
                    .FirstOrDefault();

                    if (checkRcm != null)
                    {
                        return Ok(checkRcm);
                    }

                }

                return NoContent();

            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "GetRcmDataAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetRcmDataAsync");
                return BadRequest();
            }

        }
        
        [HttpPost("check-duplicates")]
        public async Task<IActionResult> checkDuplicates([FromBody] Rcm rcm)
        {
            RcmService rcmService = new RcmService(_soxContext, _config);
            SoxTrackerService soxTrackerService = new SoxTrackerService(_soxContext, _config);
            bool status = false;
            try
            {
                if (!await rcmService.CheckDuplicateRCMDB(rcm))
                    status = true;


                // status = true;
            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSaveToDatabase");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SaveToDatabase");
                if (_environment.IsDevelopment())
                {
                    return BadRequest(ex.ToString());
                }
                else
                {
                    return BadRequest();
                }

            }
            if (status)
            {
                return Ok();
            }
            else
            {
                return NoContent();
            }



        }



        //private string ReplaceTagHtmlParagraph(string source)
        //{
        //    string output = source.Replace("<p>", string.Empty);
        //    output = output.Replace("<b />", string.Empty);
        //    output = output.Replace("</p>", "\n\n");

        //    //remove other html tag
        //    output = Regex.Replace(output, "<.*?>", String.Empty);

        //    return output;
        //}

    }
}
