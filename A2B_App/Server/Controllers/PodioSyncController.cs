using A2B_App.Server.Data;

using A2B_App.Server.Services;
using A2B_App.Shared.Podio;
using A2B_App.Shared.Skype;
using A2B_App.Shared.Sox;
using A2B_App.Shared.Time;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Internal;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using PodioAPI;
using PodioAPI.Models;
using PodioAPI.Utils;
using PodioAPI.Utils.ItemFields;
using ShareFile.Api.Client.Requests.Filters;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Runtime.ExceptionServices;
using System.Threading.Tasks;

namespace A2B_App.Server.Controllers
{
    [Authorize]
    [ApiController]
    [Route("api/[controller]")]
    public class PodioSyncController : ControllerBase
    {

        private readonly IConfiguration _config;
        private readonly ILogger<PodioSyncController> _logger;
        private readonly IWebHostEnvironment _environment;
        private readonly SoxContext _soxContext;
        private readonly TimeContext _timeContext;
        private readonly UserContext _userContext;

        public PodioSyncController(
            IConfiguration config,
            ILogger<PodioSyncController> logger,
            IWebHostEnvironment environment,
            SoxContext soxContext,
            TimeContext timeContext,
            UserContext userContext
        )
        {
            _config = config;
            _logger = logger;
            _environment = environment;
            _soxContext = soxContext;
            _timeContext = timeContext;
            _userContext = userContext;
        }

        [AllowAnonymous]
        [HttpPost("employee")]
        public async Task<IActionResult> SyncEmployeeAsync([FromBody] SyncDateRange syncDateRange)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("TimePodioApp").GetSection("EmpAppId").Value;
                PodioAppKey.AppToken = _config.GetSection("TimePodioApp").GetSection("EmpAppToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);
                //bool isHookSave, isItemSave;
                //TimeService TimeService;
                FormatService formatService = new FormatService();


                if (podio.IsAuthenticated())
                {
                    var filter = new Dictionary<string, object>
                    {
                        {
                            "last_edit_on",
                            new  { from = syncDateRange.startDate, to = syncDateRange.endDate}
                        },
                    };

                    PodioCollection<Item> collectionCheck = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 1, 0, filters: filter, sortDesc: true);

                    if (collectionCheck.Filtered != 0)
                    {
                        int loop = (int)Math.Ceiling((double)collectionCheck.Filtered / 500);
                        int offset = 0;

                        #region Set FieldId 
                        int nameField = int.Parse(_config.GetSection("TimePodioApp").GetSection("EmpField").GetSection("Name").Value);
                        int orgField = int.Parse(_config.GetSection("TimePodioApp").GetSection("EmpField").GetSection("Organization").Value);
                        int statField = int.Parse(_config.GetSection("TimePodioApp").GetSection("EmpField").GetSection("Status").Value);
                        int contactField = int.Parse(_config.GetSection("TimePodioApp").GetSection("EmpField").GetSection("Contact").Value);
                        int userField = int.Parse(_config.GetSection("TimePodioApp").GetSection("EmpField").GetSection("UserId").Value);
                        int emailField = int.Parse(_config.GetSection("TimePodioApp").GetSection("EmpField").GetSection("Email").Value);
                        int skypeNameField = int.Parse(_config.GetSection("TimePodioApp").GetSection("EmpField").GetSection("SkypeName").Value);
                        int skypeObjField = int.Parse(_config.GetSection("TimePodioApp").GetSection("EmpField").GetSection("SkypeObj").Value);
                        int profileField = int.Parse(_config.GetSection("TimePodioApp").GetSection("EmpField").GetSection("ProfileId").Value);
                        #endregion


                        WriteLog writeLog = new WriteLog();

                        List<TeamMember> listTeamMember = new List<TeamMember>();

                        //get filtered items and stored in list
                        for (int i = 0; i < loop; i++)
                        {
                            PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 500, offset, filters: filter, sortDesc: true);
                            
                            //writeLog.Display(collection);
                            //Console.WriteLine($"----------------------------------------------------------------");
                            foreach (var item in collection.Items)
                            {
                                TeamMember teamMember = new TeamMember();
                                PodioRef podioRef = new PodioRef();

                                #region Podio Item
                                podioRef.ItemId = (int)item.ItemId;
                                podioRef.UniqueId = item.AppItemIdFormatted.ToString();
                                podioRef.Revision = item.CurrentRevision.Revision;
                                podioRef.Link = item.Link.ToString();
                                podioRef.CreatedBy = item.CreatedBy.Name.ToString();
                                podioRef.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());

                                #region podio fields

                                TextItemField txtMemberName = item.Field<TextItemField>(nameField);
                                teamMember.Name = txtMemberName.Value;

                                CategoryItemField catOrg = item.Field<CategoryItemField>(orgField);
                                IEnumerable<CategoryItemField.Option> listOrg = catOrg.Options;
                                teamMember.Organization = listOrg.Select(x => x.Text).FirstOrDefault();

                                CategoryItemField catStatus = item.Field<CategoryItemField>(statField);
                                IEnumerable<CategoryItemField.Option> listStatus = catStatus.Options;
                                teamMember.Status = listStatus.Select(x => x.Text).FirstOrDefault();

                                EmailItemField emailItemField = item.Field<EmailItemField>(emailField);
                                IEnumerable<EmailPhoneFieldResult> emails = emailItemField.Value;
                                if(emails != null && emails.Any())
                                {
                                    List<string> listEmail = new List<string>();
                                    foreach (var itemEmail in emails)
                                    {
                                        listEmail.Add(itemEmail.Value);       
                                    }

                                    teamMember.Email = string.Join(",", listEmail);

                                }

                                TextItemField txtUserId = item.Field<TextItemField>(userField);
                                teamMember.UserId = !string.IsNullOrEmpty(txtUserId.Value) ? int.Parse(formatService.ReplaceTagHtmlParagraph2(txtUserId.Value.ToString(), false)) : 0;

                                TextItemField txtSkypeName = item.Field<TextItemField>(skypeNameField);
                                teamMember.SkypeName = !string.IsNullOrEmpty(txtSkypeName.Value) ? formatService.ReplaceTagHtmlParagraph2(txtSkypeName.Value.ToString(), false) : string.Empty;

                                TextItemField txtSkypeObjRaw = item.Field<TextItemField>(skypeObjField);
                                teamMember.SkypeObjRaw = !string.IsNullOrEmpty(txtSkypeObjRaw.Value) ? formatService.ReplaceTagHtmlParagraph2(txtSkypeObjRaw.Value.ToString(), false) : string.Empty;


                                if(!string.IsNullOrEmpty(teamMember.SkypeObjRaw))
                                {
                                    if(TryParseJSONSkypeObj(teamMember.SkypeObjRaw, out SkypeObj skypeObj))
                                    {
                                        teamMember.SkypeObj = skypeObj;
                                        //var deserializeObj = JsonConvert.DeserializeObject<SkypeObj>(teamMember.SkypeObjRaw);
                                    }
                                    
                                }

                                ContactItemField contactItem = item.Field<ContactItemField>(contactField);
                                IEnumerable<Contact> contacts = contactItem.Contacts;

                                TeamMemberDetail teamMemberDetail = new TeamMemberDetail();
                                ProfileImage profileImage = new ProfileImage();

                                //check if contact is not null
                                if (contacts != null && contacts.Any())
                                {
                                    //get first item
                                    var firstItem = contacts.FirstOrDefault();
                                    teamMemberDetail.About = firstItem.About;
                                    teamMemberDetail.BirthDate = firstItem.BirthDate;
                                    teamMemberDetail.Address = firstItem.Address != null && firstItem.Address.Any() ? string.Join(", ",firstItem.Address) : string.Empty;
                                    teamMemberDetail.Zip = firstItem.Zip;
                                    teamMemberDetail.City = firstItem.City;
                                    teamMemberDetail.Country = firstItem.Country;
                                    teamMemberDetail.State = firstItem.State;
                                    teamMemberDetail.Mobile = firstItem.Phone != null && firstItem.Phone.Any() ? string.Join(", ", firstItem.Phone) : string.Empty;
                                    teamMemberDetail.JobTitle = firstItem.Title != null && firstItem.Title.Any() ? string.Join(", ", firstItem.Title) : string.Empty;
                                    teamMemberDetail.LinkedInURL = firstItem.LinkedIn;
                                    teamMemberDetail.Url = firstItem.Url != null && firstItem.Url.Any() ? firstItem.Url.ToString() : string.Empty;
                                    List<Skills> listSkills = new List<Skills>();
                                    if (firstItem.Skill != null && firstItem.Skill.Any())
                                    {
                                        foreach (var skillItem in firstItem.Skill)
                                        {
                                            Skills skills = new Skills();
                                            skills.Skill = skillItem;
                                            listSkills.Add(skills);
                                        }
                                    }
                                    profileImage.Link = firstItem.Image != null ? firstItem.Image.Link : string.Empty;
                                    teamMember.ProfileId = firstItem.ProfileId;

                                }

                                teamMember.ProfileImage = profileImage;
                                teamMember.PodioRef = podioRef;
                                teamMember.TeamMemberDetail = teamMemberDetail;
                                listTeamMember.Add(teamMember);
                                //writeLog.Display(teamMember);

                                #endregion

                                #endregion

                                //listTimeCode.Add(timeCode);
                            }

                            offset += 500;
                        }

                        if(listTeamMember != null && listTeamMember.Any())
                        {
                            foreach (var item in listTeamMember)
                            {
                                using (var context = _userContext.Database.BeginTransaction())
                                {
                                    var checkTeamMember = _userContext.TeamMember
                                        .Where(x => x.UserId.Equals(item.UserId))
                                        .Include(inner => inner.SkypeObj)
                                        .Include(img => img.ProfileImage)
                                        .Include(podio => podio.PodioRef)
                                        .Include(detail => detail.TeamMemberDetail)
                                        .AsNoTracking()
                                        .FirstOrDefault();
                                    if(checkTeamMember == null) //insert
                                    {
                                        _userContext.Add(item);
                                        await _userContext.SaveChangesAsync();
                                        context.Commit();
                                    }
                                    else //update
                                    {
                                        item.Id = checkTeamMember.Id;
                                        checkTeamMember.Name = item.Name;
                                        checkTeamMember.Email = item.Email;
                                        checkTeamMember.Organization = item.Organization;
                                        checkTeamMember.Status = item.Status;
                                        checkTeamMember.UserId = item.UserId;
                                        checkTeamMember.ProfileId = item.ProfileId;
                                        checkTeamMember.SkypeName = item.SkypeName;
                                        checkTeamMember.SkypeObjRaw = item.SkypeObjRaw;
                                        _userContext.Entry(checkTeamMember).State = EntityState.Modified;


                                        //skype address
                                        if(checkTeamMember.SkypeObj == null)
                                        {
                                            if(item.SkypeObj != null)
                                            {
                                                checkTeamMember.SkypeObj = item.SkypeObj;
                                                _userContext.Entry(checkTeamMember.SkypeObj).State = EntityState.Added;
                                            }     
                                        }
                                        else
                                        {
                                            item.SkypeObj.Sys_id = checkTeamMember.SkypeObj.Sys_id;
                                            checkTeamMember.SkypeObj = item.SkypeObj;
                                            _userContext.Entry(checkTeamMember.SkypeObj).State = EntityState.Modified;
                                        }


                                        //profile image
                                        if (checkTeamMember.ProfileImage == null)
                                        {
                                            if (item.ProfileImage != null)
                                            {
                                                checkTeamMember.ProfileImage = item.ProfileImage;
                                                _userContext.Entry(checkTeamMember.ProfileImage).State = EntityState.Added;
                                            }
                                        }
                                        else
                                        {
                                            item.ProfileImage.Id = checkTeamMember.ProfileImage.Id;
                                            checkTeamMember.ProfileImage = item.ProfileImage;
                                            _userContext.Entry(checkTeamMember.ProfileImage).State = EntityState.Modified;
                                        }


                                        //podio reference
                                        if (checkTeamMember.PodioRef == null)
                                        {
                                            if (item.PodioRef != null)
                                            {
                                                checkTeamMember.PodioRef = item.PodioRef;
                                                _userContext.Entry(checkTeamMember.PodioRef).State = EntityState.Added;
                                            }
                                        }
                                        else
                                        {
                                            item.PodioRef.Id = checkTeamMember.PodioRef.Id;
                                            checkTeamMember.PodioRef = item.PodioRef;
                                            _userContext.Entry(checkTeamMember.PodioRef).State = EntityState.Modified;
                                        }

                                        //team member detail
                                        if (checkTeamMember.TeamMemberDetail == null)
                                        {
                                            if (item.TeamMemberDetail != null)
                                            {
                                                checkTeamMember.TeamMemberDetail = item.TeamMemberDetail;
                                                _userContext.Entry(checkTeamMember.TeamMemberDetail).State = EntityState.Added;
                                            }
                                        }
                                        else
                                        {
                                            item.TeamMemberDetail.Id = checkTeamMember.TeamMemberDetail.Id;
                                            checkTeamMember.TeamMemberDetail = item.TeamMemberDetail;
                                            _userContext.Entry(checkTeamMember.TeamMemberDetail).State = EntityState.Modified;
                                        }

         

                                        await _userContext.SaveChangesAsync();
                                        context.Commit();
                                    }
                                }
                            }
                        }

                    }

                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncEmployeeAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncEmployeeAsync");
                return BadRequest(ex.ToString());
            }

            return Ok();

        }

        [AllowAnonymous]
        [HttpPost("mastertime")]
        public async Task<IActionResult> SyncMasterTimeAsync([FromBody] SyncDateRange syncDateRange)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("TimePodioApp").GetSection("TimeCodeAppId").Value;
                PodioAppKey.AppToken = _config.GetSection("TimePodioApp").GetSection("TimeCodeAppToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);
                //bool isHookSave, isItemSave;
                //TimeService TimeService;


                if (podio.IsAuthenticated())
                {
                    var filter = new Dictionary<string, object>
                    {
                        {
                            "last_edit_on",
                            new  { from = syncDateRange.startDate, to = syncDateRange.endDate}
                        },
                    };

                    //PodioCollection<Item> collectionCheck = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 1, 0, filters: filter, sortDesc: true);

                    //if (collectionCheck.Filtered != 0)
                    //{
                    //    int loop = (int)Math.Ceiling((double)collectionCheck.Filtered / 500);
                    //    int offset = 0;
                    //    List<TimeCode> listTimeCode = new List<TimeCode>();
                    //    List<TimeCode> listTimeCodeAdd = new List<TimeCode>();
                    //    List<TimeCode> listTimeCodeUpdate = new List<TimeCode>();

                    //    //get filtered items and stored in list
                    //    for (int i = 0; i < loop; i++)
                    //    {
                    //        PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 500, offset, filters: filter, sortDesc: true);

                    //        foreach (var item in collection.Items)
                    //        {
                    //            TimeCode timeCode = new TimeCode();

                    //            #region Podio Item
                    //            timeCode.PodioItemId = (int)item.ItemId;
                    //            timeCode.PodioUniqueId = item.AppItemIdFormatted.ToString();
                    //            timeCode.PodioRevision = item.CurrentRevision.Revision;
                    //            timeCode.PodioLink = item.Link.ToString();
                    //            timeCode.CreatedBy = item.CreatedBy.Name.ToString();
                    //            timeCode.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());

                    //            #region podio fields

                    //            //client reference field
                    //            AppItemField clientRef = item.Field<AppItemField>("client-relationship");
                    //            IEnumerable<Item> clientAppRef = clientRef.Items;
                    //            timeCode.ClientRef = clientAppRef.Select(x => x.Title).FirstOrDefault();
                    //            timeCode.ClientRefId = clientAppRef.Select(x => x.ItemId).FirstOrDefault();

                    //            //project reference field
                    //            AppItemField projectRef = item.Field<AppItemField>("project-code-relationship");
                    //            IEnumerable<Item> projectAppRef = projectRef.Items;
                    //            timeCode.ProjectRef = projectAppRef.Select(x => x.Title).FirstOrDefault();
                    //            timeCode.ProjectRefId = projectAppRef.Select(x => x.ItemId).FirstOrDefault();

                    //            //task reference field
                    //            AppItemField taskRef = item.Field<AppItemField>("task-code-relationship");
                    //            IEnumerable<Item> taskAppRef = taskRef.Items;
                    //            timeCode.TaskRef = taskAppRef.Select(x => x.Title).FirstOrDefault();
                    //            timeCode.TaskRefId = taskAppRef.Select(x => x.ItemId).FirstOrDefault();


                    //            //Risk field
                    //            CategoryItemField catStatus = item.Field<CategoryItemField>("status");
                    //            IEnumerable<CategoryItemField.Option> listcatRisk = catStatus.Options;
                    //            timeCode.Status = listcatRisk.Select(x => x.Text).FirstOrDefault();

                    //            #endregion

                    //            #endregion

                    //            listTimeCode.Add(timeCode);
                    //        }

                    //        offset += 500;
                    //    }

                    //    if (listTimeCode != null && listTimeCode.Count() > 0)
                    //    {
                    //        using (var context = _timeContext.Database.BeginTransaction())
                    //        {

                    //            //bool isExists = false;


                    //            foreach (var item in listTimeCode)
                    //            {
                    //                var timeCodeCheck = _timeContext.TimeCode.FirstOrDefault(id => id.PodioItemId == (int)item.PodioItemId);
                    //                if (timeCodeCheck != null)
                    //                {

                    //                    timeCodeCheck.ClientRef = item.ClientRef;
                    //                    timeCodeCheck.ClientRefId = item.ClientRefId;
                    //                    timeCodeCheck.ProjectRef = item.ProjectRef;
                    //                    timeCodeCheck.ProjectRefId = item.ProjectRefId;
                    //                    timeCodeCheck.TaskRef = item.TaskRef;
                    //                    timeCodeCheck.TaskRefId = item.TaskRefId;
                    //                    timeCodeCheck.Status = item.Status;
                    //                    timeCodeCheck.PodioRevision = item.PodioRevision;
                    //                    timeCodeCheck.PodioUniqueId = item.PodioUniqueId;

                    //                    var clientRefCheck = _timeContext.ClientReference.FirstOrDefault(id => id.PodioItemId == (int)item.ClientRefId);
                    //                    if (clientRefCheck != null)
                    //                    {
                    //                        timeCodeCheck.ClientCode = clientRefCheck.ClientCode;
                    //                    }

                    //                    listTimeCodeUpdate.Add(timeCodeCheck);
                    //                }
                    //                else
                    //                {
                    //                    listTimeCodeAdd.Add(item);
                    //                }
                    //            }

                    //            if (listTimeCodeAdd != null && listTimeCodeAdd.Count() > 0)
                    //            {
                    //                _timeContext.AddRange(listTimeCodeAdd);
                    //            }

                    //            if (listTimeCodeUpdate != null && listTimeCodeUpdate.Count() > 0)
                    //            {
                    //                _timeContext.UpdateRange(listTimeCodeUpdate);
                    //            }

                    //            await _timeContext.SaveChangesAsync();
                    //            context.Commit();

                    //        }
                    //    }

                    //}

                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncTimeCodeAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncTimeCodeAsync");
                return BadRequest(ex.ToString());
            }

            return Ok();

        }


        //Trigger Sync > Time > Time Code
        [HttpPost("podio/sync/time/timecode")]
        public async Task<IActionResult> SyncTimeCodeAsync([FromBody] SyncDateRange syncDateRange)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("TimePodioApp").GetSection("TimeCodeAppId").Value;
                PodioAppKey.AppToken = _config.GetSection("TimePodioApp").GetSection("TimeCodeAppToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);
                //bool isHookSave, isItemSave;
                //TimeService TimeService;


                if (podio.IsAuthenticated())
                {
                    var filter = new Dictionary<string, object>
                    {
                        {
                            "last_edit_on",
                            new  { from = syncDateRange.startDate, to = syncDateRange.endDate}
                        },
                    };

                    PodioCollection<Item> collectionCheck = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 1, 0, filters: filter, sortDesc: true);

                    if (collectionCheck.Filtered != 0)
                    {
                        int loop = (int)Math.Ceiling((double)collectionCheck.Filtered / 500);
                        int offset = 0;
                        List<TimeCode> listTimeCode = new List<TimeCode>();
                        List<TimeCode> listTimeCodeAdd = new List<TimeCode>();
                        List<TimeCode> listTimeCodeUpdate = new List<TimeCode>();

                        string fieldClientRef = _config.GetSection("TimePodioApp").GetSection("TimeCodeField").GetSection("ClientRef").Value;
                        string fieldProjectRef = _config.GetSection("TimePodioApp").GetSection("TimeCodeField").GetSection("ProjectRef").Value;
                        string fieldTaskRef = _config.GetSection("TimePodioApp").GetSection("TimeCodeField").GetSection("TaskRef").Value;
                        string fieldStatus = _config.GetSection("TimePodioApp").GetSection("TimeCodeField").GetSection("CatStatus").Value;

                        //get filtered items and stored in list
                        for (int i = 0; i < loop; i++)
                        {
                            PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 500, offset, filters: filter, sortDesc: true);

                            foreach (var item in collection.Items)
                            {
                                TimeCode timeCode = new TimeCode();
                                PodioRef podioRef = new PodioRef();
                                #region Podio Item
                                podioRef.ItemId = (int)item.ItemId;
                                podioRef.UniqueId = item.AppItemIdFormatted.ToString();
                                podioRef.Revision = item.CurrentRevision.Revision;
                                podioRef.Link = item.Link.ToString();
                                podioRef.CreatedBy = item.CreatedBy.Name.ToString();
                                podioRef.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());

                                #region podio fields

                                //client reference field
                                AppItemField clientRef = item.Field<AppItemField>(fieldClientRef);
                                IEnumerable<Item> clientAppRef = clientRef.Items;
                                timeCode.ClientRef = clientAppRef.Select(x => x.Title).FirstOrDefault();
                                timeCode.ClientRefId = clientAppRef.Select(x => x.ItemId).FirstOrDefault();

                                //project reference field
                                AppItemField projectRef = item.Field<AppItemField>(fieldProjectRef);
                                IEnumerable<Item> projectAppRef = projectRef.Items;
                                timeCode.ProjectRef = projectAppRef.Select(x => x.Title).FirstOrDefault();
                                timeCode.ProjectRefId = projectAppRef.Select(x => x.ItemId).FirstOrDefault();

                                //task reference field
                                AppItemField taskRef = item.Field<AppItemField>(fieldTaskRef);
                                IEnumerable<Item> taskAppRef = taskRef.Items;
                                timeCode.TaskRef = taskAppRef.Select(x => x.Title).FirstOrDefault();
                                timeCode.TaskRefId = taskAppRef.Select(x => x.ItemId).FirstOrDefault();


                                //Risk field
                                CategoryItemField catStatus = item.Field<CategoryItemField>(fieldStatus);
                                IEnumerable<CategoryItemField.Option> listcatRisk = catStatus.Options;
                                timeCode.Status = listcatRisk.Select(x => x.Text).FirstOrDefault();


                                timeCode.PodioRef = podioRef;
                                #endregion

                                #endregion

                                listTimeCode.Add(timeCode);
                            }

                            offset += 500;
                        }

                        if (listTimeCode != null && listTimeCode.Count() > 0)
                        {
                            using (var context = _timeContext.Database.BeginTransaction())
                            {

                                //bool isExists = false;


                                foreach (var item in listTimeCode)
                                {
                                    var timeCodeCheck = _timeContext.TimeCode.FirstOrDefault(id => id.PodioRef.ItemId == item.PodioRef.ItemId);
                                    if (timeCodeCheck != null)
                                    {

                                        timeCodeCheck.ClientRef = item.ClientRef;
                                        timeCodeCheck.ClientRefId = item.ClientRefId;
                                        timeCodeCheck.ProjectRef = item.ProjectRef;
                                        timeCodeCheck.ProjectRefId = item.ProjectRefId;
                                        timeCodeCheck.TaskRef = item.TaskRef;
                                        timeCodeCheck.TaskRefId = item.TaskRefId;
                                        timeCodeCheck.Status = item.Status;

                                        PodioRef podioRef = new PodioRef();
                                        podioRef = timeCodeCheck.PodioRef;
                                        podioRef.Revision = item.PodioRef.Revision;
                                        podioRef.UniqueId = item.PodioRef.UniqueId;
                                        timeCodeCheck.PodioRef = podioRef;
   

                                        var clientRefCheck = _timeContext.ClientReference.FirstOrDefault(id => id.PodioRef.ItemId == (int)item.ClientRefId);
                                        if (clientRefCheck != null)
                                        {
                                            timeCodeCheck.ClientCode = clientRefCheck.ClientCode;
                                        }

                                        listTimeCodeUpdate.Add(timeCodeCheck);
                                    }
                                    else
                                    {
                                        listTimeCodeAdd.Add(item);
                                    }
                                }

                                if (listTimeCodeAdd != null && listTimeCodeAdd.Count() > 0)
                                {
                                    _timeContext.AddRange(listTimeCodeAdd);
                                }

                                if (listTimeCodeUpdate != null && listTimeCodeUpdate.Count() > 0)
                                {
                                    _timeContext.UpdateRange(listTimeCodeUpdate);
                                }

                                await _timeContext.SaveChangesAsync();
                                context.Commit();

                            }
                        }

                    }

                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncTimeCodeAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncTimeCodeAsync");
                return BadRequest(ex.ToString());
            }

            return Ok();

        }


        //Trigger Sync > Time > Clients
        [HttpPost("podio/sync/time/clientreference")]
        public async Task<IActionResult> SyncClientReferenceAsync([FromBody] SyncDateRange syncDateRange)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("TimePodioApp").GetSection("ClientAppId").Value;
                PodioAppKey.AppToken = _config.GetSection("TimePodioApp").GetSection("ClientAppToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                //bool isHookSave, isItemSave;
                //A2B_App.Server.Services.TimeService TimeService;


                if (podio.IsAuthenticated())
                {
                    var filter = new Dictionary<string, object>
                    {
                        {
                            "last_edit_on",
                            new  { from = syncDateRange.startDate, to = syncDateRange.endDate}
                        },
                    };

                    PodioCollection<Item> collectionCheck = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 1, 0, filters: filter, sortDesc: true);

                    if (collectionCheck.Filtered != 0)
                    {
                        int loop = (int)Math.Ceiling((double)collectionCheck.Filtered / 500);
                        int offset = 0;
                        List<ClientReference> listCientRef = new List<ClientReference>();
                        List<ClientReference> listClientRefAdd = new List<ClientReference>();
                        List<ClientReference> listClientRefUpdate = new List<ClientReference>();

                        string fieldClientRef = _config.GetSection("TimePodioApp").GetSection("ClientField").GetSection("ClientRef").Value;
                        string fieldUniqueId = _config.GetSection("TimePodioApp").GetSection("ClientField").GetSection("UniqueId").Value;
                        string fieldHostDomain = _config.GetSection("TimePodioApp").GetSection("ClientField").GetSection("HostedDomain").Value;

                        //get filtered items and stored in list
                        for (int i = 0; i < loop; i++)
                        {
                            PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 500, offset, filters: filter, sortDesc: true);

                            foreach (var item in collection.Items)
                            {
                                ClientReference clientRef = new ClientReference();
                                PodioRef podioRef = new PodioRef();
                                #region Podio Item
                                podioRef.ItemId = (int)item.ItemId;
                                podioRef.UniqueId = item.AppItemIdFormatted.ToString();
                                podioRef.Revision = item.CurrentRevision.Revision;
                                podioRef.Link = item.Link.ToString();
                                podioRef.CreatedBy = item.CreatedBy.Name.ToString();
                                podioRef.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());

                                #region podio fields

                                //Client field
                                TextItemField txtClient = item.Field<TextItemField>(fieldClientRef);
                                clientRef.ClientRef = txtClient.Value;

                                //Client code field
                                TextItemField txtClientCode = item.Field<TextItemField>(fieldUniqueId);
                                clientRef.ClientCode = txtClientCode.Value;

                                //Client hosted domain
                                TextItemField txtHostedDomain = item.Field<TextItemField>(fieldHostDomain);
                                clientRef.ClientHostedDomain = txtHostedDomain.Value;

                                clientRef.PodioRef = podioRef;
                                #endregion

                                #endregion

                                listCientRef.Add(clientRef);
                            }

                            offset += 500;
                        }

                        if (listCientRef != null && listCientRef.Count() > 0)
                        {
                            using (var context = _timeContext.Database.BeginTransaction())
                            {
                                //bool isExists = false;

                                foreach (var item in listCientRef)
                                {
                                    var clientRefCheck = _timeContext.ClientReference.FirstOrDefault(id => id.PodioRef.ItemId == (int)item.PodioRef.ItemId);
                                    if (clientRefCheck != null)
                                    {
                                        clientRefCheck.ClientCode = item.ClientCode;
                                        clientRefCheck.ClientRef = item.ClientRef;
                                        clientRefCheck.ClientHostedDomain = item.ClientHostedDomain;

                                        PodioRef podioRef = new PodioRef();
                                        podioRef = clientRefCheck.PodioRef;
                                        podioRef.Revision = item.PodioRef.Revision;
                                        podioRef.UniqueId = item.PodioRef.UniqueId;
                                        clientRefCheck.PodioRef = podioRef;

                                        listClientRefUpdate.Add(clientRefCheck);
                                    }
                                    else
                                    {
                                        listClientRefAdd.Add(item);
                                    }
                                }

                                if (listClientRefAdd != null && listClientRefAdd.Count() > 0)
                                    _timeContext.AddRange(listClientRefAdd);

                                if (listClientRefUpdate != null && listClientRefUpdate.Count() > 0)
                                    _timeContext.UpdateRange(listClientRefUpdate);

                                await _timeContext.SaveChangesAsync();
                                context.Commit();
                            }
                        }

                    }

                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncClientReferenceAsync");
                return BadRequest(ex.ToString());
            }

            return Ok();

        }


        //Trigger Sync > Time > Project Reference
        [HttpPost("podio/sync/time/projectreference")]
        public async Task<IActionResult> SyncProjectReferenceAsync([FromBody] SyncDateRange syncDateRange)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("TimePodioApp").GetSection("ProjectAppId").Value;
                PodioAppKey.AppToken = _config.GetSection("TimePodioApp").GetSection("ProjectAppToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                //bool isHookSave, isItemSave;
                //TimeService TimeService;


                if (podio.IsAuthenticated())
                {
                    var filter = new Dictionary<string, object>
                    {
                        {
                            "last_edit_on",
                            new  { from = syncDateRange.startDate, to = syncDateRange.endDate}
                        },
                    };

                    PodioCollection<Item> collectionCheck = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 1, 0, filters: filter, sortDesc: true);

                    if (collectionCheck.Filtered != 0)
                    {
                        int loop = (int)Math.Ceiling((double)collectionCheck.Filtered / 500);
                        int offset = 0;
                        List<ProjectReference> listProjectRef = new List<ProjectReference>();
                        List<ProjectReference> listProjectRefAdd = new List<ProjectReference>();
                        List<ProjectReference> listProjectRefUpdate = new List<ProjectReference>();

                        string fieldProjectRef = _config.GetSection("TimePodioApp").GetSection("ProjectField").GetSection("ProjectRef").Value;
                        //get filtered items and stored in list
                        for (int i = 0; i < loop; i++)
                        {
                            PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 500, offset, filters: filter, sortDesc: true);

                            foreach (var item in collection.Items)
                            {
                                ProjectReference projectReference = new ProjectReference();
                                PodioRef podioRef = new PodioRef();

                                #region Podio Item
                                podioRef.ItemId = (int)item.ItemId;
                                podioRef.UniqueId = item.AppItemIdFormatted.ToString();
                                podioRef.Revision = item.CurrentRevision.Revision;
                                podioRef.Link = item.Link.ToString();
                                podioRef.CreatedBy = item.CreatedBy.Name.ToString();
                                podioRef.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());

                                #region podio fields

                                //Task field
                                TextItemField txtProject = item.Field<TextItemField>(fieldProjectRef);
                                projectReference.ProjectRef = txtProject.Value;

                                #endregion

                                #endregion

                                listProjectRef.Add(projectReference);
                            }

                            offset += 500;
                        }

                        if (listProjectRef != null && listProjectRef.Count() > 0)
                        {
                            using (var context = _timeContext.Database.BeginTransaction())
                            {

                                //bool isExists = false;


                                foreach (var item in listProjectRef)
                                {
                                    var projectRefCheck = _timeContext.ProjectReference.FirstOrDefault(id => id.PodioRef.ItemId == (int)item.PodioRef.ItemId);
                                    if (projectRefCheck != null)
                                    {
                                        projectRefCheck.ProjectRef = item.ProjectRef;

                                        PodioRef podioRef = new PodioRef();
                                        podioRef = projectRefCheck.PodioRef;
                                        podioRef.Revision = item.PodioRef.Revision;
                                        podioRef.UniqueId = item.PodioRef.UniqueId;
                                        projectRefCheck.PodioRef = podioRef;

                                        listProjectRefUpdate.Add(projectRefCheck);
                                    }
                                    else
                                    {
                                        listProjectRefAdd.Add(item);
                                    }
                                }

                                if (listProjectRefAdd != null && listProjectRefAdd.Count() > 0)
                                    _timeContext.AddRange(listProjectRefAdd);

                                if (listProjectRefUpdate != null && listProjectRefUpdate.Count() > 0)
                                    _timeContext.UpdateRange(listProjectRefUpdate);

                                await _timeContext.SaveChangesAsync();
                                context.Commit();
                            }
                        }

                    }

                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncProjectReferenceAsync");
                return BadRequest(ex.ToString());
            }

            return Ok();

        }


        //Trigger Sync > Time > Task Reference
        [HttpPost("podio/sync/time/taskreference")]
        public async Task<IActionResult> SyncTaskReferenceAsync([FromBody] SyncDateRange syncDateRange)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("TimePodioApp").GetSection("TaskAppId").Value;
                PodioAppKey.AppToken = _config.GetSection("TimePodioApp").GetSection("TaskAppToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                if (podio.IsAuthenticated())
                {
                    var filter = new Dictionary<string, object>
                    {
                        {
                            "last_edit_on",
                            new  { from = syncDateRange.startDate, to = syncDateRange.endDate}
                        },
                    };

                    PodioCollection<Item> collectionCheck = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 1, 0, filters: filter, sortDesc: true);

                    if (collectionCheck.Filtered != 0)
                    {
                        int loop = (int)Math.Ceiling((double)collectionCheck.Filtered / 500);
                        int offset = 0;
                        List<TaskReference> listTimeCode = new List<TaskReference>();
                        List<TaskReference> listTimeCodeAdd = new List<TaskReference>();
                        List<TaskReference> listTimeCodeUpdate = new List<TaskReference>();

                        string fieldTaskRef = _config.GetSection("TimePodioApp").GetSection("TaskField").GetSection("TaskRef").Value;

                        //get filtered items and stored in list
                        for (int i = 0; i < loop; i++)
                        {
                            PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 500, offset, filters: filter, sortDesc: true);

                            foreach (var item in collection.Items)
                            {
                                TaskReference taskReference = new TaskReference();
                                PodioRef podioRef = new PodioRef();

                                #region Podio Item
                                podioRef.ItemId = (int)item.ItemId;
                                podioRef.UniqueId = item.AppItemIdFormatted.ToString();
                                podioRef.Revision = item.CurrentRevision.Revision;
                                podioRef.Link = item.Link.ToString();
                                podioRef.CreatedBy = item.CreatedBy.Name.ToString();
                                podioRef.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());

                                #region podio fields


                                //Task field
                                TextItemField txtTask = item.Field<TextItemField>(fieldTaskRef);
                                taskReference.TaskRef = txtTask.Value;

                                #endregion

                                #endregion

                                listTimeCode.Add(taskReference);
                            }

                            offset += 500;
                        }

                        if (listTimeCode != null && listTimeCode.Count() > 0)
                        {
                            using (var context = _timeContext.Database.BeginTransaction())
                            {

                                //bool isExists = false;


                                foreach (var item in listTimeCode)
                                {
                                    var taskRefCheck = _timeContext.TaskReference.FirstOrDefault(id => id.PodioRef.ItemId == (int)item.PodioRef.ItemId);
                                    if (taskRefCheck != null)
                                    {
                                        taskRefCheck.TaskRef = item.TaskRef;

                                        PodioRef podioRef = new PodioRef();
                                        podioRef = taskRefCheck.PodioRef;
                                        podioRef.Revision = item.PodioRef.Revision;
                                        podioRef.UniqueId = item.PodioRef.UniqueId;
                                        taskRefCheck.PodioRef = podioRef;

                                        listTimeCodeUpdate.Add(taskRefCheck);
                                    }
                                    else
                                    {
                                        listTimeCodeAdd.Add(item);
                                    }
                                }

                                if (listTimeCodeAdd != null && listTimeCodeAdd.Count() > 0)
                                    _timeContext.AddRange(listTimeCodeAdd);

                                if (listTimeCodeUpdate != null && listTimeCodeUpdate.Count() > 0)
                                    _timeContext.UpdateRange(listTimeCodeUpdate);

                                await _timeContext.SaveChangesAsync();
                                context.Commit();
                            }
                        }

                    }

                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncTaskReferenceAsync");
                return BadRequest(ex.ToString());
            }

            return Ok();

        }


        //Trigger Sync > Sample Selection > Client
        [HttpPost("podio/sync/sampleselection/client")]
        public async Task<IActionResult> SyncSampleSelectionClientAsync([FromBody] SyncDateRange syncDateRange)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("SampleSelectionPodioApp").GetSection("ClientAppId").Value;
                PodioAppKey.AppToken = _config.GetSection("SampleSelectionPodioApp").GetSection("ClientAppToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                if (podio.IsAuthenticated())
                {
                    var filter = new Dictionary<string, object>
                    {
                        {
                            "last_edit_on",
                            new  { from = syncDateRange.startDate, to = syncDateRange.endDate}
                        },
                    };

                    PodioCollection<Item> collectionCheck = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 1, 0, filters: filter, sortDesc: true);

                    if (collectionCheck.Filtered != 0)
                    {
                        int loop = (int)Math.Ceiling((double)collectionCheck.Filtered / 500);
                        int offset = 0;
                        List<ClientSs> listClientSs = new List<ClientSs>();
                        List<ClientSs> listClientSsAdd = new List<ClientSs>();
                        List<ClientSs> listClientSsUpdate = new List<ClientSs>();

                        //get filtered items and stored in list
                        for (int i = 0; i < loop; i++)
                        {
                            PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 500, offset, filters: filter, sortDesc: true);

                            foreach (var item in collection.Items)
                            {
                                ClientSs clientSs = new ClientSs();

                                #region Podio Item
                                clientSs.PodioItemId = (int)item.ItemId;
                                clientSs.PodioUniqueId = item.AppItemIdFormatted.ToString();
                                clientSs.PodioRevision = item.CurrentRevision.Revision;
                                clientSs.PodioLink = item.Link.ToString();
                                clientSs.CreatedBy = item.CreatedBy.Name.ToString();
                                clientSs.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());
                                clientSs.ItemId = (int)item.ItemId;

                                #region Podio fields

                                //name
                                TextItemField txtTitle = item.Field<TextItemField>("title");
                                clientSs.Name = txtTitle.Value;


                                //client reference field
                                AppItemField clientRef = item.Field<AppItemField>("client-relationship");
                                IEnumerable<Item> clientAppRef = clientRef.Items;
                                clientSs.ClientName = clientAppRef.Select(x => x.Title).FirstOrDefault();
                                //clientSs.ItemId = clientAppRef.Select(x => x.ItemId).FirstOrDefault();
                                clientSs.ClientItemId = clientAppRef.Select(x => x.ItemId).FirstOrDefault();

                                //Client Id
                                CalculationItemField calClientId = item.Field<CalculationItemField>("client-id");
                                clientSs.ClientCode = calClientId.ValueAsString;

                                //external auditor field
                                TextItemField calcExtAuditor = item.Field<TextItemField>("external-auditor");
                                clientSs.ExternalAuditor = calcExtAuditor.Value;

                                //Round 1 Percent
                                NumericItemField numRound1Percent = item.Field<NumericItemField>("round-1-percent");
                                clientSs.Percent = numRound1Percent.Value != null ? (int)numRound1Percent.Value : 0;

                                //Round 1 Percent
                                NumericItemField numRound2Percent = item.Field<NumericItemField>("round-2-percent");
                                clientSs.PercentRound2 = numRound2Percent.Value != null ? (int)numRound2Percent.Value : 0;

                                //Sharefile ID (Save File)
                                TextItemField calcSharefileId = item.Field<TextItemField>("sharefile-id-save-file");
                                clientSs.SharefileId = calcSharefileId.Value;

                                //SF Key Report ID (Screenshot)
                                TextItemField txtSharefileSsId = item.Field<TextItemField>("sf-key-report-id");
                                clientSs.SharefileScreenshotId = txtSharefileSsId.Value;

                                //SF Key Report ID (Report)
                                TextItemField txtSharefileReportId = item.Field<TextItemField>("sf-key-report-id-report");
                                clientSs.SharefileReportId = txtSharefileReportId.Value;

                                #endregion

                                #endregion

                                listClientSs.Add(clientSs);
                            }

                            offset += 500;
                        }

                        if (listClientSs != null && listClientSs.Count() > 0)
                        {
                            using (var context = _soxContext.Database.BeginTransaction())
                            {

                                foreach (var item in listClientSs)
                                {
                                    var clientSsCheck = _soxContext.ClientSs.FirstOrDefault(id => id.PodioItemId == (int)item.PodioItemId);
                                    if (clientSsCheck != null)
                                    {
                                        //clientSsCheck.Id = 0; // for error testing
                                        clientSsCheck.PodioRevision = item.PodioRevision;
                                        clientSsCheck.PodioUniqueId = item.PodioUniqueId;
                                        clientSsCheck.ClientName = item.ClientName;
                                        clientSsCheck.ExternalAuditor = item.ExternalAuditor;
                                        clientSsCheck.ItemId = item.ItemId;
                                        clientSsCheck.ClientItemId = item.ClientItemId;
                                        clientSsCheck.SharefileId = item.SharefileId;
                                        clientSsCheck.Percent = item.Percent;
                                        clientSsCheck.PercentRound2 = item.PercentRound2;
                                        clientSsCheck.SharefileScreenshotId = item.SharefileScreenshotId;
                                        clientSsCheck.Name = item.Name;
                                        clientSsCheck.SharefileReportId = item.SharefileReportId;
                                        clientSsCheck.ClientCode = item.ClientCode;
                                        listClientSsUpdate.Add(clientSsCheck);
                                    }
                                    else
                                    {
                                        listClientSsAdd.Add(item);
                                    }
                                }

                                if (listClientSsAdd != null && listClientSsAdd.Count() > 0)
                                    _soxContext.AddRange(listClientSsAdd);

                                if (listClientSsUpdate != null && listClientSsUpdate.Count() > 0)
                                    _soxContext.UpdateRange(listClientSsUpdate);

                                await _soxContext.SaveChangesAsync();
                                context.Commit();
                            }
                        }

                    }

                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncSampleSelectionClientAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncSampleSelectionClientAsync");
                return BadRequest(ex.ToString());
            }

            return Ok();

        }


        //Trigger Sync > Sample Selection > Client
        [HttpPost("podio/sync/sampleselection/matrix")]
        public async Task<IActionResult> SyncSampleSelectionMatrixAsync([FromBody] SyncDateRange syncDateRange)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("SampleSelectionPodioApp").GetSection("MatrixAppId").Value;
                PodioAppKey.AppToken = _config.GetSection("SampleSelectionPodioApp").GetSection("MatrixAppToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                if (podio.IsAuthenticated())
                {
                    var filter = new Dictionary<string, object>
                    {
                        {
                            "last_edit_on",
                            new  { from = syncDateRange.startDate, to = syncDateRange.endDate}
                        },
                    };

                    PodioCollection<Item> collectionCheck = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 1, 0, filters: filter, sortDesc: true);

                    if (collectionCheck.Filtered != 0)
                    {
                        int loop = (int)Math.Ceiling((double)collectionCheck.Filtered / 500);
                        int offset = 0;
                        List<Matrix> listMatrix = new List<Matrix>();
                        List<Matrix> listMatrixAdd = new List<Matrix>();
                        List<Matrix> listMatrixUpdate = new List<Matrix>();

                        //get filtered items and stored in list
                        for (int i = 0; i < loop; i++)
                        {
                            PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 500, offset, filters: filter, sortDesc: true);

                            foreach (var item in collection.Items)
                            {
                                Matrix matrix = new Matrix();

                                #region Podio Item
                                matrix.PodioItemId = (int)item.ItemId;
                                matrix.PodioUniqueId = item.AppItemIdFormatted.ToString();
                                matrix.PodioRevision = item.CurrentRevision.Revision;
                                matrix.PodioLink = item.Link.ToString();
                                matrix.CreatedBy = item.CreatedBy.Name.ToString();
                                matrix.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());

                                #region Podio fields

                                //client reference field
                                AppItemField clientRef = item.Field<AppItemField>("client");
                                IEnumerable<Item> clientAppRef = clientRef.Items;
                                matrix.ClientName = clientAppRef.Select(x => x.Title).FirstOrDefault();
                                matrix.ClientItemId = clientAppRef.Select(x => x.ItemId).FirstOrDefault();

                                var clientCheck = _soxContext.ClientSs.FirstOrDefault(id => id.PodioItemId == matrix.ClientItemId);
                                if (clientCheck != null)
                                {
                                    matrix.ClientCode = clientCheck.ClientCode;
                                }

                                //External Auditor
                                CalculationItemField calcExternalAuditor = item.Field<CalculationItemField>("external-auditor");
                                matrix.ExternalAuditor = calcExternalAuditor.ValueAsString;

                                //Frequency
                                CategoryItemField catFrequency = item.Field<CategoryItemField>("frequency");
                                IEnumerable<CategoryItemField.Option> listCatFrequency = catFrequency.Options;
                                matrix.Frequency = listCatFrequency.Select(x => x.Text).FirstOrDefault();

                                //Risk
                                CategoryItemField catRisk = item.Field<CategoryItemField>("risk");
                                IEnumerable<CategoryItemField.Option> listCatRisk = catRisk.Options;
                                matrix.Risk = listCatRisk.Select(x => x.Text).FirstOrDefault();

                                //Value
                                NumericItemField numValue = item.Field<NumericItemField>("value");
                                matrix.SizeValue = numValue.Value != null ? (int)numValue.Value : 0;

                                //Population Start
                                NumericItemField numPopStart = item.Field<NumericItemField>("population-start");
                                matrix.StartPopulation = numPopStart.Value != null ? (int)numPopStart.Value : 0;

                                //Population End
                                NumericItemField numPopEnd = item.Field<NumericItemField>("population-end");
                                matrix.EndPopulation = numPopEnd.Value != null ? (int)numPopEnd.Value : 0;

                                #endregion

                                #endregion

                                listMatrix.Add(matrix);
                            }

                            offset += 500;
                        }

                        if (listMatrix != null && listMatrix.Count() > 0)
                        {
                            using (var context = _soxContext.Database.BeginTransaction())
                            {

                                foreach (var item in listMatrix)
                                {
                                    var matrixCheck = _soxContext.Matrix.FirstOrDefault(id => id.PodioItemId == (int)item.PodioItemId);
                                    if (matrixCheck != null)
                                    {
                                        matrixCheck.PodioRevision = item.PodioRevision;
                                        matrixCheck.PodioUniqueId = item.PodioUniqueId;

                                        matrixCheck.ClientName = item.ClientName;
                                        matrixCheck.ClientCode = item.ClientCode;
                                        matrixCheck.ClientItemId = item.ClientItemId;
                                        matrixCheck.ExternalAuditor = item.ExternalAuditor;
                                        matrixCheck.Frequency = item.Frequency;
                                        matrixCheck.Risk = item.Risk;
                                        matrixCheck.SizeValue = item.SizeValue;
                                        matrixCheck.StartPopulation = item.StartPopulation;
                                        matrixCheck.EndPopulation = item.EndPopulation;
                                        listMatrixUpdate.Add(matrixCheck);
                                    }
                                    else
                                    {
                                        listMatrixAdd.Add(item);
                                    }
                                }

                                if (listMatrixAdd != null && listMatrixAdd.Count() > 0)
                                    _soxContext.AddRange(listMatrixAdd);

                                if (listMatrixUpdate != null && listMatrixUpdate.Count() > 0)
                                    _soxContext.UpdateRange(listMatrixUpdate);

                                await _soxContext.SaveChangesAsync();
                                context.Commit();
                            }
                        }

                    }

                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncTaskReferenceAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncTaskReferenceAsync");
                return BadRequest(ex.ToString());
            }

            return Ok();

        }


        [HttpPost("podio/sync/rcm/cta")]
        public async Task<IActionResult> SyncRcmCtaAsync([FromBody] SyncDateRange syncDateRange)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("RcmPodioApp").GetSection("CtaId").Value;
                PodioAppKey.AppToken = _config.GetSection("RcmPodioApp").GetSection("CtaToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                if (podio.IsAuthenticated())
                {

                    var filter = new Dictionary<string, object>
                    {
                        {
                            "last_edit_on",
                            new  { from = syncDateRange.startDate, to = syncDateRange.endDate}
                        },
                    };

                    PodioCollection<Item> collectionCheck = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 1, 0, filters: filter, sortDesc: true);

                    if (collectionCheck.Filtered != 0)
                    {
                        int loop = (int)Math.Ceiling((double)collectionCheck.Filtered / 500);
                        int offset = 0;
                        List<RcmCta> listMatrix = new List<RcmCta>();
                        List<RcmCta> listMatrixAdd = new List<RcmCta>();
                        List<RcmCta> listMatrixUpdate = new List<RcmCta>();

                        //get filtered items and stored in list
                        for (int i = 0; i < loop; i++)
                        {
                            PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 500, offset, filters: filter, sortDesc: true);

                            foreach (var item in collection.Items)
                            {
                                RcmCta rcmCta = new RcmCta();

                                #region Podio Item
                                rcmCta.PodioItemId = (int)item.ItemId;
                                rcmCta.PodioUniqueId = item.AppItemIdFormatted.ToString();
                                rcmCta.PodioRevision = item.CurrentRevision.Revision;
                                rcmCta.PodioLink = item.Link.ToString();
                                rcmCta.CreatedBy = item.CreatedBy.Name.ToString();
                                rcmCta.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());

                                #region sample selection fields

                                //1. Client Name
                                AppItemField clientApp = item.Field<AppItemField>("1-client-name");
                                IEnumerable<Item> clientAppRef = clientApp.Items;
                                rcmCta.ClientName = clientAppRef.Select(x => x.Title).FirstOrDefault();

                                //1. Client Code
                                CalculationItemField calcClientCode = item.Field<CalculationItemField>("1-client-code");
                                rcmCta.ClientCode = calcClientCode.ValueAsString;


                                //2. What is the sub-process?
                                CategoryItemField catSubProcess = item.Field<CategoryItemField>("what-is-the-sub-process");
                                IEnumerable<CategoryItemField.Option> listSubProcess = catSubProcess.Options;
                                rcmCta.SubProcess = listSubProcess.Select(x => x.Text).FirstOrDefault();

                                //3. What is the financial statement element?
                                TextItemField txtFinStatement = item.Field<TextItemField>("what-is-the-financial-statement-element-2");
                                rcmCta.FinancialStatementElement = txtFinStatement.Value;

                                //4. What is the specific risk?
                                TextItemField txtSpecificRisk = item.Field<TextItemField>("what-is-the-specific-risk");
                                rcmCta.SpecificRisk = txtSpecificRisk.Value;

                                //5.a. Does it show completeness & accuracy?
                                CategoryItemField q5a = item.Field<CategoryItemField>("5a-does-it-show-completeness-accuracy");
                                IEnumerable<CategoryItemField.Option> listQ5A = q5a.Options;
                                rcmCta.Q5ACompletenessAccuracy = listQ5A.Select(x => x.Text).FirstOrDefault();
                                

                                //5.b. Does it show existence & occurrence?
                                CategoryItemField q5b = item.Field<CategoryItemField>("5b-does-it-show-existence-occurrence");
                                IEnumerable<CategoryItemField.Option> listQ5B = q5b.Options;
                                rcmCta.Q5BExistenceOccur = listQ5B.Select(x => x.Text).FirstOrDefault();

                                //5.c. Does it show presentation & disclosure?
                                CategoryItemField q5c = item.Field<CategoryItemField>("5c-does-it-show-presentation-disclosure");
                                IEnumerable<CategoryItemField.Option> listQ5C = q5c.Options;
                                rcmCta.Q5CPresentationDisclose = listQ5C.Select(x => x.Text).FirstOrDefault();

                                //5.d. Does it show rights & obligations?
                                CategoryItemField q5d = item.Field<CategoryItemField>("5d-does-it-show-rights-obligations");
                                IEnumerable<CategoryItemField.Option> listQ5D = q5d.Options;
                                rcmCta.Q5DRightObligation = listQ5D.Select(x => x.Text).FirstOrDefault();

                                //5.e. Does it show valuation & allocation?
                                CategoryItemField q5e = item.Field<CategoryItemField>("5e-does-it-show-valuation-allocation");
                                IEnumerable<CategoryItemField.Option> listQ5E = q5e.Options;
                                rcmCta.Q5EValuationAlloc = listQ5E.Select(x => x.Text).FirstOrDefault();

                                //6. What is the control ID?
                                CategoryItemField controlId = item.Field<CategoryItemField>("6-what-is-the-control-id-3");
                                IEnumerable<CategoryItemField.Option> listControlId = controlId.Options;
                                rcmCta.ControlId = listControlId.Select(x => x.Text).FirstOrDefault();

                                //7. What is the control activity?
                                TextItemField txtControlActivity = item.Field<TextItemField>("what-is-the-control-activity");
                                rcmCta.ControlActivity = txtControlActivity.Value;

                                //8. What is the control objective?
                                TextItemField txtControlObj = item.Field<TextItemField>("8-what-is-the-control-objective");
                                rcmCta.ControlObjective = txtControlObj.Value;

                                //9. When is the control in place date?
                                TextItemField txtControlPlace = item.Field<TextItemField>("when-is-the-control-in-place-2");
                                rcmCta.ControlPlaceDate = txtControlPlace.Value;

                                //10. Who is the control owner?
                                CategoryItemField controlOwner = item.Field<CategoryItemField>("who-is-the-control-owner");
                                IEnumerable<CategoryItemField.Option> listControlOwner = controlOwner.Options;
                                rcmCta.ControlOwner = listControlOwner.Select(x => x.Text).FirstOrDefault();

                                //11. What is the control frequency?
                                CategoryItemField frequency = item.Field<CategoryItemField>("what-is-the-control-frequency");
                                IEnumerable<CategoryItemField.Option> listFrequency = frequency.Options;
                                rcmCta.Frequency = listFrequency.Select(x => x.Text).FirstOrDefault();

                                //12. What is the entity?
                                CategoryItemField entity = item.Field<CategoryItemField>("what-is-the-entity-2");
                                IEnumerable<CategoryItemField.Option> listEntity = entity.Options;
                                rcmCta.Entity = listEntity.Select(x => x.Text).FirstOrDefault();

                                //12.a. Notes
                                TextItemField txtNotes = item.Field<TextItemField>("notes");
                                rcmCta.Notes = txtNotes.Value;

                                //13. Is it a Key/ Non-Key control?
                                CategoryItemField keyControl = item.Field<CategoryItemField>("is-it-a-key-non-key-control");
                                IEnumerable<CategoryItemField.Option> listKeyControl = keyControl.Options;
                                rcmCta.KeyControl = listKeyControl.Select(x => x.Text).FirstOrDefault();

                                //14. What is the control type?
                                CategoryItemField controlType = item.Field<CategoryItemField>("what-is-the-control-type");
                                IEnumerable<CategoryItemField.Option> listControlType = controlType.Options;
                                rcmCta.ControlType = listControlType.Select(x => x.Text).FirstOrDefault();

                                //15. What is the nature of procedure?
                                CategoryItemField natureProc = item.Field<CategoryItemField>("what-is-the-nature-of-procedure");
                                IEnumerable<CategoryItemField.Option> listNatureProc = natureProc.Options;
                                rcmCta.NatureProcedure = listNatureProc.Select(x => x.Text).FirstOrDefault();

                                //16. Is it a fraud control?
                                CategoryItemField fraudControl = item.Field<CategoryItemField>("fraud-control");
                                IEnumerable<CategoryItemField.Option> listFraudControl = fraudControl.Options;
                                rcmCta.FraudControl = listFraudControl.Select(x => x.Text).FirstOrDefault();

                                //17. What is the risk level?
                                CategoryItemField risk = item.Field<CategoryItemField>("risk-level");
                                IEnumerable<CategoryItemField.Option> listRisk = risk.Options;
                                rcmCta.Risk = listRisk.Select(x => x.Text).FirstOrDefault();

                                //18. Is it a management review control?
                                CategoryItemField revControl = item.Field<CategoryItemField>("management-review-control");
                                IEnumerable<CategoryItemField.Option> listRevControl = revControl.Options;
                                rcmCta.ReviewControl = listEntity.Select(x => x.Text).FirstOrDefault();

                                //19. What is the Testing Procedure?
                                TextItemField txtTestProc = item.Field<TextItemField>("test-procedures");
                                rcmCta.TestingProcedure = txtTestProc.Value;

                                //Population File Required
                                CategoryItemField populationFileReq = item.Field<CategoryItemField>("populationfilerequired");
                                IEnumerable<CategoryItemField.Option> listpopulationFileReq = populationFileReq.Options;
                                rcmCta.PopulationFileRequired = listpopulationFileReq.Select(x => x.Text).FirstOrDefault();


                                #endregion

                                #endregion

                                listMatrix.Add(rcmCta);
                            }

                            offset += 500;
                        }

                        if (listMatrix != null && listMatrix.Count() > 0)
                        {
                            using (var context = _soxContext.Database.BeginTransaction())
                            {

                                foreach (var item in listMatrix)
                                {
                                    var rcmCheck = _soxContext.RcmCta.FirstOrDefault(id => id.PodioItemId == (int)item.PodioItemId);
                                    if (rcmCheck != null)
                                    {
                                        rcmCheck.ClientName = item.ClientName;
                                        rcmCheck.ClientCode = item.ClientCode;
                                        rcmCheck.SubProcess = item.SubProcess;
                                        rcmCheck.FinancialStatementElement = item.FinancialStatementElement;
                                        rcmCheck.SpecificRisk = item.SpecificRisk;
                                        
                                        rcmCheck.Q5ACompletenessAccuracy = item.Q5ACompletenessAccuracy;  
                                        rcmCheck.Q5BExistenceOccur = item.Q5BExistenceOccur;
                                        rcmCheck.Q5CPresentationDisclose = item.Q5CPresentationDisclose;
                                        rcmCheck.Q5DRightObligation = item.Q5DRightObligation;
                                        rcmCheck.Q5EValuationAlloc = item.Q5EValuationAlloc;
                                        rcmCheck.ControlOwner = item.ControlOwner;
                                        rcmCheck.ControlPurpose = item.ControlPurpose;
                                        rcmCheck.ControlId = item.ControlId;
                                        rcmCheck.ControlActivity = item.ControlActivity;
                                        rcmCheck.ControlObjective = item.ControlObjective;
                                        rcmCheck.TestingPeriod = item.TestingPeriod;
                                        rcmCheck.ControlType = item.ControlType;
                                        rcmCheck.TestValidation = item.TestValidation;
                                        rcmCheck.MethodUsed = item.MethodUsed;
                                        rcmCheck.Frequency = item.Frequency;
                                        rcmCheck.Entity = item.Entity;
                                        rcmCheck.Notes = item.Notes;
                                        rcmCheck.KeyControl = item.KeyControl;
                                        rcmCheck.ControlPlaceDate = item.ControlPlaceDate;
                                        rcmCheck.NatureProcedure = item.NatureProcedure;
                                        rcmCheck.FraudControl = item.FraudControl;
                                        rcmCheck.ReviewControl = item.ReviewControl;
                                        rcmCheck.Risk = item.Risk;
                                        rcmCheck.PerformTesting = item.PerformTesting;
                                        rcmCheck.TestingProcedure = item.TestingProcedure;
                                        rcmCheck.Reviewer = item.Reviewer;
                                        rcmCheck.PopulationFileRequired = item.PopulationFileRequired;
                                        rcmCheck.PodioUniqueId = item.PodioUniqueId;
                                        rcmCheck.PodioRevision = item.PodioRevision;
                                        rcmCheck.PodioLink = item.PodioLink;

                                        listMatrixUpdate.Add(rcmCheck);
                                    }
                                    else
                                    {
                                        listMatrixAdd.Add(item);
                                    }
                                }

                                if (listMatrixAdd != null && listMatrixAdd.Count() > 0)
                                    _soxContext.AddRange(listMatrixAdd);

                                if (listMatrixUpdate != null && listMatrixUpdate.Count() > 0)
                                    _soxContext.UpdateRange(listMatrixUpdate);

                                await _soxContext.SaveChangesAsync();
                                context.Commit();
                            }
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncRcmCtaAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncRcmCtaAsync");
                return BadRequest(ex.ToString());
            }

            return Ok();

        }

        [AllowAnonymous]
        [HttpPost("podio/sync/rcm")]
        public async Task<IActionResult> SyncRcmAsync([FromBody] SyncDateRange syncDateRange)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("RcmPodioApp").GetSection("RcmAppId").Value;
                PodioAppKey.AppToken = _config.GetSection("RcmPodioApp").GetSection("RcmAppToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);
                int totalItems = 0;
                int total = 0;
                int count = 0;

                if (podio.IsAuthenticated())
                {

                    var filter = new Dictionary<string, object>
                    {
                        {
                            "last_edit_on",
                            new  { from = syncDateRange.startDate, to = syncDateRange.endDate}
                        },
                    };


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
                    int startDateField = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("StartDate").Value);
                    int endDateField = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("EndDate").Value);
                    int assignToField = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("AssignTo").Value);
                    int reviewerField = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Reviewer").Value);
                    int durationField = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Duration").Value);
                    int statusField = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Status").Value);

                    PodioCollection<Item> collectionCheck = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 1, 0, filters: filter, sortDesc: true);

                    List<Rcm> listRcm = new List<Rcm>();
                    List<Rcm> listRcmAdd = new List<Rcm>();
                    List<Rcm> listRcmUpdate = new List<Rcm>();

                    if (syncDateRange.limit == 0 && syncDateRange.offset == 0)
                    {
                        if (collectionCheck.Filtered != 0)
                        {
                            int loop = (int)Math.Ceiling((double)collectionCheck.Filtered / 500);
                            int offset = 0;

                            //get filtered items and stored in list
                            for (int i = 0; i < loop; i++)
                            {
                                PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 500, offset, filters: filter, sortDesc: true);

                                foreach (var item in collection.Items)
                                {
                                    Rcm rcm = new Rcm();
                                    count++;
                                    #region Podio Item
                                    rcm.PodioItemId = (int)item.ItemId;
                                    rcm.PodioUniqueId = item.AppItemIdFormatted.ToString();
                                    rcm.PodioRevision = item.CurrentRevision.Revision;
                                    rcm.PodioLink = item.Link.ToString();
                                    rcm.CreatedBy = item.CreatedBy.Name.ToString();
                                    rcm.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());

                                    
                                    #region Rcm Fields

                                    //Client Reference
                                    AppItemField clientApp = item.Field<AppItemField>(q2Field);
                                    IEnumerable<Item> clientAppRef = clientApp.Items;
                                    rcm.ClientName = clientAppRef.Select(x => x.Title).FirstOrDefault();
                                    rcm.ClientItemId = clientAppRef.Select(x => x.ItemId).FirstOrDefault();

                                    var clientCheck = _soxContext.ClientSs.FirstOrDefault(id => id.PodioItemId == rcm.ClientItemId);
                                    if (clientCheck != null)
                                    {
                                        rcm.ClientCode = clientCheck.ClientCode;
                                    }

                                    //Client Name
                                    TextItemField txtClientName = item.Field<TextItemField>("title");
                                    rcm.ClientNameText = txtClientName.Value;

                                    //For what year is the RCM for?
                                    TextItemField txtFy = item.Field<TextItemField>(q1Field);
                                    rcm.FY = txtFy.Value;

                                    //Process
                                    TextItemField txtProcess = item.Field<TextItemField>(q3Field);
                                    rcm.Process = txtProcess.Value;

                                    //Sub-process
                                    TextItemField txtSubProcess = item.Field<TextItemField>(q4Field);
                                    rcm.Subprocess = txtSubProcess.Value;

                                    //Control Objective
                                    TextItemField txtControlObj = item.Field<TextItemField>(q5Field);
                                    rcm.ControlObjective = txtControlObj.Value;

                                    //Specific Risk
                                    TextItemField txtSpecRisk = item.Field<TextItemField>(q6Field);
                                    rcm.SpecificRisk = txtSpecRisk.Value;
                                    
                                   

                                    //Financial Statement Elements
                                    TextItemField txtFinStatement = item.Field<TextItemField>(q7Field);
                                    rcm.FinStatemenElement = txtFinStatement.Value;

                                    //Completeness & Accuracy
                                    TextItemField txtCompleteness = item.Field<TextItemField>(q8Field);
                                    rcm.CompletenessAccuracy = txtCompleteness.Value ?? string.Empty;
                                    if (rcm.CompletenessAccuracy.ToLower().Contains("yes"))
                                    {
                                        rcm.FinancialStatementAssert = rcm.FinancialStatementAssert + "Does it show completeness & accuracy?";
                                        //test 051721
                                        rcm.CompletenessAccuracy = "Yes";
                                    }
                                       
                                    else
                                    {
                                        rcm.CompletenessAccuracy = "No";
                                    }
                                        

                                    //Existence & Occurrence
                                    //TextItemField txtExistence = item.Field<TextItemField>("existence-occurrence");
                                    TextItemField txtExistence = item.Field<TextItemField>(q8ExistenceField);
                                    rcm.ExistenceDisclosure = txtExistence.Value ?? string.Empty;
                                    if (rcm.ExistenceDisclosure.ToLower().Contains("yes"))  
                                    {
                                        rcm.FinancialStatementAssert = rcm.FinancialStatementAssert + "Does it show existence & occurrence?";
                                        rcm.ExistenceDisclosure = "Yes";
                                    }
                                    else
                                    {
                                        rcm.ExistenceDisclosure = "No";
                                    }

                                    //Presentation & Disclosure
                                    //TextItemField txtPresentation = item.Field<TextItemField>("presentation-disclosure");
                                    TextItemField txtPresentation = item.Field<TextItemField>(q8PresentationField);
                                    rcm.PresentationDisclosure = txtPresentation.Value ?? string.Empty;
                                    if (rcm.PresentationDisclosure.ToLower().Contains("yes"))
                                    {
                                        rcm.FinancialStatementAssert = rcm.FinancialStatementAssert + "Does it show presentation & disclosure?";
                                        rcm.PresentationDisclosure = "Yes";
                                    }
                                    else
                                    {
                                        rcm.PresentationDisclosure = "No";
                                    }

                                    //Rights & Obligations
                                    //TextItemField txtRightsObligation = item.Field<TextItemField>("rights-obligations");
                                    TextItemField txtRightsObligation = item.Field<TextItemField>(q8RightsField);
                                    rcm.RigthsObligation = txtRightsObligation.Value ?? string.Empty;
                                    if (rcm.RigthsObligation.ToLower().Contains("yes"))
                                       
                                    {
                                        rcm.FinancialStatementAssert = rcm.FinancialStatementAssert + "Does it show rights & obligations?";
                                        rcm.RigthsObligation = "Yes";

                                    }
                                    else
                                    {
                                        rcm.RigthsObligation = "No";
                                    }


                                    //Valuation & Allocation
                                    //TextItemField txtValuation = item.Field<TextItemField>("valuation-allocation");
                                    TextItemField txtValuation = item.Field<TextItemField>(q8ValuationField);
                                    rcm.ValuationAlloc = txtValuation.Value ?? string.Empty;
                                    if (rcm.ValuationAlloc.ToLower().Contains("yes"))
                                        
                                    {
                                        rcm.FinancialStatementAssert = rcm.FinancialStatementAssert + "Does it show valuation & allocation??";
                                        rcm.ValuationAlloc = "Yes";
                                    }
                                    else
                                    {
                                        rcm.ValuationAlloc = "No";
                                    }

                                    //Control ID
                                    TextItemField txtControlId = item.Field<TextItemField>(q9Field);
                                    rcm.ControlId = txtControlId.Value;

                                    //Control Activity FY19
                                    TextItemField txtControlActFy19 = item.Field<TextItemField>(q10Field);
                                    rcm.ControlActivityFy19 = txtControlActFy19.Value;

                                    //Control Activity FY19
                                    TextItemField txtShortDesc = item.Field<TextItemField>(shortDescField);
                                    rcm.ShortDescription = txtShortDesc.Value;

                                    //Control In Place Date
                                    TextItemField txtControlDate = item.Field<TextItemField>(q11Field);
                                    rcm.ControlPlaceDate = txtControlDate.Value;

                                    //Control Owner
                                    TextItemField txtControlOwner = item.Field<TextItemField>(q12Field);
                                    rcm.ControlOwner = txtControlOwner.Value;

                                    //Control Frequency
                                    TextItemField txtControlFreq = item.Field<TextItemField>(q13Field);
                                    rcm.ControlFrequency = txtControlFreq.Value;

                                    //Key/ Non-Key
                                    TextItemField txtKey = item.Field<TextItemField>(q14Field);
                                    rcm.Key = txtKey.Value;

                                    //Control Type
                                    TextItemField txtControlType = item.Field<TextItemField>(q15Field);
                                    rcm.ControlType = txtControlType.Value;

                                    //Nature of Procedure
                                    TextItemField txtNature = item.Field<TextItemField>(q16Field);
                                    rcm.NatureProc = txtNature.Value;

                                    //Fraud Control
                                    TextItemField txtFraudControl = item.Field<TextItemField>(q17Field);
                                    rcm.FraudControl = txtFraudControl.Value;

                                    //Risk Level
                                    TextItemField txtRisk = item.Field<TextItemField>(q18Field);
                                    rcm.RiskLvl = txtRisk.Value;

                                    //Management Review Control
                                    TextItemField txtMgmtControl = item.Field<TextItemField>(q19Field);
                                    rcm.ManagementRevControl = txtMgmtControl.Value;

                                    //PBC List
                                    TextItemField txtPbcList = item.Field<TextItemField>(q20Field);
                                    rcm.PbcList = txtPbcList.Value;

                                    //Test procedures
                                    TextItemField txtTestProc = item.Field<TextItemField>(q21Field);
                                    rcm.TestProc = txtTestProc.Value;

                                    //Start Date
                                    DateItemField startDateItem = item.Field<DateItemField>(startDateField);
                                    rcm.StartDate = startDateItem.StartDate.HasValue ? startDateItem.StartDate : null;

                                    //End Date
                                    DateItemField endDateItem = item.Field<DateItemField>(endDateField);
                                    rcm.EndDate = endDateItem.StartDate.HasValue ? endDateItem.StartDate : null;


                                    //Assign to
                                    AppItemField appFieldAssignTo = item.Field<AppItemField>(assignToField);
                                    IEnumerable<Item> assignToItems = appFieldAssignTo.Items;
                                    rcm.AssignTo = assignToItems.Select(x => x.Title).FirstOrDefault();


                                    //Reviewer
                                    AppItemField appFieldReviewer = item.Field<AppItemField>(reviewerField);
                                    IEnumerable<Item> reviewerItems = appFieldReviewer.Items;
                                    rcm.Reviewer = reviewerItems.Select(x => x.Title).FirstOrDefault();

                                    //Status
                                    CategoryItemField catStatus = item.Field<CategoryItemField>(statusField);
                                    IEnumerable<CategoryItemField.Option> listcatRisk = catStatus.Options;
                                    rcm.Status = listcatRisk.Select(x => x.Text).FirstOrDefault();

                                    //Duration
                                    DurationItemField durationItem = item.Field<DurationItemField>(durationField);
                                    rcm.Duration = durationItem.Value;

                                    #endregion

                                    #endregion

                                    listRcm.Add(rcm);
                                }

                                offset += 500;
                            }

                            total = count >= collectionCheck.Filtered ? collectionCheck.Filtered : count * (syncDateRange.offset + 1);
                            totalItems = collectionCheck.Filtered;


                        }
                    }
                    else
                    {
                        if (collectionCheck.Filtered != 0)
                        {
                            int offset = syncDateRange.offset;
                            PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), syncDateRange.limit, offset, filters: filter, sortDesc: true);

                            foreach (var item in collection.Items)
                            {
                                Rcm rcm = new Rcm();
                                count++;
                                #region Podio Item
                                rcm.PodioItemId = (int)item.ItemId;
                                rcm.PodioUniqueId = item.AppItemIdFormatted.ToString();
                                rcm.PodioRevision = item.CurrentRevision.Revision;
                                rcm.PodioLink = item.Link.ToString();
                                rcm.CreatedBy = item.CreatedBy.Name.ToString();
                                rcm.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());

                                #region Rcm Fields

                                //Client Reference
                                AppItemField clientApp = item.Field<AppItemField>(q2Field);
                                IEnumerable<Item> clientAppRef = clientApp.Items;
                                rcm.ClientName = clientAppRef.Select(x => x.Title).FirstOrDefault();
                                rcm.ClientItemId = clientAppRef.Select(x => x.ItemId).FirstOrDefault();

                                var clientCheck = _soxContext.ClientSs.FirstOrDefault(id => id.PodioItemId == rcm.ClientItemId);
                                if (clientCheck != null)
                                {
                                    rcm.ClientCode = clientCheck.ClientCode;
                                }

                                //Client Name
                                TextItemField txtClientName = item.Field<TextItemField>("title");
                                rcm.ClientNameText = txtClientName.Value;

                                //For what year is the RCM for?
                                TextItemField txtFy = item.Field<TextItemField>(q1Field);
                                rcm.FY = txtFy.Value;

                                //Process
                                TextItemField txtProcess = item.Field<TextItemField>(q3Field);
                                rcm.Process = txtProcess.Value;

                                //Sub-process
                                TextItemField txtSubProcess = item.Field<TextItemField>(q4Field);
                                rcm.Subprocess = txtSubProcess.Value;

                                //Control Objective
                                TextItemField txtControlObj = item.Field<TextItemField>(q5Field);
                                rcm.ControlObjective = txtControlObj.Value;

                                //Specific Risk
                                TextItemField txtSpecRisk = item.Field<TextItemField>(q6Field);
                                rcm.SpecificRisk = txtSpecRisk.Value;

                                //Financial Statement Elements
                                TextItemField txtFinStatement = item.Field<TextItemField>(q7Field);
                                rcm.FinStatemenElement = txtFinStatement.Value;

                                //Completeness & Accuracy
                                //TextItemField txtCompleteness = item.Field<TextItemField>(q8Field);
                                //rcm.CompletenessAccuracy = txtCompleteness.Value;

                                //TextItemField txtCompleteness = item.Field<TextItemField>("completeness - accuracy");
                                //rcm.CompletenessAccuracy = txtCompleteness.Value;

                                //Financial Statement Elements
                                TextItemField txtCompleteness = item.Field<TextItemField>(q8Field);
                                rcm.CompletenessAccuracy = txtCompleteness.Value;
                                //Existence & Occurrence
                                //TextItemField txtExistence = item.Field<TextItemField>("existence-occurrence");
                                //rcm.ExistenceDisclosure = txtExistence.Value;
                                TextItemField txtExistence = item.Field<TextItemField>(q8ExistenceField);
                                rcm.ExistenceDisclosure = txtExistence.Value;


                                //Presentation & Disclosure
                                TextItemField txtPresentation = item.Field<TextItemField>("presentation-disclosure");
                                rcm.PresentationDisclosure = txtPresentation.Value;

                                //Rights & Obligations
                                TextItemField txtRightsObligation = item.Field<TextItemField>("rights-obligations");
                                rcm.RigthsObligation = txtRightsObligation.Value;

                                //Valuation & Allocation
                                TextItemField txtValuation = item.Field<TextItemField>("valuation-allocation");
                                rcm.ValuationAlloc = txtValuation.Value;

                                //Control ID
                                TextItemField txtControlId = item.Field<TextItemField>(q9Field);
                                rcm.ControlId = txtControlId.Value;

                                //Control Activity FY19
                                TextItemField txtControlActFy19 = item.Field<TextItemField>(q10Field);
                                rcm.ControlActivityFy19 = txtControlActFy19.Value;

                                //Control Activity FY19
                                TextItemField txtShortDesc = item.Field<TextItemField>(shortDescField);
                                rcm.ShortDescription = txtShortDesc.Value;

                                //Control In Place Date
                                TextItemField txtControlDate = item.Field<TextItemField>(q11Field);
                                rcm.ControlPlaceDate = txtControlDate.Value;

                                //Control Owner
                                TextItemField txtControlOwner = item.Field<TextItemField>(q12Field);
                                rcm.ControlOwner = txtControlOwner.Value;

                                //Control Frequency
                                TextItemField txtControlFreq = item.Field<TextItemField>(q13Field);
                                rcm.ControlFrequency = txtControlFreq.Value;

                                //Key/ Non-Key
                                TextItemField txtKey = item.Field<TextItemField>(q14Field);
                                rcm.Key = txtKey.Value;

                                //Control Type
                                TextItemField txtControlType = item.Field<TextItemField>(q15Field);
                                rcm.ControlType = txtControlType.Value;

                                //Nature of Procedure
                                TextItemField txtNature = item.Field<TextItemField>(q16Field);
                                rcm.NatureProc = txtNature.Value;

                                //Fraud Control
                                TextItemField txtFraudControl = item.Field<TextItemField>(q17Field);
                                rcm.FraudControl = txtFraudControl.Value;

                                //Risk Level
                                TextItemField txtRisk = item.Field<TextItemField>(q18Field);
                                rcm.RiskLvl = txtRisk.Value;

                                //Management Review Control
                                TextItemField txtMgmtControl = item.Field<TextItemField>(q19Field);
                                rcm.ManagementRevControl = txtMgmtControl.Value;

                                //PBC List
                                TextItemField txtPbcList = item.Field<TextItemField>(q20Field);
                                rcm.PbcList = txtPbcList.Value;

                                //Test procedures
                                TextItemField txtTestProc = item.Field<TextItemField>(q21Field);
                                rcm.TestProc = txtTestProc.Value;

                                //Start Date
                                DateItemField startDateItem = item.Field<DateItemField>(startDateField);
                                rcm.StartDate = startDateItem.StartDate.HasValue ? startDateItem.StartDate : null;

                                //End Date
                                DateItemField endDateItem = item.Field<DateItemField>(endDateField);
                                rcm.EndDate = endDateItem.StartDate.HasValue ? endDateItem.StartDate : null;


                                //Assign to
                                AppItemField appFieldAssignTo = item.Field<AppItemField>(assignToField);
                                IEnumerable<Item> assignToItems = appFieldAssignTo.Items;
                                rcm.AssignTo = assignToItems.Select(x => x.Title).FirstOrDefault();


                                //Reviewer
                                AppItemField appFieldReviewer = item.Field<AppItemField>(reviewerField);
                                IEnumerable<Item> reviewerItems = appFieldReviewer.Items;
                                rcm.Reviewer = reviewerItems.Select(x => x.Title).FirstOrDefault();

                                //Status
                                CategoryItemField catStatus = item.Field<CategoryItemField>(statusField);
                                IEnumerable<CategoryItemField.Option> listcatRisk = catStatus.Options;
                                rcm.Status = listcatRisk.Select(x => x.Text).FirstOrDefault();

                                //Duration
                                DurationItemField durationItem = item.Field<DurationItemField>(durationField);
                                rcm.Duration = durationItem.Value;

                                #endregion

                                #endregion

                                listRcm.Add(rcm);
                            }

                            count = count * (syncDateRange.offset + 1);
                            total = count >= collection.Filtered ? collection.Filtered : count;
                            totalItems = collection.Filtered;
                        }
                    }


                    if (listRcm != null && listRcm.Count() > 0)
                    {
                        using (var context = _soxContext.Database.BeginTransaction())
                        {

                            foreach (var item in listRcm)
                            {
                                var rcmCheck = _soxContext.Rcm.FirstOrDefault(id => id.PodioItemId == (int)item.PodioItemId);
                                if (rcmCheck != null)
                                {
                                    rcmCheck.ClientName = item.ClientName;
                                    rcmCheck.ClientCode = item.ClientCode;
                                    rcmCheck.ClientItemId = item.ClientItemId;
                                    rcmCheck.ClientNameText = item.ClientNameText;
                                    rcmCheck.FY = item.FY;
                                    rcmCheck.Process = item.Process;
                                    rcmCheck.Subprocess = item.Subprocess;
                                    rcmCheck.ControlObjective = item.ControlObjective;
                                    rcmCheck.SpecificRisk = item.SpecificRisk;
                                    rcmCheck.FinStatemenElement = item.FinStatemenElement;
                                    rcmCheck.FinancialStatementAssert = item.FinancialStatementAssert;
                                    rcmCheck.CompletenessAccuracy = item.CompletenessAccuracy;
                                    rcmCheck.ExistenceDisclosure = item.ExistenceDisclosure;
                                    rcmCheck.PresentationDisclosure = item.PresentationDisclosure;
                                    rcmCheck.RigthsObligation = item.RigthsObligation;
                                    rcmCheck.ValuationAlloc = item.ValuationAlloc;
                                    rcmCheck.ControlId = item.ControlId;
                                    rcmCheck.ControlActivityFy19 = item.ControlActivityFy19;
                                    rcmCheck.ControlPlaceDate = item.ControlPlaceDate;
                                    rcmCheck.ControlOwner = item.ControlOwner;
                                    rcmCheck.ControlFrequency = item.ControlFrequency;
                                    rcmCheck.Key = item.Key;
                                    rcmCheck.ControlType = item.ControlType;
                                    rcmCheck.NatureProc = item.NatureProc;
                                    rcmCheck.FraudControl = item.FraudControl;
                                    rcmCheck.RiskLvl = item.RiskLvl;
                                    rcmCheck.ManagementRevControl = item.ManagementRevControl;
                                    rcmCheck.PbcList = item.PbcList;
                                    rcmCheck.TestProc = item.TestProc;
                                    rcmCheck.Status = item.Status;
                                    rcmCheck.PodioUniqueId = item.PodioUniqueId;
                                    rcmCheck.PodioRevision = item.PodioRevision;
                                    rcmCheck.PodioLink = item.PodioLink;
                                    rcmCheck.ShortDescription = item.ShortDescription;

                                    listRcmUpdate.Add(rcmCheck);
                                }
                                else
                                {
                                    listRcmAdd.Add(item);
                                }
                            }

                            if (listRcmAdd != null && listRcmAdd.Count() > 0)
                                _soxContext.AddRange(listRcmAdd);

                            if (listRcmUpdate != null && listRcmUpdate.Count() > 0)
                                _soxContext.UpdateRange(listRcmUpdate);

                            await _soxContext.SaveChangesAsync();
                            context.Commit();
                        }
                    }


                }

                string saveItems = $"{total}/{totalItems}";
                return Ok(new { syncDateRange.limit, totalItems, saveItems });

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncRcmAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncRcmAsync");
                return BadRequest(ex.ToString());
            }

        }


        //[AllowAnonymous]
        [HttpPost("podio/sync/rcmprocess")]
        public async Task<IActionResult> SyncRcmProcessAsync([FromBody] SyncDateRange syncDateRange)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("RcmPodioApp").GetSection("RcmProcessAppId").Value;
                PodioAppKey.AppToken = _config.GetSection("RcmPodioApp").GetSection("RcmProcessAppToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                if (podio.IsAuthenticated())
                {

                    var filter = new Dictionary<string, object>
                    {
                        {
                            "last_edit_on",
                            new  { from = syncDateRange.startDate, to = syncDateRange.endDate}
                        },
                    };

                    PodioCollection<Item> collectionCheck = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 1, 0, filters: filter, sortDesc: true);

                    if (collectionCheck.Filtered != 0)
                    {
                        int loop = (int)Math.Ceiling((double)collectionCheck.Filtered / 500);
                        int offset = 0;
                        List<RcmProcess> listRcmProcess = new List<RcmProcess>();
                        List<RcmProcess> listRcmAdd = new List<RcmProcess>();
                        List<RcmProcess> listRcmUpdate = new List<RcmProcess>();
                        //get filtered items and stored in list
                        for (int i = 0; i < loop; i++)
                        {
                            PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 500, offset, filters: filter, sortDesc: true);

                            foreach (var item in collection.Items)
                            {
                                RcmProcess rcmProcess = new RcmProcess();

                                #region Podio Item
                                rcmProcess.PodioItemId = (int)item.ItemId;
                                rcmProcess.PodioUniqueId = item.AppItemIdFormatted.ToString();
                                rcmProcess.PodioRevision = item.CurrentRevision.Revision;
                                rcmProcess.PodioLink = item.Link.ToString();
                                rcmProcess.CreatedBy = item.CreatedBy.Name.ToString();
                                rcmProcess.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());

                                #region Fields

                                //get process value
                                TextItemField textProcess = item.Field<TextItemField>(int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmProcessField").GetSection("Process").Value.ToString()));
                                rcmProcess.Process = textProcess.Value;

                                #endregion

                                #endregion

                                listRcmProcess.Add(rcmProcess);
                            }

                            offset += 500;
                        }

                        if (listRcmProcess != null && listRcmProcess.Count() > 0)
                        {
                            using (var context = _soxContext.Database.BeginTransaction())
                            {

                                foreach (var item in listRcmProcess)
                                {
                                    var rcmCheck = _soxContext.RcmProcess.FirstOrDefault(id => id.PodioItemId == (int)item.PodioItemId);
                                    if (rcmCheck != null)
                                    {
  
                                        rcmCheck.Process = item.Process; 
                                        rcmCheck.PodioUniqueId = item.PodioUniqueId;
                                        rcmCheck.PodioRevision = item.PodioRevision;
                                        rcmCheck.PodioLink = item.PodioLink;

                                        listRcmUpdate.Add(rcmCheck);
                                    }
                                    else
                                    {
                                        listRcmAdd.Add(item);
                                    }
                                }

                                if (listRcmAdd != null && listRcmAdd.Count() > 0)
                                    _soxContext.AddRange(listRcmAdd);

                                if (listRcmUpdate != null && listRcmUpdate.Count() > 0)
                                    _soxContext.UpdateRange(listRcmUpdate);

                                await _soxContext.SaveChangesAsync();
                                context.Commit();
                            }
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncRcmProcessAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncRcmProcessAsync");
                return BadRequest(ex.ToString());
            }

            return Ok();

        }

        //[AllowAnonymous]
        [HttpPost("podio/sync/rcmsubprocess")]
        public async Task<IActionResult> SyncRcmSubProcessAsync([FromBody] SyncDateRange syncDateRange)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("RcmPodioApp").GetSection("RcmSubProcessAppId").Value;
                PodioAppKey.AppToken = _config.GetSection("RcmPodioApp").GetSection("RcmSubProcessAppToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                if (podio.IsAuthenticated())
                {

                    var filter = new Dictionary<string, object>
                    {
                        {
                            "last_edit_on",
                            new  { from = syncDateRange.startDate, to = syncDateRange.endDate}
                        },
                    };

                    PodioCollection<Item> collectionCheck = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 1, 0, filters: filter, sortDesc: true);

                    if (collectionCheck.Filtered != 0)
                    {
                        int loop = (int)Math.Ceiling((double)collectionCheck.Filtered / 500);
                        int offset = 0;
                        List<RcmSubProcess> listRcmSubprocess = new List<RcmSubProcess>();
                        List<RcmSubProcess> listRcmAdd = new List<RcmSubProcess>();
                        List<RcmSubProcess> listRcmUpdate = new List<RcmSubProcess>();
                        //get filtered items and stored in list
                        for (int i = 0; i < loop; i++)
                        {
                            PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 500, offset, filters: filter, sortDesc: true);

                            foreach (var item in collection.Items)
                            {
                                RcmSubProcess rcmSubProcess = new RcmSubProcess();

                                #region Podio Item
                                rcmSubProcess.PodioItemId = (int)item.ItemId;
                                rcmSubProcess.PodioUniqueId = item.AppItemIdFormatted.ToString();
                                rcmSubProcess.PodioRevision = item.CurrentRevision.Revision;
                                rcmSubProcess.PodioLink = item.Link.ToString();
                                rcmSubProcess.CreatedBy = item.CreatedBy.Name.ToString();
                                rcmSubProcess.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());

                                #region Fields

                                //get process value
                                TextItemField textProcess = item.Field<TextItemField>(int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmSubProcessField").GetSection("SubProcess").Value.ToString()));
                                rcmSubProcess.SubProcess = textProcess.Value;

                                #endregion

                                #endregion

                                listRcmSubprocess.Add(rcmSubProcess);
                            }

                            offset += 500;
                        }

                        if (listRcmSubprocess != null && listRcmSubprocess.Count() > 0)
                        {
                            using (var context = _soxContext.Database.BeginTransaction())
                            {

                                foreach (var item in listRcmSubprocess)
                                {
                                    var rcmCheck = _soxContext.RcmSubProcess.FirstOrDefault(id => id.PodioItemId == (int)item.PodioItemId);
                                    if (rcmCheck != null)
                                    {

                                        rcmCheck.SubProcess = item.SubProcess;
                                        rcmCheck.PodioUniqueId = item.PodioUniqueId;
                                        rcmCheck.PodioRevision = item.PodioRevision;
                                        rcmCheck.PodioLink = item.PodioLink;

                                        listRcmUpdate.Add(rcmCheck);
                                    }
                                    else
                                    {
                                        listRcmAdd.Add(item);
                                    }
                                }

                                if (listRcmAdd != null && listRcmAdd.Count() > 0)
                                    _soxContext.AddRange(listRcmAdd);

                                if (listRcmUpdate != null && listRcmUpdate.Count() > 0)
                                    _soxContext.UpdateRange(listRcmUpdate);

                                await _soxContext.SaveChangesAsync();
                                context.Commit();
                            }
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncRcmSubProcessAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncRcmSubProcessAsync");
                return BadRequest(ex.ToString());
            }

            return Ok();

        }


        //[AllowAnonymous]
        [HttpPost("podio/sync/rcmcontrolid")]
        public async Task<IActionResult> SyncRcmControlIdAsync([FromBody] SyncDateRange syncDateRange)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("RcmPodioApp").GetSection("RcmControlIdAppId").Value;
                PodioAppKey.AppToken = _config.GetSection("RcmPodioApp").GetSection("RcmControlIdAppToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                if (podio.IsAuthenticated())
                {

                    var filter = new Dictionary<string, object>
                    {
                        {
                            "last_edit_on",
                            new  { from = syncDateRange.startDate, to = syncDateRange.endDate}
                        },
                    };

                    PodioCollection<Item> collectionCheck = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 1, 0, filters: filter, sortDesc: true);

                    if (collectionCheck.Filtered != 0)
                    {
                        int loop = (int)Math.Ceiling((double)collectionCheck.Filtered / 500);
                        int offset = 0;
                        List<RcmControlId> listRcmControlId = new List<RcmControlId>();
                        List<RcmControlId> listRcmAdd = new List<RcmControlId>();
                        List<RcmControlId> listRcmUpdate = new List<RcmControlId>();
                        //get filtered items and stored in list
                        for (int i = 0; i < loop; i++)
                        {
                            PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 500, offset, filters: filter, sortDesc: true);

                            foreach (var item in collection.Items)
                            {
                                RcmControlId rcmControlId = new RcmControlId();

                                #region Podio Item
                                rcmControlId.PodioItemId = (int)item.ItemId;
                                rcmControlId.PodioUniqueId = item.AppItemIdFormatted.ToString();
                                rcmControlId.PodioRevision = item.CurrentRevision.Revision;
                                rcmControlId.PodioLink = item.Link.ToString();
                                rcmControlId.CreatedBy = item.CreatedBy.Name.ToString();
                                rcmControlId.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());

                                #region Fields

                                //get process value
                                TextItemField textProcess = item.Field<TextItemField>(int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmControlIdField").GetSection("ControlId").Value.ToString()));
                                rcmControlId.ControlId = textProcess.Value;

                                #endregion

                                #endregion

                                listRcmControlId.Add(rcmControlId);
                            }

                            offset += 500;
                        }

                        if (listRcmControlId != null && listRcmControlId.Count() > 0)
                        {
                            using (var context = _soxContext.Database.BeginTransaction())
                            {

                                foreach (var item in listRcmControlId)
                                {
                                    var rcmCheck = _soxContext.RcmControlId.FirstOrDefault(id => id.PodioItemId == (int)item.PodioItemId);
                                    if (rcmCheck != null)
                                    {

                                        rcmCheck.ControlId = item.ControlId;
                                        rcmCheck.PodioUniqueId = item.PodioUniqueId;
                                        rcmCheck.PodioRevision = item.PodioRevision;
                                        rcmCheck.PodioLink = item.PodioLink;

                                        listRcmUpdate.Add(rcmCheck);
                                    }
                                    else
                                    {
                                        listRcmAdd.Add(item);
                                    }
                                }

                                if (listRcmAdd != null && listRcmAdd.Count() > 0)
                                    _soxContext.AddRange(listRcmAdd);

                                if (listRcmUpdate != null && listRcmUpdate.Count() > 0)
                                    _soxContext.UpdateRange(listRcmUpdate);

                                await _soxContext.SaveChangesAsync();
                                context.Commit();
                            }
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncRcmControlIdAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncRcmControlIdAsync");
                return BadRequest(ex.ToString());
            }

            return Ok();

        }


        //[AllowAnonymous]
        [HttpPost("rcm/questionnaire/fields")]
        public async Task<IActionResult> SyncRcmQuestionnaireFields()
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("RcmPodioApp").GetSection("RcmQuestionnaireAppId").Value;
                PodioAppKey.AppToken = _config.GetSection("RcmPodioApp").GetSection("RcmQuestionnaireAppToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(int.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                if (podio.IsAuthenticated())
                {
                    Debug.WriteLine("Podio Authenticated");
                    int count = 1;
                    int position = 1;
                    var podioResult = await podio.ApplicationService.GetApp(int.Parse(PodioAppKey.AppId));
                    List<RcmQuestionnaireField> listRcmQuestionnaireField = new List<RcmQuestionnaireField>();
                    if (podioResult.Fields != null && podioResult.Fields.Count > 0)
                    {

                        if (podioResult.Fields.Count > 0)
                        {

                            using (var context = _soxContext.Database.BeginTransaction())
                            {
                                //Remove all fields associated with app id
                                var checkRcmQuestionnaireField = _soxContext.RcmQuestionnaireField
                                    .Where(
                                        x => x.AppId.Equals(PodioAppKey.AppId))
                                    .FirstOrDefault();
                                if (checkRcmQuestionnaireField != null)
                                {

                                    _soxContext.RemoveRange(_soxContext.RcmQuestionnaireFieldOptions.Where(id => id.AppId == int.Parse(PodioAppKey.AppId)));
                                    _soxContext.RemoveRange(_soxContext.RcmQuestionnaireField.Where(id => id.AppId == PodioAppKey.AppId));
                                    await _soxContext.SaveChangesAsync();
                                    context.Commit();
                                }
                            }


                            foreach (var field in podioResult.Fields)
                            {
                                if (field.Status == "active")
                                {
                                    using (var context = _soxContext.Database.BeginTransaction())
                                    {
                                        //writeLog.Display(field);
                                        RcmQuestionnaireField rcmQuestionnaireField = new RcmQuestionnaireField();
                                        rcmQuestionnaireField.QuestionString = field.Label;
                                        rcmQuestionnaireField.Type = field.Type;
                                        rcmQuestionnaireField.AppId = PodioAppKey.AppId;
                                        rcmQuestionnaireField.FieldId = field.FieldId;
                                        rcmQuestionnaireField.CreatedOn = DateTime.Now;
                                        rcmQuestionnaireField.UpdatedOn = DateTime.Now;
                                        rcmQuestionnaireField.Position = position;
                                        List<RcmQuestionnaireFieldOptions> listOptions = new List<RcmQuestionnaireFieldOptions>();

                                        if (field.InternalConfig.Description != string.Empty)
                                        {
                                            rcmQuestionnaireField.Description = field.InternalConfig.Description;
                                        }

                                        if (field.Type == "category")
                                        {
                                            foreach (var option in field.InternalConfig.Settings)
                                            {
                                                if (option.Value.HasValues)
                                                {

                                                    foreach (var item in option.Value)
                                                    {
                                                        if (item["status"].ToString() == "active")
                                                        {
                                                            RcmQuestionnaireFieldOptions RcmOption = new RcmQuestionnaireFieldOptions();
                                                            //Console.WriteLine($"{item["text"]}");
                                                            RcmOption.OptionName = item["text"].ToString();
                                                            RcmOption.CreatedOn = DateTime.Now;
                                                            RcmOption.UpdatedOn = DateTime.Now;
                                                            RcmOption.AppId = int.Parse(rcmQuestionnaireField.AppId);
                                                            RcmOption.RcmQuestionnaire = rcmQuestionnaireField;
                                                            RcmOption.OptionId = $"{rcmQuestionnaireField.AppId}{count}";
                                                            listOptions.Add(RcmOption);
                                                            count++;
                                                        }

                                                    }

                                                    rcmQuestionnaireField.Options = listOptions;
                                                }

                                            }
                                        }

                                        if (field.Type == "date")
                                        {
                                            Debug.WriteLine($"{field.InternalConfig.Settings["end"]}");
                                            if (field.InternalConfig.Settings["end"].ToString() == "enabled")
                                                rcmQuestionnaireField.Tag = "enabled";
                                            else
                                                rcmQuestionnaireField.Tag = string.Empty;

                                        }

                                        if (field.Type == "text")
                                        {
                                            if (field.InternalConfig.Settings["size"].ToString() == "large")
                                            {
                                                rcmQuestionnaireField.Tag = "large";
                                            }
                                        }

                                        if (field.Type == "image")
                                        {
                                            //image field 
                                        }

                                        if (field.Type == "app")
                                        {
                                            //image field 
                                        }

                                        



                                        _soxContext.Add(rcmQuestionnaireField);
                                        await _soxContext.SaveChangesAsync();
                                        context.Commit();
                                        position++;
                                    }
                                }
                            }

                        }

                        return Ok();
                    }

                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncRcmQuestionnaireFields");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncRcmQuestionnaireFields");
                return BadRequest(ex.ToString());
            }

            return NoContent();

        }


        //[AllowAnonymous]
        [HttpPost("rcm/questionnaire")]
        public async Task<IActionResult> SyncRcmQuestionnaireDefine([FromBody] SyncDateRange syncDateRange)
        {
            //Console.WriteLine("Triggered process");
            try
            {
                #region Start Syncing Process
                //Initiliaze
                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();

                //Get API keys in settings.json
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("RcmPodioApp").GetSection("RcmQuestionnaireAppId").Value;
                PodioAppKey.AppToken = _config.GetSection("RcmPodioApp").GetSection("RcmQuestionnaireAppToken").Value;

                //connect to podio using api keys and search for item using podio api filter
                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(int.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                if (podio.IsAuthenticated())
                {
                    //podio filter use in querying data
                    var filter = new Dictionary<string, object>
                    {
                        {
                            "last_edit_on",
                            new  { from = syncDateRange.startDate, to = syncDateRange.endDate}
                        },
                    };

                    //get collection
                    PodioCollection<Item> collectionCheck = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 1, 0, filters: filter, sortDesc: true);

                    if (collectionCheck.Filtered != 0)
                    {
                        int loop = (int)Math.Ceiling((double)collectionCheck.Filtered / 500);
                        int offset = 0;
                        WriteLog writeLog = new WriteLog();
                        PodioCollection<Item> pubCollection = new PodioCollection<Item>();
                        List<RcmQuestionnaire> listRcmQuestionnaire = new List<RcmQuestionnaire>();
                        //get filtered items and stored in list
                        for (int i = 0; i < loop; i++)
                        {
                            PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 500, offset, filters: filter, sortDesc: true);
                            //PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 2, offset, filters: filter, sortDesc: true);
                            if (collection.Items != null && collection.Items.Count() > 0)
                            {
                                foreach (var item in collection.Items)
                                {
                                    
                                    RcmQuestionnaire rcmQuestionnaire = new RcmQuestionnaire();
                                    rcmQuestionnaire.PodioItemId = (int)item.ItemId;
                                    rcmQuestionnaire.PodioUniqueId = item.AppItemIdFormatted.ToString();
                                    rcmQuestionnaire.PodioRevision = item.CurrentRevision.Revision;
                                    rcmQuestionnaire.PodioLink = item.Link.ToString();
                                    rcmQuestionnaire.CreatedBy = item.CreatedBy.Name.ToString();
                                    rcmQuestionnaire.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());

                                    //foreach (var field in item.Fields.ToList())
                                    //{
                                    //}

                                    #region 1. For what year is the RCM for?
                                    //Question 1
                                    rcmQuestionnaire.Q1FieldId = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmQuestionnaire").GetSection("Q1").Value);
                                        CategoryItemField catFY = item.Field<CategoryItemField>(rcmQuestionnaire.Q1FieldId);
                                        IEnumerable<CategoryItemField.Option> listCat = catFY.Options != null ? catFY.Options : null;
                                        rcmQuestionnaire.Q1Label = catFY.Label; //get label

                                        if (listCat != null && listCat.Count() > 0) //check if category has value then get all category text
                                        {
                                            List<RcmFY> listRcmFy = new List<RcmFY>();
                                            foreach (var itemFy in listCat.ToList())
                                            {
                                                if (itemFy.Status == "active")
                                                {
                                                    listRcmFy.Add(new RcmFY { FY = itemFy.Text, RcmQuestionnaire = rcmQuestionnaire });
                                                }
                                            }
                                            rcmQuestionnaire.ListQ1Year = listRcmFy;
                                        }
                                        #endregion

                                    #region 2. Client
                                    //Question 2
                                    rcmQuestionnaire.Q2FieldId = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmQuestionnaire").GetSection("Q2").Value);
                                    AppItemField clientField = item.Field<AppItemField>(rcmQuestionnaire.Q2FieldId);
                                    IEnumerable<Item> relatedClientItems = clientField.Items != null ? clientField.Items : null;
                                    rcmQuestionnaire.Q2Label = clientField.Label; //get label
                                        
                                    if (relatedClientItems != null && relatedClientItems.Count() > 0)
                                    {
                                        rcmQuestionnaire.Q2Client = relatedClientItems.Select(x => x.Title).FirstOrDefault();
                                        rcmQuestionnaire.Q2ClientItemId = relatedClientItems.Select(x => x.ItemId).FirstOrDefault();
                                    }
                                    #endregion

                                    #region 3. Process
                                    //Question 3
                                    rcmQuestionnaire.Q3FieldId = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmQuestionnaire").GetSection("Q3").Value);
                                    AppItemField processField = item.Field<AppItemField>(rcmQuestionnaire.Q3FieldId);
                                    IEnumerable<Item> relatedProcessItems = processField.Items != null ? processField.Items : null;
                                    rcmQuestionnaire.Q3Label = processField.Label; //get label
                                        
                                    if (relatedProcessItems != null && relatedProcessItems.Count() > 0)
                                    {
                                        rcmQuestionnaire.Q3Process = relatedProcessItems.Select(x => x.Title).FirstOrDefault();
                                        rcmQuestionnaire.Q3ProcessItemId = relatedProcessItems.Select(x => x.ItemId).FirstOrDefault();
                                    }
                                    #endregion

                                    #region 4. Sub-Process
                                    //Question 4
                                    rcmQuestionnaire.Q4FieldId = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmQuestionnaire").GetSection("Q4").Value);
                                    AppItemField subProcessField = item.Field<AppItemField>(rcmQuestionnaire.Q4FieldId);
                                    IEnumerable<Item> relatedSubProcessItems = subProcessField.Items != null ? subProcessField.Items : null;
                                    rcmQuestionnaire.Q4Label = subProcessField.Label; //get label
 
                                    if (relatedSubProcessItems != null && relatedSubProcessItems.Count() > 0)
                                    {
                                        rcmQuestionnaire.Q4SubProcess = relatedSubProcessItems.Select(x => x.Title).FirstOrDefault();
                                        rcmQuestionnaire.Q4SubProcessItemId = relatedSubProcessItems.Select(x => x.ItemId).FirstOrDefault();
                                    }
                                    #endregion

                                    #region 5. What is the control objective
                                    //Question 5
                                    rcmQuestionnaire.Q5FieldId = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmQuestionnaire").GetSection("Q5").Value);
                                    TextItemField controlObjField = item.Field<TextItemField>(rcmQuestionnaire.Q5FieldId);
                                    rcmQuestionnaire.Q5ControlObjective = controlObjField.Value != null ? controlObjField.Value.ToString() : string.Empty;
                                    rcmQuestionnaire.Q5ControlObjective = rcmQuestionnaire.Q5ControlObjective.Replace("<p>", String.Empty).Replace("</p>", String.Empty);
                                    rcmQuestionnaire.Q5Label = controlObjField.Label;

                                    #endregion

                                    #region 6. What is the specific risk 
                                    //Question 6
                                    rcmQuestionnaire.Q6FieldId = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmQuestionnaire").GetSection("Q6").Value);
                                    TextItemField riskField = item.Field<TextItemField>(rcmQuestionnaire.Q6FieldId);
                                    rcmQuestionnaire.Q6SpecificRisk = riskField.Value != null ? riskField.Value.ToString() : string.Empty;
                                    rcmQuestionnaire.Q6SpecificRisk = rcmQuestionnaire.Q6SpecificRisk.Replace("<p>", String.Empty).Replace("</p>", String.Empty);
                                    rcmQuestionnaire.Q6Label = riskField.Label;
                                    #endregion

                                    #region 7. What is the financial statement element
                                    //Question 7
                                    rcmQuestionnaire.Q7FieldId = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmQuestionnaire").GetSection("Q7").Value);
                                    AppItemField finStateElementField = item.Field<AppItemField>(rcmQuestionnaire.Q7FieldId);
                                    IEnumerable<Item> relatedFinStateElementItems = finStateElementField.Items != null ? finStateElementField.Items : null;
                                    rcmQuestionnaire.Q7Label = finStateElementField.Label; //get label

                                    if (relatedFinStateElementItems != null && relatedFinStateElementItems.Count() > 0)
                                    {
                                        List<RcmFinancialStatement> listFinStatement = new List<RcmFinancialStatement>();
                                        foreach (var finStatement in relatedFinStateElementItems.ToList())
                                        {
                                            RcmFinancialStatement rcmFinState = new RcmFinancialStatement();
                                            rcmFinState.FinStatement = finStatement.Title;
                                            rcmFinState.PodioItemId = finStatement.ItemId;
                                            rcmFinState.PodioUniqueId = finStatement.AppItemIdFormatted != null ? finStatement.AppItemIdFormatted.ToString() : string.Empty;
                                            rcmFinState.PodioRevision = finStatement.CurrentRevision != null ? finStatement.CurrentRevision.Revision : 0;
                                            rcmFinState.PodioLink = finStatement.Link.ToString();
                                            rcmFinState.CreatedBy = finStatement.CreatedBy.Name.ToString();
                                            rcmFinState.CreatedOn = DateTime.Parse(finStatement.CreatedOn.ToString());

                                            listFinStatement.Add(rcmFinState);
                                        }
                                        rcmQuestionnaire.ListQ7FinStatementElement = listFinStatement;
                                    }
                                    #endregion

                                    #region 8. Financial Statement Assertions

                                    //Question 8
                                    rcmQuestionnaire.Q8FieldId = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmQuestionnaire").GetSection("Q8").Value);
                                    CategoryItemField catFinAssert = item.Field<CategoryItemField>(rcmQuestionnaire.Q8FieldId);
                                    IEnumerable<CategoryItemField.Option> listFinAssert = catFinAssert.Options != null ? catFinAssert.Options : null;
                                    rcmQuestionnaire.Q8Label = catFinAssert.Label; //get label

                                    if (listFinAssert != null && listFinAssert.Count() > 0) //check if category has value then get all category text
                                    {
                                        List<RcmFinancialStatementAssert> listFinStateAssert = new List<RcmFinancialStatementAssert>();
                                        foreach (var itemFinAssert in listFinAssert.ToList())
                                        {
                                            if (itemFinAssert.Status == "active")
                                            {
                                                listFinStateAssert.Add(new RcmFinancialStatementAssert { FinStatementAssert = itemFinAssert.Text, RcmQuestionnaire = rcmQuestionnaire });
                                            }
                                        }
                                        rcmQuestionnaire.ListQ8FinStatementAssert = listFinStateAssert;
                                    }
                                    #endregion

                                    #region 9. What is the Control ID

                                    //Question 9
                                    rcmQuestionnaire.Q9FieldId = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmQuestionnaire").GetSection("Q9").Value);
                                    AppItemField controlIdField = item.Field<AppItemField>(rcmQuestionnaire.Q9FieldId);
                                    IEnumerable<Item> relatedControlIdItems = controlIdField.Items != null ? controlIdField.Items : null;
                                    rcmQuestionnaire.Q9Label = controlIdField.Label; //get label

                                    if (relatedControlIdItems != null && relatedControlIdItems.Count() > 0)
                                    {
                                        List<RcmControlId> listControlId = new List<RcmControlId>();
                                        foreach (var controlId in relatedControlIdItems.ToList())
                                        {
                                            RcmControlId rcmControlId = new RcmControlId();
                                            rcmControlId.ControlId = controlId.Title;
                                            rcmControlId.PodioItemId = controlId.ItemId;
                                            rcmControlId.PodioUniqueId = controlId.AppItemIdFormatted != null ? controlId.AppItemIdFormatted.ToString() : string.Empty;
                                            rcmControlId.PodioRevision = controlId.CurrentRevision != null ? controlId.CurrentRevision.Revision : 0;
                                            rcmControlId.PodioLink = controlId.Link.ToString();
                                            rcmControlId.CreatedBy = controlId.CreatedBy.Name.ToString();
                                            rcmControlId.CreatedOn = DateTime.Parse(controlId.CreatedOn.ToString());

                                            listControlId.Add(rcmControlId);
                                        }
                                        rcmQuestionnaire.ListQ9ControlId = listControlId;
                                    }

                                    #endregion

                                    #region 10. What is the control activity
                                    //Question 10
                                    rcmQuestionnaire.Q10FieldId = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmQuestionnaire").GetSection("Q10").Value);
                                    TextItemField riskControlActivity = item.Field<TextItemField>(rcmQuestionnaire.Q10FieldId);
                                    rcmQuestionnaire.Q10ControlActivity = riskControlActivity.Value != null ? riskControlActivity.Value.ToString() : string.Empty;
                                    rcmQuestionnaire.Q10ControlActivity = rcmQuestionnaire.Q10ControlActivity.Replace("<p>", String.Empty).Replace("</p>", String.Empty);
                                    rcmQuestionnaire.Q10Label = riskControlActivity.Label;
                                    #endregion

                                    #region 11. When is the control in place date
                                    //Question 11
                                    rcmQuestionnaire.Q11FieldId = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmQuestionnaire").GetSection("Q11").Value);
                                    TextItemField riskPlaceDate = item.Field<TextItemField>(rcmQuestionnaire.Q11FieldId);
                                    rcmQuestionnaire.Q11ControlInPlace = riskPlaceDate.Value != null ? riskPlaceDate.Value.ToString() : string.Empty;
                                    rcmQuestionnaire.Q11ControlInPlace = rcmQuestionnaire.Q11ControlInPlace.Replace("<p>", String.Empty).Replace("</p>", String.Empty);
                                    rcmQuestionnaire.Q11Label = riskPlaceDate.Label;
                                    #endregion

                                    #region 12. Who is the control owner
                                    //Question 12
                                    rcmQuestionnaire.Q12FieldId = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmQuestionnaire").GetSection("Q12").Value);
                                    AppItemField appControlOwner = item.Field<AppItemField>(rcmQuestionnaire.Q12FieldId);
                                    IEnumerable<Item> relatedControlOwnerItems = appControlOwner.Items != null ? appControlOwner.Items : null;
                                    rcmQuestionnaire.Q12Label = appControlOwner.Label; //get label

                                    if (relatedControlOwnerItems != null && relatedControlOwnerItems.Count() > 0)
                                    {
                                        List<RcmControlOwner> listControlOwner = new List<RcmControlOwner>();
                                        foreach (var controlOwner in relatedControlOwnerItems.ToList())
                                        {
                                            RcmControlOwner rcmControlOwner = new RcmControlOwner();
                                            rcmControlOwner.ControlOwner = controlOwner.Title;
                                            rcmControlOwner.PodioItemId = controlOwner.ItemId;
                                            rcmControlOwner.PodioUniqueId = controlOwner.AppItemIdFormatted != null ? controlOwner.AppItemIdFormatted.ToString() : string.Empty;
                                            rcmControlOwner.PodioRevision = controlOwner.CurrentRevision != null ? controlOwner.CurrentRevision.Revision : 0;
                                            rcmControlOwner.PodioLink = controlOwner.Link.ToString();
                                            rcmControlOwner.CreatedBy = controlOwner.CreatedBy.Name.ToString();
                                            rcmControlOwner.CreatedOn = DateTime.Parse(controlOwner.CreatedOn.ToString());

                                            listControlOwner.Add(rcmControlOwner);
                                        }
                                        rcmQuestionnaire.ListQ12ControlOwner = listControlOwner;
                                    }
                                    #endregion

                                    #region 13. What is the control frequency
                                    //Question 13
                                    rcmQuestionnaire.Q13FieldId = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmQuestionnaire").GetSection("Q13").Value);
                                    AppItemField appControlFreq = item.Field<AppItemField>(rcmQuestionnaire.Q13FieldId);
                                    IEnumerable<Item> relatedControlFreqItems = appControlFreq.Items != null ? appControlFreq.Items : null;
                                    rcmQuestionnaire.Q13Label = appControlFreq.Label; //get label

                                    if (relatedControlFreqItems != null && relatedControlFreqItems.Count() > 0)
                                    {
                                        List<RcmFrequency> listControlFreq = new List<RcmFrequency>();
                                        foreach (var controlFreq in relatedControlFreqItems.ToList())
                                        {
                                            RcmFrequency rcmControlFreq = new RcmFrequency();
                                            rcmControlFreq.Frequency = controlFreq.Title;
                                            rcmControlFreq.PodioItemId = controlFreq.ItemId;
                                            rcmControlFreq.PodioUniqueId = controlFreq.AppItemIdFormatted != null ? controlFreq.AppItemIdFormatted.ToString() : string.Empty;
                                            rcmControlFreq.PodioRevision = controlFreq.CurrentRevision != null ? controlFreq.CurrentRevision.Revision : 0;
                                            rcmControlFreq.PodioLink = controlFreq.Link.ToString();
                                            rcmControlFreq.CreatedBy = controlFreq.CreatedBy.Name.ToString();
                                            rcmControlFreq.CreatedOn = DateTime.Parse(controlFreq.CreatedOn.ToString());

                                            listControlFreq.Add(rcmControlFreq);
                                        }
                                        rcmQuestionnaire.ListQ13Frequency = listControlFreq;
                                    }
                                    #endregion

                                    #region 14. Is it a key/ non-key control
                                    //Question 14
                                    rcmQuestionnaire.Q14FieldId = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmQuestionnaire").GetSection("Q14").Value);
                                    CategoryItemField catKey = item.Field<CategoryItemField>(rcmQuestionnaire.Q14FieldId);
                                    IEnumerable<CategoryItemField.Option> listCatKey = catKey.Options != null ? catKey.Options : null;
                                    rcmQuestionnaire.Q14Label = catKey.Label; //get label

                                    if (listCatKey != null && listCatKey.Count() > 0) //check if category has value then get all category text
                                    {
                                        List<RcmControlKey> listCategoryKey = new List<RcmControlKey>();
                                        foreach (var itemCatKey in listCatKey.ToList())
                                        {
                                            if (itemCatKey.Status == "active")
                                            {
                                                listCategoryKey.Add(new RcmControlKey { Option = itemCatKey.Text, RcmQuestionnaire = rcmQuestionnaire });
                                            }
                                        }
                                        rcmQuestionnaire.ListQ14ControlKey = listCategoryKey;
                                    }
                                    #endregion

                                    #region 15. What is the control type
                                    //Question 15
                                    rcmQuestionnaire.Q15FieldId = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmQuestionnaire").GetSection("Q15").Value);
                                    AppItemField appControlType = item.Field<AppItemField>(rcmQuestionnaire.Q15FieldId);
                                    IEnumerable<Item> relatedControlTypeItems = appControlType.Items != null ? appControlType.Items : null;
                                    rcmQuestionnaire.Q15Label = appControlType.Label; //get label

                                    if (relatedControlTypeItems != null && relatedControlTypeItems.Count() > 0)
                                    {
                                        List<RcmControlType> listControlType = new List<RcmControlType>();
                                        foreach (var controlType in relatedControlTypeItems.ToList())
                                        {
                                            RcmControlType rcmControlType = new RcmControlType();
                                            rcmControlType.ControlType = controlType.Title;
                                            rcmControlType.PodioItemId = controlType.ItemId;
                                            rcmControlType.PodioUniqueId = controlType.AppItemIdFormatted != null ? controlType.AppItemIdFormatted.ToString() : string.Empty;
                                            rcmControlType.PodioRevision = controlType.CurrentRevision != null ? controlType.CurrentRevision.Revision : 0;
                                            rcmControlType.PodioLink = controlType.Link.ToString();
                                            rcmControlType.CreatedBy = controlType.CreatedBy.Name.ToString();
                                            rcmControlType.CreatedOn = DateTime.Parse(controlType.CreatedOn.ToString());

                                            listControlType.Add(rcmControlType);
                                        }
                                        rcmQuestionnaire.ListQ15ControlType = listControlType;
                                    }
                                    #endregion

                                    #region 16. What is the nature of procedure
                                    //Question 16
                                    rcmQuestionnaire.Q16FieldId = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmQuestionnaire").GetSection("Q16").Value);
                                    AppItemField appNatProc = item.Field<AppItemField>(rcmQuestionnaire.Q16FieldId);
                                    IEnumerable<Item> relatedNatProcItems = appNatProc.Items != null ? appNatProc.Items : null;
                                    rcmQuestionnaire.Q16Label = appNatProc.Label; //get label

                                    if (relatedNatProcItems != null && relatedNatProcItems.Count() > 0)
                                    {
                                        List<RcmNatureProcedure> listNatProc = new List<RcmNatureProcedure>();
                                        foreach (var natureProc in relatedNatProcItems.ToList())
                                        {
                                            RcmNatureProcedure rcmNatProc = new RcmNatureProcedure();
                                            rcmNatProc.NatureProcedure = natureProc.Title;
                                            rcmNatProc.PodioItemId = natureProc.ItemId;
                                            rcmNatProc.PodioUniqueId = natureProc.AppItemIdFormatted != null ? natureProc.AppItemIdFormatted.ToString() : string.Empty;
                                            rcmNatProc.PodioRevision = natureProc.CurrentRevision != null ? natureProc.CurrentRevision.Revision : 0;
                                            rcmNatProc.PodioLink = natureProc.Link.ToString();
                                            rcmNatProc.CreatedBy = natureProc.CreatedBy.Name.ToString();
                                            rcmNatProc.CreatedOn = DateTime.Parse(natureProc.CreatedOn.ToString());

                                            listNatProc.Add(rcmNatProc);
                                        }
                                        rcmQuestionnaire.ListQ16NatureProcedure = listNatProc;
                                    }
                                    #endregion

                                    #region 17. Is it a fraud control
                                    //Question 17
                                    rcmQuestionnaire.Q17FieldId = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmQuestionnaire").GetSection("Q17").Value);
                                    CategoryItemField catFraudControl = item.Field<CategoryItemField>(rcmQuestionnaire.Q17FieldId);
                                    IEnumerable<CategoryItemField.Option> listFraudControl = catFraudControl.Options != null ? catFraudControl.Options : null;
                                    rcmQuestionnaire.Q8Label = catFraudControl.Label; //get label

                                    if (listFraudControl != null && listFraudControl.Count() > 0) //check if category has value then get all category text
                                    {
                                        List<RcmFraudControl> listControlFraud= new List<RcmFraudControl>();
                                        foreach (var itemFraudControl in listFraudControl.ToList())
                                        {
                                            if (itemFraudControl.Status == "active")
                                            {
                                                listControlFraud.Add(new RcmFraudControl { Option = itemFraudControl.Text, RcmQuestionnaire = rcmQuestionnaire });
                                            }
                                        }
                                        rcmQuestionnaire.ListQ17FraudControl = listControlFraud;
                                    }
                                    #endregion

                                    #region 18. What is the risk level
                                    //Question 18
                                    rcmQuestionnaire.Q18FieldId = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmQuestionnaire").GetSection("Q18").Value);
                                    CategoryItemField catRiskControl = item.Field<CategoryItemField>(rcmQuestionnaire.Q18FieldId);
                                    IEnumerable<CategoryItemField.Option> listCatRiskControl = catRiskControl.Options != null ? catRiskControl.Options : null;
                                    rcmQuestionnaire.Q18Label = catRiskControl.Label; //get label

                                    if (listCatRiskControl != null && listCatRiskControl.Count() > 0) //check if category has value then get all category text
                                    {
                                        List<RcmRiskLevel> listRiskLevel = new List<RcmRiskLevel>();
                                        foreach (var itemRiskLevel in listCatRiskControl.ToList())
                                        {
                                            if (itemRiskLevel.Status == "active")
                                            {
                                                listRiskLevel.Add(new RcmRiskLevel { Option = itemRiskLevel.Text, RcmQuestionnaire = rcmQuestionnaire });
                                            }
                                        }
                                        rcmQuestionnaire.ListQ18RiskLevel = listRiskLevel;
                                    }
                                    #endregion

                                    #region 19. Is it a management review control
                                    //Question 19
                                    rcmQuestionnaire.Q19FieldId = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmQuestionnaire").GetSection("Q19").Value);
                                    AppItemField appMgmtRevControl = item.Field<AppItemField>(rcmQuestionnaire.Q19FieldId);
                                    IEnumerable<Item> relatedMgmtRevControlItems = appMgmtRevControl.Items != null ? appMgmtRevControl.Items : null;
                                    rcmQuestionnaire.Q19Label = appMgmtRevControl.Label; //get label

                                    if (relatedMgmtRevControlItems != null && relatedMgmtRevControlItems.Count() > 0)
                                    {
                                        List<RcmManagementReviewControl> listMgtmControl = new List<RcmManagementReviewControl>();
                                        foreach (var controlId in relatedMgmtRevControlItems.ToList())
                                        {
                                            RcmManagementReviewControl rcmMgmtControl = new RcmManagementReviewControl();
                                            rcmMgmtControl.MgmtReviewControl = controlId.Title;
                                            rcmMgmtControl.PodioItemId = controlId.ItemId;
                                            rcmMgmtControl.PodioUniqueId = controlId.AppItemIdFormatted != null ? controlId.AppItemIdFormatted.ToString() : string.Empty;
                                            rcmMgmtControl.PodioRevision = controlId.CurrentRevision != null ? controlId.CurrentRevision.Revision : 0;
                                            rcmMgmtControl.PodioLink = controlId.Link.ToString();
                                            rcmMgmtControl.CreatedBy = controlId.CreatedBy.Name.ToString();
                                            rcmMgmtControl.CreatedOn = DateTime.Parse(controlId.CreatedOn.ToString());

                                            listMgtmControl.Add(rcmMgmtControl);
                                        }
                                        rcmQuestionnaire.ListQ19MgmtReviewControl = listMgtmControl;
                                    }
                                    #endregion

                                    #region 20. What PBC's (supporting documents) are needed to test this control
                                    //Question 20
                                    rcmQuestionnaire.Q20FieldId = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmQuestionnaire").GetSection("Q20").Value);
                                    TextItemField neededPBC = item.Field<TextItemField>(rcmQuestionnaire.Q20FieldId);
                                    rcmQuestionnaire.Q20PBCNeeded = neededPBC.Value != null ? neededPBC.Value.ToString() : string.Empty;
                                    rcmQuestionnaire.Q20PBCNeeded = rcmQuestionnaire.Q20PBCNeeded.Replace("<p>", String.Empty).Replace("</p>", String.Empty);
                                    rcmQuestionnaire.Q20Label = riskPlaceDate.Label;
                                    #endregion

                                    #region 21. What is the testing procedure
                                    //Question 21
                                    rcmQuestionnaire.Q21FieldId = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmQuestionnaire").GetSection("Q21").Value);
                                    TextItemField testProc = item.Field<TextItemField>(rcmQuestionnaire.Q21FieldId);
                                    rcmQuestionnaire.Q21TestProcedure = testProc.Value != null ? testProc.Value.ToString() : string.Empty;
                                    rcmQuestionnaire.Q21TestProcedure = rcmQuestionnaire.Q21TestProcedure.Replace("<p>", String.Empty).Replace("</p>", String.Empty);
                                    rcmQuestionnaire.Q21Label = riskPlaceDate.Label;
                                    #endregion


                                    #region Status
                                    //Status
                                    int fieldStatus = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmQuestionnaire").GetSection("Status").Value);
                                    CategoryItemField catStatus = item.Field<CategoryItemField>(fieldStatus);
                                    IEnumerable<CategoryItemField.Option> listStatus = catStatus.Options != null ? catStatus.Options : null;

                                    if (listStatus != null && listStatus.Count() > 0) //check if category has value then get all category text
                                    {
                                        List<RcmFY> listRcmFy = new List<RcmFY>();
                                        foreach (var itemStatus in listStatus.ToList())
                                        {
                                            if (itemStatus.Status == "active")
                                            {
                                                rcmQuestionnaire.Status = itemStatus.Text;
                                            }
                                        }
                                    }
                                    #endregion

                                    

                                    listRcmQuestionnaire.Add(rcmQuestionnaire);
                                }
                                //pubCollection = collection;
                            }

                            offset += 500;
                        }

                        foreach (var item in listRcmQuestionnaire)
                        {
                            using (var context = _soxContext.Database.BeginTransaction())
                            {
                                //check if podio item id is not zero
                                if (item.PodioItemId != 0)
                                {
                                    //check if already exists
                                    var checkRcmQuestionnaire = _soxContext.RcmQuestionnaire
                                        .Where(x => x.PodioItemId.Equals(item.PodioItemId))
                                        //.Include(x => x.ListQ1Year)
                                        //.Include(x => x.ListQ7FinStatementElement)
                                        //.Include(x => x.ListQ8FinStatementAssert)
                                        //.Include(x => x.ListQ9ControlId)
                                        //.Include(x => x.ListQ12ControlOwner)
                                        //.Include(x => x.ListQ13Frequency)
                                        //.Include(x => x.ListQ14ControlKey)
                                        //.Include(x => x.ListQ15ControlType)
                                        //.Include(x => x.ListQ16NatureProcedure)
                                        //.Include(x => x.ListQ17FraudControl)
                                        //.Include(x => x.ListQ18RiskLevel)
                                        //.Include(x => x.ListQ19MgmtReviewControl)
                                        .FirstOrDefault();
                                    if (checkRcmQuestionnaire != null)
                                    {
                                        item.Id = checkRcmQuestionnaire.Id;
                                        _soxContext.Entry(checkRcmQuestionnaire).CurrentValues.SetValues(item);
                                        await _soxContext.SaveChangesAsync();

                                        #region RcmFY
                                        //Remove RcmFY
                                        var checkQ1Year = _soxContext.RcmFY.Where(x => x.RcmQuestionnaire.Id.Equals(checkRcmQuestionnaire.Id)).ToList();
                                        if (checkQ1Year != null)
                                        {
                                            foreach (var itemYear in checkQ1Year)
                                            {
                                                _soxContext.Remove(itemYear);
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }

                                        //Add RcmFY
                                        if(item.ListQ1Year != null && item.ListQ1Year.Any())
                                        {
                                            item.ListQ1Year = item.ListQ1Year.Select(x => { x.RcmQuestionnaire = checkRcmQuestionnaire; return x; }).ToList();
                                            foreach (var itemYear in item.ListQ1Year)
                                            {

                                                _soxContext.Entry(itemYear).State = EntityState.Added;
                                                _soxContext.Entry(checkRcmQuestionnaire).State = EntityState.Unchanged;
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }
                                        #endregion

                                        #region RcmFinancialStatement
                                        //Remove RcmFinancialStatement
                                        var checkFinStatement = _soxContext.RcmFinancialStatement.Where(x => x.RcmQuestionnaire.Id.Equals(checkRcmQuestionnaire.Id)).ToList();
                                        if (checkFinStatement != null)
                                        {
                                            foreach (var itemVal in checkFinStatement)
                                            {
                                                _soxContext.Remove(itemVal);
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }

                                        //Add RcmFinancialStatement
                                        if (item.ListQ7FinStatementElement != null && item.ListQ7FinStatementElement.Any())
                                        {
                                            item.ListQ7FinStatementElement = item.ListQ7FinStatementElement.Select(x => { x.RcmQuestionnaire = checkRcmQuestionnaire; return x; }).ToList();
                                            foreach (var itemVal in item.ListQ7FinStatementElement)
                                            {

                                                _soxContext.Entry(itemVal).State = EntityState.Added;
                                                _soxContext.Entry(checkRcmQuestionnaire).State = EntityState.Unchanged;
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }
                                        #endregion

                                        #region RcmFinancialStatementAssert
                                        //Remove RcmFinancialStatementAssert
                                        var checkFinStatementAssert = _soxContext.RcmFinancialStatementAssert.Where(x => x.RcmQuestionnaire.Id.Equals(checkRcmQuestionnaire.Id)).ToList();
                                        if (checkFinStatementAssert != null)
                                        {
                                            foreach (var itemVal in checkFinStatementAssert)
                                            {
                                                _soxContext.Remove(itemVal);
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }

                                        //Add RcmFinancialStatementAssert
                                        if (item.ListQ8FinStatementAssert != null && item.ListQ8FinStatementAssert.Any())
                                        {
                                            item.ListQ8FinStatementAssert = item.ListQ8FinStatementAssert.Select(x => { x.RcmQuestionnaire = checkRcmQuestionnaire; return x; }).ToList();
                                            foreach (var itemVal in item.ListQ8FinStatementAssert)
                                            {
                                                _soxContext.Entry(itemVal).State = EntityState.Added;
                                                _soxContext.Entry(checkRcmQuestionnaire).State = EntityState.Unchanged;
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }
                                        #endregion

                                        #region RcmControlId
                                        //Remove RcmControlId
                                        var checkControlId = _soxContext.RcmControlId.Where(x => x.RcmQuestionnaire.Id.Equals(checkRcmQuestionnaire.Id)).ToList();
                                        if (checkControlId != null)
                                        {
                                            foreach (var itemVal in checkControlId)
                                            {
                                                _soxContext.Remove(itemVal);
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }

                                        //Add RcmControlId
                                        if (item.ListQ9ControlId != null && item.ListQ9ControlId.Any())
                                        {
                                            item.ListQ9ControlId = item.ListQ9ControlId.Select(x => { x.RcmQuestionnaire = checkRcmQuestionnaire; return x; }).ToList();
                                            foreach (var itemVal in item.ListQ9ControlId)
                                            {

                                                _soxContext.Entry(itemVal).State = EntityState.Added;
                                                _soxContext.Entry(checkRcmQuestionnaire).State = EntityState.Unchanged;
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }
                                        #endregion

                                        #region RcmControlOwner
                                        //Remove RcmControlOwner
                                        var checkControlOwner = _soxContext.RcmControlOwner.Where(x => x.RcmQuestionnaire.Id.Equals(checkRcmQuestionnaire.Id)).ToList();
                                        if (checkControlOwner != null)
                                        {
                                            foreach (var itemYear in checkControlOwner)
                                            {
                                                _soxContext.Remove(itemYear);
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }

                                        //Add RcmControlOwner
                                        if (item.ListQ12ControlOwner != null && item.ListQ12ControlOwner.Any())
                                        {
                                            item.ListQ12ControlOwner = item.ListQ12ControlOwner.Select(x => { x.RcmQuestionnaire = checkRcmQuestionnaire; return x; }).ToList();
                                            foreach (var itemVal in item.ListQ12ControlOwner)
                                            {

                                                _soxContext.Entry(itemVal).State = EntityState.Added;
                                                _soxContext.Entry(checkRcmQuestionnaire).State = EntityState.Unchanged;
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }
                                        #endregion


                                        #region RcmFrequency
                                        //Remove RcmFrequency
                                        var checkFrequency= _soxContext.RcmFrequency.Where(x => x.RcmQuestionnaire.Id.Equals(checkRcmQuestionnaire.Id)).ToList();
                                        if (checkFrequency != null)
                                        {
                                            foreach (var itemVal in checkFrequency)
                                            {
                                                _soxContext.Remove(itemVal);
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }

                                        //Add RcmFrequency
                                        if (item.ListQ13Frequency != null && item.ListQ13Frequency.Any())
                                        {
                                            item.ListQ13Frequency = item.ListQ13Frequency.Select(x => { x.RcmQuestionnaire = checkRcmQuestionnaire; return x; }).ToList();
                                            foreach (var itemVal in item.ListQ13Frequency)
                                            {

                                                _soxContext.Entry(itemVal).State = EntityState.Added;
                                                _soxContext.Entry(checkRcmQuestionnaire).State = EntityState.Unchanged;
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }
                                        #endregion


                                        #region RcmControlKey
                                        //Remove RcmControlKey
                                        var checkControlKey = _soxContext.RcmControlKey.Where(x => x.RcmQuestionnaire.Id.Equals(checkRcmQuestionnaire.Id)).ToList();
                                        if (checkControlKey != null)
                                        {
                                            foreach (var itemVal in checkControlKey)
                                            {
                                                _soxContext.Remove(itemVal);
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }

                                        //Add RcmControlKey
                                        if (item.ListQ14ControlKey != null && item.ListQ14ControlKey.Any())
                                        {
                                            item.ListQ14ControlKey = item.ListQ14ControlKey.Select(x => { x.RcmQuestionnaire = checkRcmQuestionnaire; return x; }).ToList();
                                            foreach (var itemVal in item.ListQ14ControlKey)
                                            {

                                                _soxContext.Entry(itemVal).State = EntityState.Added;
                                                _soxContext.Entry(checkRcmQuestionnaire).State = EntityState.Unchanged;
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }
                                        #endregion


                                        #region RcmControlType
                                        //Remove RcmControlType
                                        var checkControlType = _soxContext.RcmControlType.Where(x => x.RcmQuestionnaire.Id.Equals(checkRcmQuestionnaire.Id)).ToList();
                                        if (checkControlType != null)
                                        {
                                            foreach (var itemVal in checkControlType)
                                            {
                                                _soxContext.Remove(itemVal);
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }

                                        //Add RcmControlType
                                        if (item.ListQ15ControlType != null && item.ListQ15ControlType.Any())
                                        {
                                            item.ListQ15ControlType = item.ListQ15ControlType.Select(x => { x.RcmQuestionnaire = checkRcmQuestionnaire; return x; }).ToList();
                                            foreach (var itemVal in item.ListQ15ControlType)
                                            {

                                                _soxContext.Entry(itemVal).State = EntityState.Added;
                                                _soxContext.Entry(checkRcmQuestionnaire).State = EntityState.Unchanged;
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }
                                        #endregion


                                        #region RcmNatureProcedure
                                        //Remove RcmNatureProcedure
                                        var checkNatProc = _soxContext.RcmNatureProcedure.Where(x => x.RcmQuestionnaire.Id.Equals(checkRcmQuestionnaire.Id)).ToList();
                                        if (checkNatProc != null)
                                        {
                                            foreach (var itemVal in checkNatProc)
                                            {
                                                _soxContext.Remove(itemVal);
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }

                                        //Add RcmNatureProcedure
                                        if (item.ListQ16NatureProcedure != null && item.ListQ16NatureProcedure.Any())
                                        {
                                            item.ListQ16NatureProcedure = item.ListQ16NatureProcedure.Select(x => { x.RcmQuestionnaire = checkRcmQuestionnaire; return x; }).ToList();
                                            foreach (var itemVal in item.ListQ16NatureProcedure)
                                            {

                                                _soxContext.Entry(itemVal).State = EntityState.Added;
                                                _soxContext.Entry(checkRcmQuestionnaire).State = EntityState.Unchanged;
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }
                                        #endregion


                                        #region RcmFraudControl
                                        //Remove RcmFraudControl
                                        var checkFraudControl = _soxContext.RcmFraudControl.Where(x => x.RcmQuestionnaire.Id.Equals(checkRcmQuestionnaire.Id)).ToList();
                                        if (checkFraudControl != null)
                                        {
                                            foreach (var itemVal in checkFraudControl)
                                            {
                                                _soxContext.Remove(itemVal);
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }

                                        //Add RcmFraudControl
                                        if (item.ListQ17FraudControl != null && item.ListQ17FraudControl.Any())
                                        {
                                            item.ListQ17FraudControl = item.ListQ17FraudControl.Select(x => { x.RcmQuestionnaire = checkRcmQuestionnaire; return x; }).ToList();
                                            foreach (var itemVal in item.ListQ17FraudControl)
                                            {

                                                _soxContext.Entry(itemVal).State = EntityState.Added;
                                                _soxContext.Entry(checkRcmQuestionnaire).State = EntityState.Unchanged;
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }
                                        #endregion


                                        #region RcmRiskLevel
                                        //Remove RcmRiskLevel
                                        var checkRiskLevel = _soxContext.RcmRiskLevel.Where(x => x.RcmQuestionnaire.Id.Equals(checkRcmQuestionnaire.Id)).ToList();
                                        if (checkRiskLevel != null)
                                        {
                                            foreach (var itemVal in checkRiskLevel)
                                            {
                                                _soxContext.Remove(itemVal);
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }

                                        //Add RcmRiskLevel
                                        if (item.ListQ18RiskLevel != null && item.ListQ18RiskLevel.Any())
                                        {
                                            item.ListQ18RiskLevel = item.ListQ18RiskLevel.Select(x => { x.RcmQuestionnaire = checkRcmQuestionnaire; return x; }).ToList();
                                            foreach (var itemVal in item.ListQ18RiskLevel)
                                            {

                                                _soxContext.Entry(itemVal).State = EntityState.Added;
                                                _soxContext.Entry(checkRcmQuestionnaire).State = EntityState.Unchanged;
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }
                                        #endregion


                                        #region RcmManagementReviewControl
                                        //Remove RcmManagementReviewControl
                                        var checkMgmtRevControl = _soxContext.RcmManagementReviewControl.Where(x => x.RcmQuestionnaire.Id.Equals(checkRcmQuestionnaire.Id)).ToList();
                                        if (checkMgmtRevControl != null)
                                        {
                                            foreach (var itemVal in checkMgmtRevControl)
                                            {
                                                _soxContext.Remove(itemVal);
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }

                                        //Add RcmManagementReviewControl
                                        if (item.ListQ19MgmtReviewControl != null && item.ListQ19MgmtReviewControl.Any())
                                        {
                                            item.ListQ19MgmtReviewControl = item.ListQ19MgmtReviewControl.Select(x => { x.RcmQuestionnaire = checkRcmQuestionnaire; return x; }).ToList();
                                            foreach (var itemVal in item.ListQ19MgmtReviewControl)
                                            {

                                                _soxContext.Entry(itemVal).State = EntityState.Added;
                                                _soxContext.Entry(checkRcmQuestionnaire).State = EntityState.Unchanged;
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }
                                        #endregion






                                        //await _soxContext.SaveChangesAsync();
                                        context.Commit();
                                    }
                                    else
                                    {
                                        _soxContext.Add(item);
                                        await _soxContext.SaveChangesAsync();
                                        context.Commit();
                                    }
                                }
                            }

                        }

                        List<int> ItemId = new List<int>();
                        foreach (var item in listRcmQuestionnaire)
                        {
                            ItemId.Add(item.PodioItemId);
                        }

                        return Ok(ItemId);
                    }

                }
                #endregion
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncRcmQuestionnaire");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncRcmQuestionnaire");
                return BadRequest(ex.ToString());
            }

            return NoContent();

        }


        //[AllowAnonymous]
        [HttpPost("soxtracker/questionnaire")]
        public async Task<IActionResult> SyncSoxTrackerQuestionnaireDefine([FromBody] SyncDateRange syncDateRange)
        {
            //Console.WriteLine("Triggered process");
            try
            {
                //Initiliaze
                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();

                //Get API keys in settings.json
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("SoxTrackerApp").GetSection("QuestionnaireAppId").Value;
                PodioAppKey.AppToken = _config.GetSection("SoxTrackerApp").GetSection("QuestionnaireAppToken").Value;

                //connect to podio using api keys and search for item using podio api filter
                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(int.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                if (podio.IsAuthenticated())
                {
                    //podio filter use in querying data
                    var filter = new Dictionary<string, object>
                    {
                        {
                            "last_edit_on",
                            new  { from = syncDateRange.startDate, to = syncDateRange.endDate}
                        },
                    };

                    //get collection
                    PodioCollection<Item> collectionCheck = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 1, 0, filters: filter, sortDesc: true);

                    if (collectionCheck.Filtered != 0)
                    {
                        int loop = (int)Math.Ceiling((double)collectionCheck.Filtered / 500);
                        int offset = 0;
                        WriteLog writeLog = new WriteLog();
                        PodioCollection<Item> pubCollection = new PodioCollection<Item>();
                        List<SoxTrackerQuestionnaire> listSoxTrackerQuestionnaire = new List<SoxTrackerQuestionnaire>();
                        //get filtered items and stored in list
                        for (int i = 0; i < loop; i++)
                        {
                            PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 500, offset, filters: filter, sortDesc: true);
                            //PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 2, offset, filters: filter, sortDesc: true);
                            if (collection.Items != null && collection.Items.Count() > 0)
                            {
                                foreach (var item in collection.Items)
                                {

                                    SoxTrackerQuestionnaire soxTrackerQuestionnaire = new SoxTrackerQuestionnaire();
                                    List<SoxTrackerAppRelationship> listSoxTrackerAppRelationship = new List<SoxTrackerAppRelationship>();
                                    List<SoxTrackerAppCategory> listSoxTrackerAppCategory = new List<SoxTrackerAppCategory>();
                                    soxTrackerQuestionnaire.PodioItemId = (int)item.ItemId;
                                    soxTrackerQuestionnaire.PodioUniqueId = item.AppItemIdFormatted.ToString();
                                    soxTrackerQuestionnaire.PodioRevision = item.CurrentRevision.Revision;
                                    soxTrackerQuestionnaire.PodioLink = item.Link.ToString();
                                    soxTrackerQuestionnaire.CreatedBy = item.CreatedBy.Name.ToString();
                                    soxTrackerQuestionnaire.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());

                                    //foreach (var field in item.Fields.ToList())
                                    //{
                                    //}


                                    #region 6. What PBC's (supporting documents) are needed to test this control?
                                    //Question 6
                                    soxTrackerQuestionnaire.Q6FieldId = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("QuestionnaireFields").GetSection("Q6").Value);
                                    TextItemField riskField = item.Field<TextItemField>(soxTrackerQuestionnaire.Q6FieldId);
                                    soxTrackerQuestionnaire.Q6Label = riskField.Label;
                                    #endregion

                                    #region 7. Who is the PBC Owner
                                    //Question 7
                                    soxTrackerQuestionnaire.Q7FieldId = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("QuestionnaireFields").GetSection("Q7").Value);
                                    AppItemField pbcOwnerField = item.Field<AppItemField>(soxTrackerQuestionnaire.Q7FieldId);
                                    IEnumerable<Item> relatedPbcOwnerField = pbcOwnerField.Items != null ? pbcOwnerField.Items : null;
                                    soxTrackerQuestionnaire.Q7Label = pbcOwnerField.Label; //get label

                                    if (relatedPbcOwnerField != null && relatedPbcOwnerField.Count() > 0)
                                    {
                                        foreach (var finStatement in relatedPbcOwnerField.ToList())
                                        {
                                            SoxTrackerAppRelationship soxTrackerAppRelationship = new SoxTrackerAppRelationship();
                                            soxTrackerAppRelationship.Title = finStatement.Title;
                                            soxTrackerAppRelationship.PodioItemId = finStatement.ItemId;
                                            soxTrackerAppRelationship.PodioUniqueId = finStatement.AppItemIdFormatted != null ? finStatement.AppItemIdFormatted.ToString() : string.Empty;
                                            soxTrackerAppRelationship.PodioRevision = finStatement.CurrentRevision != null ? finStatement.CurrentRevision.Revision : 0;
                                            soxTrackerAppRelationship.PodioLink = finStatement.Link.ToString();
                                            soxTrackerAppRelationship.CreatedBy = finStatement.CreatedBy.Name.ToString();
                                            soxTrackerAppRelationship.CreatedOn = DateTime.Parse(finStatement.CreatedOn.ToString());
                                            soxTrackerAppRelationship.SoxTrackerQuestionnaire = soxTrackerQuestionnaire;
                                            soxTrackerAppRelationship.FieldId = soxTrackerQuestionnaire.Q7FieldId;
                                            listSoxTrackerAppRelationship.Add(soxTrackerAppRelationship);
                                            //soxTrackerQuestionnaire.ListSoxTrackerAppRelationship.Add(soxTrackerAppRelationship);
                                        }
                                    }
                                    #endregion

                                    #region 8. Does the control need a population file request?

                                    //Debug.WriteLine($"field {item.Fields.Label}");

                                    //Question 8
                                    soxTrackerQuestionnaire.Q8FieldId = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("QuestionnaireFields").GetSection("Q8").Value);
                                    CategoryItemField catPopRequired = item.Field<CategoryItemField>(soxTrackerQuestionnaire.Q8FieldId);
                                    IEnumerable<CategoryItemField.Option> listPopRequired = catPopRequired.Options != null ? catPopRequired.Options : null;
                                    soxTrackerQuestionnaire.Q8Label = catPopRequired.Label; //get label

                                    if (listPopRequired != null && listPopRequired.Count() > 0) //check if category has value then get all category text
                                    {
                                        foreach (var itemOption in listPopRequired.ToList())
                                        {
                                            if (itemOption.Status == "active")
                                            {
                                                SoxTrackerAppCategory soxTrackerAppCategory = new SoxTrackerAppCategory();
                                                soxTrackerAppCategory.Option = itemOption.Text;
                                                soxTrackerAppCategory.FieldId = soxTrackerQuestionnaire.Q8FieldId;
                                                soxTrackerAppCategory.SoxTrackerQuestionnaire = soxTrackerQuestionnaire;
                                                listSoxTrackerAppCategory.Add(soxTrackerAppCategory);
                                            }
                                        }
                                    }
                                #endregion

                                    #region 9. Is sample selection/sub-selection required?

                                    //Question 9
                                    soxTrackerQuestionnaire.Q9FieldId = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("QuestionnaireFields").GetSection("Q9").Value);
                                    CategoryItemField catSubSelection = item.Field<CategoryItemField>(soxTrackerQuestionnaire.Q9FieldId);
                                    IEnumerable<CategoryItemField.Option> listSubSelection = catSubSelection.Options != null ? catSubSelection.Options : null;
                                    soxTrackerQuestionnaire.Q9Label = catSubSelection.Label; //get label

                                    if (listSubSelection != null && listSubSelection.Count() > 0) //check if category has value then get all category text
                                    {
                                        foreach (var itemOption in listSubSelection.ToList())
                                        {
                                            if (itemOption.Status == "active")
                                            {
                                                SoxTrackerAppCategory soxTrackerAppCategory = new SoxTrackerAppCategory();
                                                soxTrackerAppCategory.Option = itemOption.Text;
                                                soxTrackerAppCategory.FieldId = soxTrackerQuestionnaire.Q9FieldId;
                                                soxTrackerAppCategory.SoxTrackerQuestionnaire = soxTrackerQuestionnaire;
                                                listSoxTrackerAppCategory.Add(soxTrackerAppCategory);
                                            }
                                        }
                                    }
                                    #endregion

                                    #region 10. Does the external auditor require samples to be tested in R3 (Q4)?

                                    //Question 10
                                    soxTrackerQuestionnaire.Q10FieldId = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("QuestionnaireFields").GetSection("Q10").Value);
                                    CategoryItemField catR3 = item.Field<CategoryItemField>(soxTrackerQuestionnaire.Q10FieldId);
                                    IEnumerable<CategoryItemField.Option> listcatR3 = catR3.Options != null ? catR3.Options : null;
                                    soxTrackerQuestionnaire.Q10Label = catR3.Label; //get label

                                    if (listcatR3 != null && listcatR3.Count() > 0) //check if category has value then get all category text
                                    {
                                        foreach (var itemOption in listcatR3.ToList())
                                        {
                                            if (itemOption.Status == "active")
                                            {
                                                SoxTrackerAppCategory soxTrackerAppCategory = new SoxTrackerAppCategory();
                                                soxTrackerAppCategory.Option = itemOption.Text;
                                                soxTrackerAppCategory.FieldId = soxTrackerQuestionnaire.Q10FieldId;
                                                soxTrackerAppCategory.SoxTrackerQuestionnaire = soxTrackerQuestionnaire;
                                                listSoxTrackerAppCategory.Add(soxTrackerAppCategory);
                                            }
                                        }
                                    }
                                    #endregion

                                    #region 11. How many samples to be tested in R3(Q4)?
                                    ////Question 11
                                    //soxTrackerQuestionnaire.Q11FieldId = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("QuestionnaireFields").GetSection("Q11").Value);
                                    //NumericItemField numR3 = item.Field<NumericItemField>(soxTrackerQuestionnaire.Q11FieldId);
                                    //soxTrackerQuestionnaire.Q11Label = numR3.Label;

                                    //Question 11
                                    soxTrackerQuestionnaire.Q11FieldId = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("QuestionnaireFields").GetSection("Q11").Value);
                                    CategoryItemField catR3Q4 = item.Field<CategoryItemField>(soxTrackerQuestionnaire.Q11FieldId);
                                    IEnumerable<CategoryItemField.Option> listcatR3Q4 = catR3Q4.Options != null ? catR3Q4.Options : null;
                                    soxTrackerQuestionnaire.Q11Label = catR3Q4.Label; //get label

                                    if (listcatR3Q4 != null && listcatR3Q4.Count() > 0) //check if category has value then get all category text
                                    {
                                        foreach (var itemOption in listcatR3Q4.ToList())
                                        {
                                            if (itemOption.Status == "active")
                                            {
                                                SoxTrackerAppCategory soxTrackerAppCategory = new SoxTrackerAppCategory();
                                                soxTrackerAppCategory.Option = itemOption.Text;
                                                soxTrackerAppCategory.FieldId = soxTrackerQuestionnaire.Q11FieldId;
                                                soxTrackerAppCategory.SoxTrackerQuestionnaire = soxTrackerQuestionnaire;
                                                listSoxTrackerAppCategory.Add(soxTrackerAppCategory);
                                            }
                                        }
                                    }

                                    #endregion

                                    #region 12. What is the PBC status (WT)
                                    //Question 12
                                    soxTrackerQuestionnaire.Q12AFieldId = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("QuestionnaireFields").GetSection("Q12A").Value);
                                    AppItemField pbcStatusWT = item.Field<AppItemField>(soxTrackerQuestionnaire.Q12AFieldId);
                                    IEnumerable<Item> relatedPbcStatusWT = pbcStatusWT.Items != null ? pbcStatusWT.Items : null;
                                    soxTrackerQuestionnaire.Q12ALabel = pbcStatusWT.Label; //get label

                                    if (relatedPbcStatusWT != null && relatedPbcStatusWT.Count() > 0)
                                    {
                                        foreach (var itemApp in relatedPbcStatusWT.ToList())
                                        {
                                            SoxTrackerAppRelationship soxTrackerAppRelationship = new SoxTrackerAppRelationship();
                                            soxTrackerAppRelationship.Title = itemApp.Title;
                                            soxTrackerAppRelationship.PodioItemId = itemApp.ItemId;
                                            soxTrackerAppRelationship.PodioUniqueId = itemApp.AppItemIdFormatted != null ? itemApp.AppItemIdFormatted.ToString() : string.Empty;
                                            soxTrackerAppRelationship.PodioRevision = itemApp.CurrentRevision != null ? itemApp.CurrentRevision.Revision : 0;
                                            soxTrackerAppRelationship.PodioLink = itemApp.Link.ToString();
                                            soxTrackerAppRelationship.CreatedBy = itemApp.CreatedBy.Name.ToString();
                                            soxTrackerAppRelationship.CreatedOn = DateTime.Parse(itemApp.CreatedOn.ToString());
                                            soxTrackerAppRelationship.SoxTrackerQuestionnaire = soxTrackerQuestionnaire;
                                            soxTrackerAppRelationship.FieldId = soxTrackerQuestionnaire.Q12AFieldId;
                                            listSoxTrackerAppRelationship.Add(soxTrackerAppRelationship);
                                            //soxTrackerQuestionnaire.ListSoxTrackerAppRelationship.Add(soxTrackerAppRelationship);
                                        }
                                    }
                                    #endregion

                                    #region 12. What is the PBC status (R1)
                                    //Question 12
                                    soxTrackerQuestionnaire.Q12BFieldId = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("QuestionnaireFields").GetSection("Q12B").Value);
                                    AppItemField pbcStatusR1 = item.Field<AppItemField>(soxTrackerQuestionnaire.Q12BFieldId);
                                    IEnumerable<Item> relatedPbcStatusR1 = pbcStatusR1.Items != null ? pbcStatusR1.Items : null;
                                    soxTrackerQuestionnaire.Q12BLabel = pbcStatusR1.Label; //get label

                                    if (relatedPbcStatusR1 != null && relatedPbcStatusR1.Count() > 0)
                                    {
                                        foreach (var itemApp in relatedPbcStatusR1.ToList())
                                        {
                                            SoxTrackerAppRelationship soxTrackerAppRelationship = new SoxTrackerAppRelationship();
                                            soxTrackerAppRelationship.Title = itemApp.Title;
                                            soxTrackerAppRelationship.PodioItemId = itemApp.ItemId;
                                            soxTrackerAppRelationship.PodioUniqueId = itemApp.AppItemIdFormatted != null ? itemApp.AppItemIdFormatted.ToString() : string.Empty;
                                            soxTrackerAppRelationship.PodioRevision = itemApp.CurrentRevision != null ? itemApp.CurrentRevision.Revision : 0;
                                            soxTrackerAppRelationship.PodioLink = itemApp.Link.ToString();
                                            soxTrackerAppRelationship.CreatedBy = itemApp.CreatedBy.Name.ToString();
                                            soxTrackerAppRelationship.CreatedOn = DateTime.Parse(itemApp.CreatedOn.ToString());
                                            soxTrackerAppRelationship.SoxTrackerQuestionnaire = soxTrackerQuestionnaire;
                                            soxTrackerAppRelationship.FieldId = soxTrackerQuestionnaire.Q12BFieldId;
                                            listSoxTrackerAppRelationship.Add(soxTrackerAppRelationship);
                                            //soxTrackerQuestionnaire.ListSoxTrackerAppRelationship.Add(soxTrackerAppRelationship);
                                        }
                                    }
                                    #endregion

                                    #region 12. What is the PBC status (R2)
                                    //Question 12
                                    soxTrackerQuestionnaire.Q12CFieldId = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("QuestionnaireFields").GetSection("Q12C").Value);
                                    AppItemField pbcStatusR2 = item.Field<AppItemField>(soxTrackerQuestionnaire.Q12CFieldId);
                                    IEnumerable<Item> relatedPbcStatusR2 = pbcStatusR2.Items != null ? pbcStatusR2.Items : null;
                                    soxTrackerQuestionnaire.Q12CLabel = pbcStatusR2.Label; //get label

                                    if (relatedPbcStatusR2 != null && relatedPbcStatusR2.Count() > 0)
                                    {
                                        foreach (var itemApp in relatedPbcStatusR2.ToList())
                                        {
                                            SoxTrackerAppRelationship soxTrackerAppRelationship = new SoxTrackerAppRelationship();
                                            soxTrackerAppRelationship.Title = itemApp.Title;
                                            soxTrackerAppRelationship.PodioItemId = itemApp.ItemId;
                                            soxTrackerAppRelationship.PodioUniqueId = itemApp.AppItemIdFormatted != null ? itemApp.AppItemIdFormatted.ToString() : string.Empty;
                                            soxTrackerAppRelationship.PodioRevision = itemApp.CurrentRevision != null ? itemApp.CurrentRevision.Revision : 0;
                                            soxTrackerAppRelationship.PodioLink = itemApp.Link.ToString();
                                            soxTrackerAppRelationship.CreatedBy = itemApp.CreatedBy.Name.ToString();
                                            soxTrackerAppRelationship.CreatedOn = DateTime.Parse(itemApp.CreatedOn.ToString());
                                            soxTrackerAppRelationship.SoxTrackerQuestionnaire = soxTrackerQuestionnaire;
                                            soxTrackerAppRelationship.FieldId = soxTrackerQuestionnaire.Q12CFieldId;
                                            listSoxTrackerAppRelationship.Add(soxTrackerAppRelationship);
                                            //soxTrackerQuestionnaire.ListSoxTrackerAppRelationship.Add(soxTrackerAppRelationship);
                                        }
                                    }
                                    #endregion

                                    #region 12. What is the PBC status (R3)
                                    //Question 12
                                    soxTrackerQuestionnaire.Q12DFieldId = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("QuestionnaireFields").GetSection("Q12D").Value);
                                    AppItemField pbcStatusR3 = item.Field<AppItemField>(soxTrackerQuestionnaire.Q12DFieldId);
                                    IEnumerable<Item> relatedPbcStatusR3 = pbcStatusR3.Items != null ? pbcStatusR3.Items : null;
                                    soxTrackerQuestionnaire.Q12DLabel = pbcStatusR3.Label; //get label

                                    if (relatedPbcStatusR3 != null && relatedPbcStatusR3.Count() > 0)
                                    {
                                        foreach (var itemApp in relatedPbcStatusR3.ToList())
                                        {
                                            SoxTrackerAppRelationship soxTrackerAppRelationship = new SoxTrackerAppRelationship();
                                            soxTrackerAppRelationship.Title = itemApp.Title;
                                            soxTrackerAppRelationship.PodioItemId = itemApp.ItemId;
                                            soxTrackerAppRelationship.PodioUniqueId = itemApp.AppItemIdFormatted != null ? itemApp.AppItemIdFormatted.ToString() : string.Empty;
                                            soxTrackerAppRelationship.PodioRevision = itemApp.CurrentRevision != null ? itemApp.CurrentRevision.Revision : 0;
                                            soxTrackerAppRelationship.PodioLink = itemApp.Link.ToString();
                                            soxTrackerAppRelationship.CreatedBy = itemApp.CreatedBy.Name.ToString();
                                            soxTrackerAppRelationship.CreatedOn = DateTime.Parse(itemApp.CreatedOn.ToString());
                                            soxTrackerAppRelationship.SoxTrackerQuestionnaire = soxTrackerQuestionnaire;
                                            soxTrackerAppRelationship.FieldId = soxTrackerQuestionnaire.Q12DFieldId;
                                            listSoxTrackerAppRelationship.Add(soxTrackerAppRelationship);
                                            //soxTrackerQuestionnaire.ListSoxTrackerAppRelationship.Add(soxTrackerAppRelationship);
                                        }
                                    }
                                    #endregion

                                    #region 13. Testing Assignment (WT)
                                    //Question 13
                                    soxTrackerQuestionnaire.Q13AFieldId = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("QuestionnaireFields").GetSection("Q13A").Value);
                                    AppItemField testingAssignWT = item.Field<AppItemField>(soxTrackerQuestionnaire.Q13AFieldId);
                                    IEnumerable<Item> relatedTestAssignWT = testingAssignWT.Items != null ? testingAssignWT.Items : null;
                                    soxTrackerQuestionnaire.Q13ALabel = testingAssignWT.Label; //get label

                                    if (relatedTestAssignWT != null && relatedTestAssignWT.Count() > 0)
                                    {
                                        foreach (var itemApp in relatedTestAssignWT.ToList())
                                        {
                                            SoxTrackerAppRelationship soxTrackerAppRelationship = new SoxTrackerAppRelationship();
                                            soxTrackerAppRelationship.Title = itemApp.Title;
                                            soxTrackerAppRelationship.PodioItemId = itemApp.ItemId;
                                            soxTrackerAppRelationship.PodioUniqueId = itemApp.AppItemIdFormatted != null ? itemApp.AppItemIdFormatted.ToString() : string.Empty;
                                            soxTrackerAppRelationship.PodioRevision = itemApp.CurrentRevision != null ? itemApp.CurrentRevision.Revision : 0;
                                            soxTrackerAppRelationship.PodioLink = itemApp.Link.ToString();
                                            soxTrackerAppRelationship.CreatedBy = itemApp.CreatedBy.Name.ToString();
                                            soxTrackerAppRelationship.CreatedOn = DateTime.Parse(itemApp.CreatedOn.ToString());
                                            soxTrackerAppRelationship.SoxTrackerQuestionnaire = soxTrackerQuestionnaire;
                                            soxTrackerAppRelationship.FieldId = soxTrackerQuestionnaire.Q13AFieldId;
                                            listSoxTrackerAppRelationship.Add(soxTrackerAppRelationship);
                                            //soxTrackerQuestionnaire.ListSoxTrackerAppRelationship.Add(soxTrackerAppRelationship);
                                        }
                                    }
                                    #endregion

                                    #region 13. Testing Assignment (R1)
                                    //Question 13
                                    soxTrackerQuestionnaire.Q13BFieldId = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("QuestionnaireFields").GetSection("Q13B").Value);
                                    AppItemField testingAssignR1 = item.Field<AppItemField>(soxTrackerQuestionnaire.Q13BFieldId);
                                    IEnumerable<Item> relatedTestAssignR1 = testingAssignR1.Items != null ? testingAssignR1.Items : null;
                                    soxTrackerQuestionnaire.Q13BLabel = testingAssignR1.Label; //get label

                                    if (relatedTestAssignR1 != null && relatedTestAssignR1.Count() > 0)
                                    {
                                        foreach (var itemApp in relatedTestAssignR1.ToList())
                                        {
                                            SoxTrackerAppRelationship soxTrackerAppRelationship = new SoxTrackerAppRelationship();
                                            soxTrackerAppRelationship.Title = itemApp.Title;
                                            soxTrackerAppRelationship.PodioItemId = itemApp.ItemId;
                                            soxTrackerAppRelationship.PodioUniqueId = itemApp.AppItemIdFormatted != null ? itemApp.AppItemIdFormatted.ToString() : string.Empty;
                                            soxTrackerAppRelationship.PodioRevision = itemApp.CurrentRevision != null ? itemApp.CurrentRevision.Revision : 0;
                                            soxTrackerAppRelationship.PodioLink = itemApp.Link.ToString();
                                            soxTrackerAppRelationship.CreatedBy = itemApp.CreatedBy.Name.ToString();
                                            soxTrackerAppRelationship.CreatedOn = DateTime.Parse(itemApp.CreatedOn.ToString());
                                            soxTrackerAppRelationship.SoxTrackerQuestionnaire = soxTrackerQuestionnaire;
                                            soxTrackerAppRelationship.FieldId = soxTrackerQuestionnaire.Q13BFieldId;
                                            listSoxTrackerAppRelationship.Add(soxTrackerAppRelationship);
                                        }
                                    }
                                    #endregion

                                    #region 13. Testing Assignment (R2)
                                    //Question 13
                                    soxTrackerQuestionnaire.Q13CFieldId = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("QuestionnaireFields").GetSection("Q13C").Value);
                                    AppItemField testingAssignR2 = item.Field<AppItemField>(soxTrackerQuestionnaire.Q13CFieldId);
                                    IEnumerable<Item> relatedTestAssignR2 = testingAssignR2.Items != null ? testingAssignR2.Items : null;
                                    soxTrackerQuestionnaire.Q13CLabel = testingAssignR2.Label; //get label

                                    if (relatedTestAssignR2 != null && relatedTestAssignR2.Count() > 0)
                                    {
                                        foreach (var itemApp in relatedTestAssignR2.ToList())
                                        {
                                            SoxTrackerAppRelationship soxTrackerAppRelationship = new SoxTrackerAppRelationship();
                                            soxTrackerAppRelationship.Title = itemApp.Title;
                                            soxTrackerAppRelationship.PodioItemId = itemApp.ItemId;
                                            soxTrackerAppRelationship.PodioUniqueId = itemApp.AppItemIdFormatted != null ? itemApp.AppItemIdFormatted.ToString() : string.Empty;
                                            soxTrackerAppRelationship.PodioRevision = itemApp.CurrentRevision != null ? itemApp.CurrentRevision.Revision : 0;
                                            soxTrackerAppRelationship.PodioLink = itemApp.Link.ToString();
                                            soxTrackerAppRelationship.CreatedBy = itemApp.CreatedBy.Name.ToString();
                                            soxTrackerAppRelationship.CreatedOn = DateTime.Parse(itemApp.CreatedOn.ToString());
                                            soxTrackerAppRelationship.SoxTrackerQuestionnaire = soxTrackerQuestionnaire;
                                            soxTrackerAppRelationship.FieldId = soxTrackerQuestionnaire.Q13CFieldId;
                                            listSoxTrackerAppRelationship.Add(soxTrackerAppRelationship);
                                            //soxTrackerQuestionnaire.ListSoxTrackerAppRelationship.Add(soxTrackerAppRelationship);
                                        }
                                    }
                                    #endregion

                                    #region 13. Testing Assignment (R3)
                                    //Question R3
                                    soxTrackerQuestionnaire.Q13DFieldId = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("QuestionnaireFields").GetSection("Q13D").Value);
                                    AppItemField testingAssignR3 = item.Field<AppItemField>(soxTrackerQuestionnaire.Q13DFieldId);
                                    IEnumerable<Item> relatedTestAssignR3 = testingAssignR3.Items != null ? testingAssignR3.Items : null;
                                    soxTrackerQuestionnaire.Q13DLabel = testingAssignR3.Label; //get label

                                    if (relatedTestAssignR3 != null && relatedTestAssignR3.Count() > 0)
                                    {
                                        foreach (var itemApp in relatedTestAssignR3.ToList())
                                        {
                                            SoxTrackerAppRelationship soxTrackerAppRelationship = new SoxTrackerAppRelationship();
                                            soxTrackerAppRelationship.Title = itemApp.Title;
                                            soxTrackerAppRelationship.PodioItemId = itemApp.ItemId;
                                            soxTrackerAppRelationship.PodioUniqueId = itemApp.AppItemIdFormatted != null ? itemApp.AppItemIdFormatted.ToString() : string.Empty;
                                            soxTrackerAppRelationship.PodioRevision = itemApp.CurrentRevision != null ? itemApp.CurrentRevision.Revision : 0;
                                            soxTrackerAppRelationship.PodioLink = itemApp.Link.ToString();
                                            soxTrackerAppRelationship.CreatedBy = itemApp.CreatedBy.Name.ToString();
                                            soxTrackerAppRelationship.CreatedOn = DateTime.Parse(itemApp.CreatedOn.ToString());
                                            soxTrackerAppRelationship.SoxTrackerQuestionnaire = soxTrackerQuestionnaire;
                                            soxTrackerAppRelationship.FieldId = soxTrackerQuestionnaire.Q13DFieldId;
                                            listSoxTrackerAppRelationship.Add(soxTrackerAppRelationship);
                                        }
                                    }
                                    #endregion

                                    #region 13. 1st Level Reviewer WT
                                    //Question 13
                                    soxTrackerQuestionnaire.Q13EFieldId = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("QuestionnaireFields").GetSection("Q13E").Value);
                                    AppItemField firstLevRevWT = item.Field<AppItemField>(soxTrackerQuestionnaire.Q13EFieldId);
                                    IEnumerable<Item> relatedFirstLevRevWT = firstLevRevWT.Items != null ? firstLevRevWT.Items : null;
                                    soxTrackerQuestionnaire.Q13ELabel = firstLevRevWT.Label; //get label

                                    if (relatedFirstLevRevWT != null && relatedFirstLevRevWT.Count() > 0)
                                    {
                                        foreach (var itemApp in relatedFirstLevRevWT.ToList())
                                        {
                                            SoxTrackerAppRelationship soxTrackerAppRelationship = new SoxTrackerAppRelationship();
                                            soxTrackerAppRelationship.Title = itemApp.Title;
                                            soxTrackerAppRelationship.PodioItemId = itemApp.ItemId;
                                            soxTrackerAppRelationship.PodioUniqueId = itemApp.AppItemIdFormatted != null ? itemApp.AppItemIdFormatted.ToString() : string.Empty;
                                            soxTrackerAppRelationship.PodioRevision = itemApp.CurrentRevision != null ? itemApp.CurrentRevision.Revision : 0;
                                            soxTrackerAppRelationship.PodioLink = itemApp.Link.ToString();
                                            soxTrackerAppRelationship.CreatedBy = itemApp.CreatedBy.Name.ToString();
                                            soxTrackerAppRelationship.CreatedOn = DateTime.Parse(itemApp.CreatedOn.ToString());
                                            soxTrackerAppRelationship.SoxTrackerQuestionnaire = soxTrackerQuestionnaire;
                                            soxTrackerAppRelationship.FieldId = soxTrackerQuestionnaire.Q13EFieldId;
                                            listSoxTrackerAppRelationship.Add(soxTrackerAppRelationship);
                                            //soxTrackerQuestionnaire.ListSoxTrackerAppRelationship.Add(soxTrackerAppRelationship);
                                        }
                                    }
                                    #endregion

                                    #region 13. 1st Level Reviewer R1
                                    //Question 13
                                    soxTrackerQuestionnaire.Q13FFieldId = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("QuestionnaireFields").GetSection("Q13F").Value);
                                    AppItemField firstLevRevR1 = item.Field<AppItemField>(soxTrackerQuestionnaire.Q13FFieldId);
                                    IEnumerable<Item> relatedFirstLevRevR1 = firstLevRevR1.Items != null ? firstLevRevR1.Items : null;
                                    soxTrackerQuestionnaire.Q13FLabel = firstLevRevR1.Label; //get label

                                    if (relatedFirstLevRevR1 != null && relatedFirstLevRevR1.Count() > 0)
                                    {
                                        foreach (var itemApp in relatedFirstLevRevR1.ToList())
                                        {
                                            SoxTrackerAppRelationship soxTrackerAppRelationship = new SoxTrackerAppRelationship();
                                            soxTrackerAppRelationship.Title = itemApp.Title;
                                            soxTrackerAppRelationship.PodioItemId = itemApp.ItemId;
                                            soxTrackerAppRelationship.PodioUniqueId = itemApp.AppItemIdFormatted != null ? itemApp.AppItemIdFormatted.ToString() : string.Empty;
                                            soxTrackerAppRelationship.PodioRevision = itemApp.CurrentRevision != null ? itemApp.CurrentRevision.Revision : 0;
                                            soxTrackerAppRelationship.PodioLink = itemApp.Link.ToString();
                                            soxTrackerAppRelationship.CreatedBy = itemApp.CreatedBy.Name.ToString();
                                            soxTrackerAppRelationship.CreatedOn = DateTime.Parse(itemApp.CreatedOn.ToString());
                                            soxTrackerAppRelationship.SoxTrackerQuestionnaire = soxTrackerQuestionnaire;
                                            soxTrackerAppRelationship.FieldId = soxTrackerQuestionnaire.Q13FFieldId;
                                            listSoxTrackerAppRelationship.Add(soxTrackerAppRelationship);
                                        }
                                    }
                                    #endregion

                                    #region 13. 1st Level Reviewer R2
                                    //Question 13
                                    soxTrackerQuestionnaire.Q13GFieldId = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("QuestionnaireFields").GetSection("Q13G").Value);
                                    AppItemField firstLevRevR2 = item.Field<AppItemField>(soxTrackerQuestionnaire.Q13GFieldId);
                                    IEnumerable<Item> relatedFirstLevRevR2 = firstLevRevR2.Items != null ? firstLevRevR2.Items : null;
                                    soxTrackerQuestionnaire.Q13GLabel = firstLevRevR2.Label; //get label

                                    if (relatedFirstLevRevR2 != null && relatedFirstLevRevR2.Count() > 0)
                                    {
                                        foreach (var itemApp in relatedFirstLevRevR2.ToList())
                                        {
                                            SoxTrackerAppRelationship soxTrackerAppRelationship = new SoxTrackerAppRelationship();
                                            soxTrackerAppRelationship.Title = itemApp.Title;
                                            soxTrackerAppRelationship.PodioItemId = itemApp.ItemId;
                                            soxTrackerAppRelationship.PodioUniqueId = itemApp.AppItemIdFormatted != null ? itemApp.AppItemIdFormatted.ToString() : string.Empty;
                                            soxTrackerAppRelationship.PodioRevision = itemApp.CurrentRevision != null ? itemApp.CurrentRevision.Revision : 0;
                                            soxTrackerAppRelationship.PodioLink = itemApp.Link.ToString();
                                            soxTrackerAppRelationship.CreatedBy = itemApp.CreatedBy.Name.ToString();
                                            soxTrackerAppRelationship.CreatedOn = DateTime.Parse(itemApp.CreatedOn.ToString());
                                            soxTrackerAppRelationship.SoxTrackerQuestionnaire = soxTrackerQuestionnaire;
                                            soxTrackerAppRelationship.FieldId = soxTrackerQuestionnaire.Q13GFieldId;
                                            listSoxTrackerAppRelationship.Add(soxTrackerAppRelationship);
                                            //soxTrackerQuestionnaire.ListSoxTrackerAppRelationship.Add(soxTrackerAppRelationship);
                                        }
                                    }
                                    #endregion

                                    #region 13. 1st Level Reviewer R3
                                    //Question R3
                                    soxTrackerQuestionnaire.Q13HFieldId = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("QuestionnaireFields").GetSection("Q13H").Value);
                                    AppItemField firstLevRevR3 = item.Field<AppItemField>(soxTrackerQuestionnaire.Q13HFieldId);
                                    IEnumerable<Item> relatedFirstLevRevR3 = firstLevRevR3.Items != null ? firstLevRevR3.Items : null;
                                    soxTrackerQuestionnaire.Q13HLabel = firstLevRevR3.Label; //get label

                                    if (relatedFirstLevRevR3 != null && relatedFirstLevRevR3.Count() > 0)
                                    {
                                        foreach (var itemApp in relatedFirstLevRevR3.ToList())
                                        {
                                            SoxTrackerAppRelationship soxTrackerAppRelationship = new SoxTrackerAppRelationship();
                                            soxTrackerAppRelationship.Title = itemApp.Title;
                                            soxTrackerAppRelationship.PodioItemId = itemApp.ItemId;
                                            soxTrackerAppRelationship.PodioUniqueId = itemApp.AppItemIdFormatted != null ? itemApp.AppItemIdFormatted.ToString() : string.Empty;
                                            soxTrackerAppRelationship.PodioRevision = itemApp.CurrentRevision != null ? itemApp.CurrentRevision.Revision : 0;
                                            soxTrackerAppRelationship.PodioLink = itemApp.Link.ToString();
                                            soxTrackerAppRelationship.CreatedBy = itemApp.CreatedBy.Name.ToString();
                                            soxTrackerAppRelationship.CreatedOn = DateTime.Parse(itemApp.CreatedOn.ToString());
                                            soxTrackerAppRelationship.SoxTrackerQuestionnaire = soxTrackerQuestionnaire;
                                            soxTrackerAppRelationship.FieldId = soxTrackerQuestionnaire.Q13HFieldId;
                                            listSoxTrackerAppRelationship.Add(soxTrackerAppRelationship);
                                        }
                                    }
                                    #endregion

                                    #region 13. 2nd Level Reviewer WT
                                    //Question 13
                                    soxTrackerQuestionnaire.Q13IFieldId = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("QuestionnaireFields").GetSection("Q13I").Value);
                                    AppItemField secondLevRevWT = item.Field<AppItemField>(soxTrackerQuestionnaire.Q13IFieldId);
                                    IEnumerable<Item> relatedSecondLevRevWT = secondLevRevWT.Items != null ? secondLevRevWT.Items : null;
                                    soxTrackerQuestionnaire.Q13ILabel = secondLevRevWT.Label; //get label

                                    if (relatedSecondLevRevWT != null && relatedSecondLevRevWT.Count() > 0)
                                    {
                                        foreach (var itemApp in relatedSecondLevRevWT.ToList())
                                        {
                                            SoxTrackerAppRelationship soxTrackerAppRelationship = new SoxTrackerAppRelationship();
                                            soxTrackerAppRelationship.Title = itemApp.Title;
                                            soxTrackerAppRelationship.PodioItemId = itemApp.ItemId;
                                            soxTrackerAppRelationship.PodioUniqueId = itemApp.AppItemIdFormatted != null ? itemApp.AppItemIdFormatted.ToString() : string.Empty;
                                            soxTrackerAppRelationship.PodioRevision = itemApp.CurrentRevision != null ? itemApp.CurrentRevision.Revision : 0;
                                            soxTrackerAppRelationship.PodioLink = itemApp.Link.ToString();
                                            soxTrackerAppRelationship.CreatedBy = itemApp.CreatedBy.Name.ToString();
                                            soxTrackerAppRelationship.CreatedOn = DateTime.Parse(itemApp.CreatedOn.ToString());
                                            soxTrackerAppRelationship.SoxTrackerQuestionnaire = soxTrackerQuestionnaire;
                                            soxTrackerAppRelationship.FieldId = soxTrackerQuestionnaire.Q13IFieldId;
                                            listSoxTrackerAppRelationship.Add(soxTrackerAppRelationship);
                                            //soxTrackerQuestionnaire.ListSoxTrackerAppRelationship.Add(soxTrackerAppRelationship);
                                        }
                                    }
                                    #endregion

                                    #region 13. 2nd Level Reviewer R1
                                    //Question 13
                                    soxTrackerQuestionnaire.Q13JFieldId = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("QuestionnaireFields").GetSection("Q13J").Value);
                                    AppItemField secondLevRevR1 = item.Field<AppItemField>(soxTrackerQuestionnaire.Q13JFieldId);
                                    IEnumerable<Item> relatedSecondLevRevR1 = secondLevRevR1.Items != null ? secondLevRevR1.Items : null;
                                    soxTrackerQuestionnaire.Q13JLabel = secondLevRevR1.Label; //get label

                                    if (relatedSecondLevRevR1 != null && relatedSecondLevRevR1.Count() > 0)
                                    {
                                        foreach (var itemApp in relatedSecondLevRevR1.ToList())
                                        {
                                            SoxTrackerAppRelationship soxTrackerAppRelationship = new SoxTrackerAppRelationship();
                                            soxTrackerAppRelationship.Title = itemApp.Title;
                                            soxTrackerAppRelationship.PodioItemId = itemApp.ItemId;
                                            soxTrackerAppRelationship.PodioUniqueId = itemApp.AppItemIdFormatted != null ? itemApp.AppItemIdFormatted.ToString() : string.Empty;
                                            soxTrackerAppRelationship.PodioRevision = itemApp.CurrentRevision != null ? itemApp.CurrentRevision.Revision : 0;
                                            soxTrackerAppRelationship.PodioLink = itemApp.Link.ToString();
                                            soxTrackerAppRelationship.CreatedBy = itemApp.CreatedBy.Name.ToString();
                                            soxTrackerAppRelationship.CreatedOn = DateTime.Parse(itemApp.CreatedOn.ToString());
                                            soxTrackerAppRelationship.SoxTrackerQuestionnaire = soxTrackerQuestionnaire;
                                            soxTrackerAppRelationship.FieldId = soxTrackerQuestionnaire.Q13JFieldId;
                                            listSoxTrackerAppRelationship.Add(soxTrackerAppRelationship);
                                        }
                                    }
                                    #endregion

                                    #region 13. 2nd Level Reviewer R2
                                    //Question 13
                                    soxTrackerQuestionnaire.Q13KFieldId = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("QuestionnaireFields").GetSection("Q13K").Value);
                                    AppItemField secondLevRevR2 = item.Field<AppItemField>(soxTrackerQuestionnaire.Q13KFieldId);
                                    IEnumerable<Item> relatedSecondLevRevR2 = secondLevRevR2.Items != null ? secondLevRevR2.Items : null;
                                    soxTrackerQuestionnaire.Q13KLabel = secondLevRevR2.Label; //get label

                                    if (relatedSecondLevRevR2 != null && relatedSecondLevRevR2.Count() > 0)
                                    {
                                        foreach (var itemApp in relatedSecondLevRevR2.ToList())
                                        {
                                            SoxTrackerAppRelationship soxTrackerAppRelationship = new SoxTrackerAppRelationship();
                                            soxTrackerAppRelationship.Title = itemApp.Title;
                                            soxTrackerAppRelationship.PodioItemId = itemApp.ItemId;
                                            soxTrackerAppRelationship.PodioUniqueId = itemApp.AppItemIdFormatted != null ? itemApp.AppItemIdFormatted.ToString() : string.Empty;
                                            soxTrackerAppRelationship.PodioRevision = itemApp.CurrentRevision != null ? itemApp.CurrentRevision.Revision : 0;
                                            soxTrackerAppRelationship.PodioLink = itemApp.Link.ToString();
                                            soxTrackerAppRelationship.CreatedBy = itemApp.CreatedBy.Name.ToString();
                                            soxTrackerAppRelationship.CreatedOn = DateTime.Parse(itemApp.CreatedOn.ToString());
                                            soxTrackerAppRelationship.SoxTrackerQuestionnaire = soxTrackerQuestionnaire;
                                            soxTrackerAppRelationship.FieldId = soxTrackerQuestionnaire.Q13KFieldId;
                                            listSoxTrackerAppRelationship.Add(soxTrackerAppRelationship);
                                            //soxTrackerQuestionnaire.ListSoxTrackerAppRelationship.Add(soxTrackerAppRelationship);
                                        }
                                    }
                                    #endregion

                                    #region 13. 2nd Level Reviewer R3
                                    //Question R3
                                    soxTrackerQuestionnaire.Q13LFieldId = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("QuestionnaireFields").GetSection("Q13L").Value);
                                    AppItemField secondLevRevR3 = item.Field<AppItemField>(soxTrackerQuestionnaire.Q13LFieldId);
                                    IEnumerable<Item> relatedSecondLevRevR3 = secondLevRevR3.Items != null ? secondLevRevR3.Items : null;
                                    soxTrackerQuestionnaire.Q13LLabel = secondLevRevR3.Label; //get label

                                    if (relatedSecondLevRevR3 != null && relatedSecondLevRevR3.Count() > 0)
                                    {
                                        foreach (var itemApp in relatedSecondLevRevR3.ToList())
                                        {
                                            SoxTrackerAppRelationship soxTrackerAppRelationship = new SoxTrackerAppRelationship();
                                            soxTrackerAppRelationship.Title = itemApp.Title;
                                            soxTrackerAppRelationship.PodioItemId = itemApp.ItemId;
                                            soxTrackerAppRelationship.PodioUniqueId = itemApp.AppItemIdFormatted != null ? itemApp.AppItemIdFormatted.ToString() : string.Empty;
                                            soxTrackerAppRelationship.PodioRevision = itemApp.CurrentRevision != null ? itemApp.CurrentRevision.Revision : 0;
                                            soxTrackerAppRelationship.PodioLink = itemApp.Link.ToString();
                                            soxTrackerAppRelationship.CreatedBy = itemApp.CreatedBy.Name.ToString();
                                            soxTrackerAppRelationship.CreatedOn = DateTime.Parse(itemApp.CreatedOn.ToString());
                                            soxTrackerAppRelationship.SoxTrackerQuestionnaire = soxTrackerQuestionnaire;
                                            soxTrackerAppRelationship.FieldId = soxTrackerQuestionnaire.Q13LFieldId;
                                            listSoxTrackerAppRelationship.Add(soxTrackerAppRelationship);
                                        }
                                    }
                                    #endregion

                                    #region 14. Reviewer checklist required?

                                    //Question 14
                                    soxTrackerQuestionnaire.Q14AFieldId = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("QuestionnaireFields").GetSection("Q14A").Value);
                                    CategoryItemField catRevCheckRequired = item.Field<CategoryItemField>(soxTrackerQuestionnaire.Q14AFieldId);
                                    IEnumerable<CategoryItemField.Option> listRevCheckRequired = catRevCheckRequired.Options != null ? catRevCheckRequired.Options : null;
                                    soxTrackerQuestionnaire.Q14ALabel = catRevCheckRequired.Label; //get label

                                    if (listRevCheckRequired != null && listRevCheckRequired.Count() > 0) //check if category has value then get all category text
                                    {
                                        foreach (var itemOption in listRevCheckRequired.ToList())
                                        {
                                            if (itemOption.Status == "active")
                                            {
                                                SoxTrackerAppCategory soxTrackerAppCategory = new SoxTrackerAppCategory();
                                                soxTrackerAppCategory.Option = itemOption.Text;
                                                soxTrackerAppCategory.FieldId = soxTrackerQuestionnaire.Q14AFieldId;
                                                soxTrackerAppCategory.SoxTrackerQuestionnaire = soxTrackerQuestionnaire;
                                                listSoxTrackerAppCategory.Add(soxTrackerAppCategory);
                                            }
                                        }
                                    }
                                    #endregion

                                    #region 14. Testing status (WT)
                                    //Question 13
                                    soxTrackerQuestionnaire.Q14BFieldId = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("QuestionnaireFields").GetSection("Q14B").Value);
                                    AppItemField testStatusWT = item.Field<AppItemField>(soxTrackerQuestionnaire.Q14BFieldId);
                                    IEnumerable<Item> relatedTestStatusWT = testStatusWT.Items != null ? testStatusWT.Items : null;
                                    soxTrackerQuestionnaire.Q14BLabel = testStatusWT.Label; //get label

                                    if (relatedTestStatusWT != null && relatedTestStatusWT.Count() > 0)
                                    {
                                        foreach (var itemApp in relatedTestStatusWT.ToList())
                                        {
                                            SoxTrackerAppRelationship soxTrackerAppRelationship = new SoxTrackerAppRelationship();
                                            soxTrackerAppRelationship.Title = itemApp.Title;
                                            soxTrackerAppRelationship.PodioItemId = itemApp.ItemId;
                                            soxTrackerAppRelationship.PodioUniqueId = itemApp.AppItemIdFormatted != null ? itemApp.AppItemIdFormatted.ToString() : string.Empty;
                                            soxTrackerAppRelationship.PodioRevision = itemApp.CurrentRevision != null ? itemApp.CurrentRevision.Revision : 0;
                                            soxTrackerAppRelationship.PodioLink = itemApp.Link.ToString();
                                            soxTrackerAppRelationship.CreatedBy = itemApp.CreatedBy.Name.ToString();
                                            soxTrackerAppRelationship.CreatedOn = DateTime.Parse(itemApp.CreatedOn.ToString());
                                            soxTrackerAppRelationship.SoxTrackerQuestionnaire = soxTrackerQuestionnaire;
                                            soxTrackerAppRelationship.FieldId = soxTrackerQuestionnaire.Q14BFieldId;
                                            listSoxTrackerAppRelationship.Add(soxTrackerAppRelationship);
                                            //soxTrackerQuestionnaire.ListSoxTrackerAppRelationship.Add(soxTrackerAppRelationship);
                                        }
                                    }
                                    #endregion

                                    #region 14. Testing status (R1)
                                    //Question 13
                                    soxTrackerQuestionnaire.Q14CFieldId = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("QuestionnaireFields").GetSection("Q14C").Value);
                                    AppItemField testStatusR1 = item.Field<AppItemField>(soxTrackerQuestionnaire.Q14CFieldId);
                                    IEnumerable<Item> relatedTestStatusR1 = testStatusR1.Items != null ? testStatusR1.Items : null;
                                    soxTrackerQuestionnaire.Q14CLabel = testStatusR1.Label; //get label

                                    if (relatedTestStatusR1 != null && relatedTestStatusR1.Count() > 0)
                                    {
                                        foreach (var itemApp in relatedTestStatusR1.ToList())
                                        {
                                            SoxTrackerAppRelationship soxTrackerAppRelationship = new SoxTrackerAppRelationship();
                                            soxTrackerAppRelationship.Title = itemApp.Title;
                                            soxTrackerAppRelationship.PodioItemId = itemApp.ItemId;
                                            soxTrackerAppRelationship.PodioUniqueId = itemApp.AppItemIdFormatted != null ? itemApp.AppItemIdFormatted.ToString() : string.Empty;
                                            soxTrackerAppRelationship.PodioRevision = itemApp.CurrentRevision != null ? itemApp.CurrentRevision.Revision : 0;
                                            soxTrackerAppRelationship.PodioLink = itemApp.Link.ToString();
                                            soxTrackerAppRelationship.CreatedBy = itemApp.CreatedBy.Name.ToString();
                                            soxTrackerAppRelationship.CreatedOn = DateTime.Parse(itemApp.CreatedOn.ToString());
                                            soxTrackerAppRelationship.SoxTrackerQuestionnaire = soxTrackerQuestionnaire;
                                            soxTrackerAppRelationship.FieldId = soxTrackerQuestionnaire.Q14CFieldId;
                                            listSoxTrackerAppRelationship.Add(soxTrackerAppRelationship);
                                        }
                                    }
                                    #endregion

                                    #region 14. Testing status (R2)
                                    //Question 13
                                    soxTrackerQuestionnaire.Q14DFieldId = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("QuestionnaireFields").GetSection("Q14D").Value);
                                    AppItemField testStatusR2 = item.Field<AppItemField>(soxTrackerQuestionnaire.Q14DFieldId);
                                    IEnumerable<Item> relatedTestStatusR2 = testStatusR2.Items != null ? testStatusR2.Items : null;
                                    soxTrackerQuestionnaire.Q14DLabel = testStatusR2.Label; //get label

                                    if (relatedTestStatusR2 != null && relatedTestStatusR2.Count() > 0)
                                    {
                                        foreach (var itemApp in relatedTestStatusR2.ToList())
                                        {
                                            SoxTrackerAppRelationship soxTrackerAppRelationship = new SoxTrackerAppRelationship();
                                            soxTrackerAppRelationship.Title = itemApp.Title;
                                            soxTrackerAppRelationship.PodioItemId = itemApp.ItemId;
                                            soxTrackerAppRelationship.PodioUniqueId = itemApp.AppItemIdFormatted != null ? itemApp.AppItemIdFormatted.ToString() : string.Empty;
                                            soxTrackerAppRelationship.PodioRevision = itemApp.CurrentRevision != null ? itemApp.CurrentRevision.Revision : 0;
                                            soxTrackerAppRelationship.PodioLink = itemApp.Link.ToString();
                                            soxTrackerAppRelationship.CreatedBy = itemApp.CreatedBy.Name.ToString();
                                            soxTrackerAppRelationship.CreatedOn = DateTime.Parse(itemApp.CreatedOn.ToString());
                                            soxTrackerAppRelationship.SoxTrackerQuestionnaire = soxTrackerQuestionnaire;
                                            soxTrackerAppRelationship.FieldId = soxTrackerQuestionnaire.Q14DFieldId;
                                            listSoxTrackerAppRelationship.Add(soxTrackerAppRelationship);
                                            //soxTrackerQuestionnaire.ListSoxTrackerAppRelationship.Add(soxTrackerAppRelationship);
                                        }
                                    }
                                    #endregion

                                    #region 14. Testing status (R3)
                                    //Question R3
                                    soxTrackerQuestionnaire.Q14EFieldId = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("QuestionnaireFields").GetSection("Q14E").Value);
                                    AppItemField testStatusR3 = item.Field<AppItemField>(soxTrackerQuestionnaire.Q14EFieldId);
                                    IEnumerable<Item> relatedTestStatusR3 = testStatusR3.Items != null ? testStatusR3.Items : null;
                                    soxTrackerQuestionnaire.Q14ELabel = secondLevRevR3.Label; //get label

                                    if (relatedTestStatusR3 != null && relatedTestStatusR3.Count() > 0)
                                    {
                                        foreach (var itemApp in relatedTestStatusR3.ToList())
                                        {
                                            SoxTrackerAppRelationship soxTrackerAppRelationship = new SoxTrackerAppRelationship();
                                            soxTrackerAppRelationship.Title = itemApp.Title;
                                            soxTrackerAppRelationship.PodioItemId = itemApp.ItemId;
                                            soxTrackerAppRelationship.PodioUniqueId = itemApp.AppItemIdFormatted != null ? itemApp.AppItemIdFormatted.ToString() : string.Empty;
                                            soxTrackerAppRelationship.PodioRevision = itemApp.CurrentRevision != null ? itemApp.CurrentRevision.Revision : 0;
                                            soxTrackerAppRelationship.PodioLink = itemApp.Link.ToString();
                                            soxTrackerAppRelationship.CreatedBy = itemApp.CreatedBy.Name.ToString();
                                            soxTrackerAppRelationship.CreatedOn = DateTime.Parse(itemApp.CreatedOn.ToString());
                                            soxTrackerAppRelationship.SoxTrackerQuestionnaire = soxTrackerQuestionnaire;
                                            soxTrackerAppRelationship.FieldId = soxTrackerQuestionnaire.Q14EFieldId;
                                            listSoxTrackerAppRelationship.Add(soxTrackerAppRelationship);
                                        }
                                    }
                                    #endregion




                                    soxTrackerQuestionnaire.ListSoxTrackerAppRelationship = listSoxTrackerAppRelationship;
                                    soxTrackerQuestionnaire.ListSoxTrackerAppCategory = listSoxTrackerAppCategory;
                                    listSoxTrackerQuestionnaire.Add(soxTrackerQuestionnaire);
                                }
                                //pubCollection = collection;
                            }

                            offset += 500;
                        }

                        foreach (var item in listSoxTrackerQuestionnaire)
                        {
                            using (var context = _soxContext.Database.BeginTransaction())
                            {
                                //check if podio item id is not zero
                                if (item.PodioItemId != 0)
                                {
                                    //check if already exists
                                    var checkSoxTrackerQuestionnaire = _soxContext.SoxTrackerQuestionnaire
                                        .Where(x => x.PodioItemId.Equals(item.PodioItemId))
                                        .FirstOrDefault();
                                    if (checkSoxTrackerQuestionnaire != null)
                                    {
                                        item.Id = checkSoxTrackerQuestionnaire.Id;
                                        _soxContext.Entry(checkSoxTrackerQuestionnaire).CurrentValues.SetValues(item);
                                        await _soxContext.SaveChangesAsync();

                                        #region SoxTrackerAppRelationship
                                        //Remove SoxTrackerAppRelationship
                                        var checkSoxTrackerAppRelationship = _soxContext.SoxTrackerAppRelationship.Where(x => x.SoxTrackerQuestionnaire.Id.Equals(checkSoxTrackerQuestionnaire.Id)).ToList();
                                        if (checkSoxTrackerAppRelationship != null)
                                        {
                                            foreach (var itemApp in checkSoxTrackerAppRelationship)
                                            {
                                                _soxContext.Remove(itemApp);
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }

                                        //Add SoxTrackerAppRelationship
                                        if (item.ListSoxTrackerAppRelationship != null && item.ListSoxTrackerAppRelationship.Any())
                                        {
                                            //item.ListQ1Year = item.ListQ1Year.Select(x => { x.RcmQuestionnaire = checkSoxTrackerQuestionnaire; return x; }).ToList();
                                            foreach (var itemApp in item.ListSoxTrackerAppRelationship)
                                            {
                                                _soxContext.Entry(itemApp).State = EntityState.Added;
                                                _soxContext.Entry(checkSoxTrackerQuestionnaire).State = EntityState.Unchanged;
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }
                                        #endregion

                                        #region SoxTrackerAppCategory
                                        //Remove SoxTrackerAppCategory
                                        var checkSoxTrackerAppCategory = _soxContext.SoxTrackerAppCategory.Where(x => x.SoxTrackerQuestionnaire.Equals(checkSoxTrackerQuestionnaire)).ToList();
                                        if (checkSoxTrackerAppCategory != null)
                                        {
                                            foreach (var itemCat in checkSoxTrackerAppCategory)
                                            {
                                                _soxContext.Remove(itemCat);
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }

                                        //Add SoxTrackerAppCategory
                                        if (item.ListSoxTrackerAppCategory != null && item.ListSoxTrackerAppCategory.Any())
                                        {
                                            //item.ListQ7FinStatementElement = item.ListQ7FinStatementElement.Select(x => { x.RcmQuestionnaire = checkSoxTrackerQuestionnaire; return x; }).ToList();
                                            foreach (var itemCat in item.ListSoxTrackerAppCategory)
                                            {

                                                _soxContext.Entry(itemCat).State = EntityState.Added;
                                                _soxContext.Entry(checkSoxTrackerQuestionnaire).State = EntityState.Unchanged;
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }
                                        #endregion

                                        //await _soxContext.SaveChangesAsync();
                                        context.Commit();
                                    }
                                    else
                                    {
                                        _soxContext.Add(item);
                                        await _soxContext.SaveChangesAsync();
                                        context.Commit();
                                    }
                                }
                            }

                        }

                        List<int> ItemId = new List<int>();
                        foreach (var item in listSoxTrackerQuestionnaire)
                        {
                            ItemId.Add(item.PodioItemId);
                        }

                        return Ok(ItemId);
                    }

                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncRcmQuestionnaire");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncRcmQuestionnaire");
                return BadRequest(ex.ToString());
            }

            return NoContent();

        }

        //[AllowAnonymous]
        [HttpPost("soxtracker")]
        public async Task<IActionResult> SyncSoxTrackerAsync([FromBody] SyncDateRange syncDateRange)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("SoxTrackerApp").GetSection("AppId").Value;
                PodioAppKey.AppToken = _config.GetSection("SoxTrackerApp").GetSection("AppToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);
                int totalItems = 0;
                int total = 0;
                int count = 0;
                #region SoxTracker Fields
                int q1Field = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q1").Value);
                int q2Field = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q2").Value);
                int q3Field = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q3").Value);
                int q4Field = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q4").Value);
                int q5Field = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q5").Value);
                int q6Field = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q6").Value);
                int q7Field = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q7").Value);
                int q7OtherField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q7Other").Value);
                int q8Field = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q8").Value);
                int q9Field = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q9").Value);
                int q10Field = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q10").Value);
                int q11Field = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q11").Value);
                int q12AField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q12A").Value);
                int q12BField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q12B").Value);
                int q12CField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q12C").Value);
                int q12DField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q12D").Value);
                int q13AField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q13A").Value);
                int q13BField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q13B").Value);
                int q13CField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q13C").Value);
                int q13DField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q13D").Value);
                int q13EField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q13E").Value);
                int q13FField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q13F").Value);
                int q13GField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q13G").Value);
                int q13HField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q13H").Value);
                int q13IField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q13I").Value);
                int q13JField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q13J").Value);
                int q13KField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q13K").Value);
                int q13LField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q13L").Value);
                int q14AField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q14A").Value);
                int q14BField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q14B").Value);
                int q14CField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q14C").Value);
                int q14DField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q14D").Value);
                int q14EField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q14E").Value);
                int q15Field = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q15").Value);
                int q16Field = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Q16").Value);
                int statusField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Status").Value);
                int durationField = int.Parse(_config.GetSection("SoxTrackerApp").GetSection("FieldId").GetSection("Duration").Value);
                #endregion


                if (podio.IsAuthenticated())
                {

                    var filter = new Dictionary<string, object>
                    {
                        {
                            "last_edit_on",
                            new  { from = syncDateRange.startDate, to = syncDateRange.endDate}
                        },
                    };

                    PodioCollection<Item> collectionCheck = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 1, 0, filters: filter, sortDesc: true);

                    List<SoxTracker> listSoxTracker = new List<SoxTracker>();
                    List<SoxTracker> listSoxTrackerAdd = new List<SoxTracker>();
                    List<SoxTracker> listSoxTrackerUpdate = new List<SoxTracker>();

                    if (syncDateRange.limit == 0 && syncDateRange.offset == 0)
                    {
                        if (collectionCheck.Filtered != 0)
                        {
                            int loop = (int)Math.Ceiling((double)collectionCheck.Filtered / 500);
                            int offset = 0;
                            
                            totalItems = collectionCheck.Filtered;
                            //get filtered items and stored in list
                            for (int i = 0; i < loop; i++)
                            {
                                PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 500, offset, filters: filter, sortDesc: true);

                                foreach (var item in collection.Items)
                                {
                                    SoxTracker soxTracker = new SoxTracker();

                                    count++;
                                    #region Podio Item
                                    soxTracker.PodioItemId = (int)item.ItemId;
                                    soxTracker.PodioUniqueId = item.AppItemIdFormatted.ToString();
                                    soxTracker.PodioRevision = item.CurrentRevision.Revision;
                                    soxTracker.PodioLink = item.Link.ToString();
                                    soxTracker.CreatedBy = item.CreatedBy.Name.ToString();
                                    soxTracker.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());

                                    #region SoxTracker Fields

                                    //1. Year
                                    TextItemField txtFy = item.Field<TextItemField>(q1Field);
                                    soxTracker.FY = txtFy.Value;

                                    //2. What is the Client Name
                                    AppItemField clientApp = item.Field<AppItemField>(q2Field);
                                    IEnumerable<Item> clientAppRef = clientApp.Items;
                                    soxTracker.ClientName = clientAppRef.Select(x => x.Title).FirstOrDefault();
                                    soxTracker.ClientItemId = clientAppRef.Select(x => x.ItemId).FirstOrDefault();

                                    //search for client code in database
                                    var clientCheck = _soxContext.ClientSs.FirstOrDefault(id => id.PodioItemId == soxTracker.ClientItemId);
                                    if (clientCheck != null)
                                    {
                                        soxTracker.ClientCode = clientCheck.ClientCode;
                                    }

                                    //3. What is the Process
                                    AppItemField processApp = item.Field<AppItemField>(q3Field);
                                    IEnumerable<Item> processAppRef = processApp.Items;
                                    soxTracker.Process = processAppRef.Select(x => x.Title).FirstOrDefault();

                                    //4. What is the Sub-Process
                                    AppItemField subprocessApp = item.Field<AppItemField>(q4Field);
                                    IEnumerable<Item> subprocessAppRef = subprocessApp.Items;
                                    soxTracker.Subprocess = subprocessAppRef.Select(x => x.Title).FirstOrDefault();

                                    //5. Control ID?
                                    AppItemField controlIdApp = item.Field<AppItemField>(q5Field);
                                    IEnumerable<Item> controlIdAppRef = controlIdApp.Items;
                                    soxTracker.ControlId = controlIdAppRef.Select(x => x.Title).FirstOrDefault();

                                    //6. What PBC's (supporting documents) are needed to test this control?
                                    TextItemField txtPbc = item.Field<TextItemField>(q6Field);
                                    soxTracker.PBC = txtPbc.Value;

                                    //7. Who is the PBC Owner
                                    AppItemField pbcOwnerApp = item.Field<AppItemField>(q7Field);
                                    IEnumerable<Item> pbcOwnerAppRef = pbcOwnerApp.Items;
                                    soxTracker.PBCOwner = pbcOwnerAppRef.Select(x => x.Title).FirstOrDefault();

                                    //7. Who is the PBC Owner (Other)
                                    TextItemField txtPbcOwner = item.Field<TextItemField>(q7OtherField);
                                    soxTracker.PBCOwnerOther = txtPbcOwner.Value;

                                    //8. Does the control need a population file request?
                                    CategoryItemField catPopFile = item.Field<CategoryItemField>(q8Field);
                                    IEnumerable<CategoryItemField.Option> listcatPopFile = catPopFile.Options;
                                    soxTracker.PopulationFileRequest = listcatPopFile.Select(x => x.Text).FirstOrDefault();

                                    //9. Is sample selection/sub-selection required?
                                    CategoryItemField catSubSelection = item.Field<CategoryItemField>(q9Field);
                                    IEnumerable<CategoryItemField.Option> listcatSubSelection = catSubSelection.Options;
                                    soxTracker.SampleSelection = listcatSubSelection.Select(x => x.Text).FirstOrDefault();

                                    //10. Does the external auditor require samples to be tested in R3 (Q4)?
                                    CategoryItemField catExternalAuditor = item.Field<CategoryItemField>(q10Field);
                                    IEnumerable<CategoryItemField.Option> listcatExternalAuditor = catExternalAuditor.Options;
                                    soxTracker.ExternalAuditorSample = listcatExternalAuditor.Select(x => x.Text).FirstOrDefault();

                                    //11. How many samples to be tested in R3(Q4)?
                                    TextItemField txtSampleTestedR3 = item.Field<TextItemField>(q11Field);
                                    soxTracker.R3Sample = txtSampleTestedR3.Value;


                                    //12. What is the PBC status (WT)
                                    AppItemField WTPBCApp = item.Field<AppItemField>(q12AField);
                                    IEnumerable<Item> WTPBCAppRef = WTPBCApp.Items;
                                    soxTracker.WTPBC = WTPBCAppRef.Select(x => x.Title).FirstOrDefault();

                                    //12. What is the PBC status (R1)
                                    AppItemField R1PBCApp = item.Field<AppItemField>(q12BField);
                                    IEnumerable<Item> R1PBCAppRef = R1PBCApp.Items;
                                    soxTracker.R1PBC = R1PBCAppRef.Select(x => x.Title).FirstOrDefault();

                                    //12. What is the PBC status (R2)
                                    AppItemField R2PBCApp = item.Field<AppItemField>(q12CField);
                                    IEnumerable<Item> R2PBCAppRef = R2PBCApp.Items;
                                    soxTracker.R2PBC = R2PBCAppRef.Select(x => x.Title).FirstOrDefault();

                                    //12. What is the PBC status (R3)
                                    AppItemField R3PBCApp = item.Field<AppItemField>(q12DField);
                                    IEnumerable<Item> R3PBCAppRef = R3PBCApp.Items;
                                    soxTracker.R3PBC = R3PBCAppRef.Select(x => x.Title).FirstOrDefault();



                                    //13. Testing Assignment (WT)
                                    AppItemField WTTesterApp = item.Field<AppItemField>(q13AField);
                                    IEnumerable<Item> WTTesterAppRef = WTTesterApp.Items;
                                    soxTracker.WTTester = WTTesterAppRef.Select(x => x.Title).FirstOrDefault();

                                    //13. Testing Assignment (R1)
                                    AppItemField R1TesterApp = item.Field<AppItemField>(q13BField);
                                    IEnumerable<Item> R1TesterAppRef = R1TesterApp.Items;
                                    soxTracker.R1Tester = R1TesterAppRef.Select(x => x.Title).FirstOrDefault();

                                    //13. Testing Assignment (R2)
                                    AppItemField R2TesterApp = item.Field<AppItemField>(q13CField);
                                    IEnumerable<Item> R2TesterAppRef = R2TesterApp.Items;
                                    soxTracker.R2Tester = R2TesterAppRef.Select(x => x.Title).FirstOrDefault();

                                    //13. Testing Assignment (R3)
                                    AppItemField R3TesterApp = item.Field<AppItemField>(q13DField);
                                    IEnumerable<Item> R3TesterAppRef = R3TesterApp.Items;
                                    soxTracker.R3Tester = R3TesterAppRef.Select(x => x.Title).FirstOrDefault();



                                    //13. 1st Level Reviewer WT
                                    AppItemField WT1LReviewerApp = item.Field<AppItemField>(q13EField);
                                    IEnumerable<Item> WT1LReviewerAppRef = WT1LReviewerApp.Items;
                                    soxTracker.WT1LReviewer = WT1LReviewerAppRef.Select(x => x.Title).FirstOrDefault();

                                    //13. 1st Level Reviewer R1
                                    AppItemField R11LReviewerApp = item.Field<AppItemField>(q13FField);
                                    IEnumerable<Item> R11LReviewerAppRef = R11LReviewerApp.Items;
                                    soxTracker.R11LReviewer = R11LReviewerAppRef.Select(x => x.Title).FirstOrDefault();

                                    //13. 1st Level Reviewer R2
                                    AppItemField R21LReviewerApp = item.Field<AppItemField>(q13GField);
                                    IEnumerable<Item> R21LReviewerAppRef = R21LReviewerApp.Items;
                                    soxTracker.R21LReviewer = R21LReviewerAppRef.Select(x => x.Title).FirstOrDefault();

                                    //13. 1st Level Reviewer R3
                                    AppItemField R31LReviewerApp = item.Field<AppItemField>(q13HField);
                                    IEnumerable<Item> R31LReviewerAppRef = R31LReviewerApp.Items;
                                    soxTracker.R31LReviewer = R31LReviewerAppRef.Select(x => x.Title).FirstOrDefault();



                                    //13. 2nd Level Reviewer WT
                                    AppItemField WT2LReviewerApp = item.Field<AppItemField>(q13IField);
                                    IEnumerable<Item> WT2LReviewerAppRef = WT2LReviewerApp.Items;
                                    soxTracker.WT2LReviewer = WT2LReviewerAppRef.Select(x => x.Title).FirstOrDefault();

                                    //13. 2nd Level Reviewer R1
                                    AppItemField R12LReviewerApp = item.Field<AppItemField>(q13JField);
                                    IEnumerable<Item> R12LReviewerAppRef = R12LReviewerApp.Items;
                                    soxTracker.R12LReviewer = R12LReviewerAppRef.Select(x => x.Title).FirstOrDefault();

                                    //13. 2nd Level Reviewer R2
                                    AppItemField R22LReviewerApp = item.Field<AppItemField>(q13KField);
                                    IEnumerable<Item> R22LReviewerAppRef = R22LReviewerApp.Items;
                                    soxTracker.R22LReviewer = R21LReviewerAppRef.Select(x => x.Title).FirstOrDefault();

                                    //13. 2nd Level Reviewer R3
                                    AppItemField R32LReviewerApp = item.Field<AppItemField>(q13LField);
                                    IEnumerable<Item> R32LReviewerAppRef = R32LReviewerApp.Items;
                                    soxTracker.R32LReviewer = R31LReviewerAppRef.Select(x => x.Title).FirstOrDefault();



                                    //14. Reviewer checklist required?
                                    CategoryItemField catRevChecklist = item.Field<CategoryItemField>(q14AField);
                                    IEnumerable<CategoryItemField.Option> listcatRevChecklist = catRevChecklist.Options;

                                    soxTracker.RCRWT = 0;
                                    soxTracker.RCRR1 = 0;
                                    soxTracker.RCRR2 = 0;
                                    soxTracker.RCRR3 = 0;

                                    if (listcatRevChecklist != null && listcatRevChecklist.Count() > 0)
                                    {
                                        foreach (var itemCheckList in listcatRevChecklist)
                                        {

                                            switch (itemCheckList.Text)
                                            {
                                                case "Walkthrough":
                                                    soxTracker.RCRWT = 1;
                                                    break;
                                                case "Round 1":
                                                    soxTracker.RCRR1 = 1;
                                                    break;
                                                case "Round 2":
                                                    soxTracker.RCRR2 = 1;
                                                    break;
                                                case "Round 3":
                                                    soxTracker.RCRR3 = 1;
                                                    break;
                                                default:
                                                    break;
                                            }

                                        }
                                    }

                                    //14. Testing status (WT)
                                    AppItemField WTTestingStatusApp = item.Field<AppItemField>(q14BField);
                                    IEnumerable<Item> WTTestingStatusAppRef = WTTestingStatusApp.Items;
                                    soxTracker.WTTestingStatus = WTTestingStatusAppRef.Select(x => x.Title).FirstOrDefault();

                                    //14. Testing status (R1)
                                    AppItemField R1TestingStatusApp = item.Field<AppItemField>(q14CField);
                                    IEnumerable<Item> R1TestingStatusAppRef = R1TestingStatusApp.Items;
                                    soxTracker.R1TestingStatus = R1TestingStatusAppRef.Select(x => x.Title).FirstOrDefault();

                                    //14. Testing status (R2)
                                    AppItemField R2TestingStatusApp = item.Field<AppItemField>(q14DField);
                                    IEnumerable<Item> R2TestingStatusAppRef = R2TestingStatusApp.Items;
                                    soxTracker.R2TestingStatus = R2TestingStatusAppRef.Select(x => x.Title).FirstOrDefault();

                                    //14. Testing status (R3)
                                    AppItemField R3TestingStatusApp = item.Field<AppItemField>(q14EField);
                                    IEnumerable<Item> R3TestingStatusAppRef = R3TestingStatusApp.Items;
                                    soxTracker.R3TestingStatus = R3TestingStatusAppRef.Select(x => x.Title).FirstOrDefault();


                                    //15. Is this a key report?
                                    TextItemField txtKeyReport = item.Field<TextItemField>(q15Field);
                                    soxTracker.KeyReport = txtKeyReport.Value;

                                    //16. Key report name
                                    TextItemField txtKeyReportName = item.Field<TextItemField>(q16Field);
                                    soxTracker.KeyReportName = txtKeyReportName.Value;


                                    //Status
                                    CategoryItemField catStatus = item.Field<CategoryItemField>(statusField);
                                    IEnumerable<CategoryItemField.Option> listcatRisk = catStatus.Options;
                                    soxTracker.Status = listcatRisk.Select(x => x.Text).FirstOrDefault();

                                    //Duration
                                    DurationItemField durationItem = item.Field<DurationItemField>("duration");
                                    soxTracker.Duration = durationItem.Value;


                                    #endregion

                                    #endregion

                                    listSoxTracker.Add(soxTracker);
                                }

                                offset += 500;
                            }

                            total = count >= collectionCheck.Filtered ? collectionCheck.Filtered : count * (syncDateRange.offset + 1);
                            totalItems = collectionCheck.Filtered;

                        }
                    }
                    else
                    {
                        if (collectionCheck.Filtered != 0)
                        {
                            
                            int offset = syncDateRange.offset;
                            PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), syncDateRange.limit, offset, filters: filter, sortDesc: true);

                            foreach (var item in collection.Items)
                            {
                                SoxTracker soxTracker = new SoxTracker();
                                count++;
                                #region Podio Item
                                soxTracker.PodioItemId = (int)item.ItemId;
                                soxTracker.PodioUniqueId = item.AppItemIdFormatted.ToString();
                                soxTracker.PodioRevision = item.CurrentRevision.Revision;
                                soxTracker.PodioLink = item.Link.ToString();
                                soxTracker.CreatedBy = item.CreatedBy.Name.ToString();
                                soxTracker.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());


                                //1. Year
                                TextItemField txtFy = item.Field<TextItemField>(q1Field);
                                soxTracker.FY = txtFy.Value;

                                //2. What is the Client Name
                                AppItemField clientApp = item.Field<AppItemField>(q2Field);
                                IEnumerable<Item> clientAppRef = clientApp.Items;
                                soxTracker.ClientName = clientAppRef.Select(x => x.Title).FirstOrDefault();
                                soxTracker.ClientItemId = clientAppRef.Select(x => x.ItemId).FirstOrDefault();

                                //search for client code in database
                                var clientCheck = _soxContext.ClientSs.FirstOrDefault(id => id.PodioItemId == soxTracker.ClientItemId);
                                if (clientCheck != null)
                                {
                                    soxTracker.ClientCode = clientCheck.ClientCode;
                                }

                                //3. What is the Process
                                AppItemField processApp = item.Field<AppItemField>(q3Field);
                                IEnumerable<Item> processAppRef = processApp.Items;
                                soxTracker.Process = processAppRef.Select(x => x.Title).FirstOrDefault();

                                //4. What is the Sub-Process
                                AppItemField subprocessApp = item.Field<AppItemField>(q4Field);
                                IEnumerable<Item> subprocessAppRef = subprocessApp.Items;
                                soxTracker.Subprocess = subprocessAppRef.Select(x => x.Title).FirstOrDefault();

                                //5. Control ID?
                                AppItemField controlIdApp = item.Field<AppItemField>(q5Field);
                                IEnumerable<Item> controlIdAppRef = controlIdApp.Items;
                                soxTracker.ControlId = controlIdAppRef.Select(x => x.Title).FirstOrDefault();

                                //6. What PBC's (supporting documents) are needed to test this control?
                                TextItemField txtPbc = item.Field<TextItemField>(q6Field);
                                soxTracker.PBC = txtPbc.Value;

                                //7. Who is the PBC Owner
                                AppItemField pbcOwnerApp = item.Field<AppItemField>(q7Field);
                                IEnumerable<Item> pbcOwnerAppRef = pbcOwnerApp.Items;
                                soxTracker.PBCOwner = pbcOwnerAppRef.Select(x => x.Title).FirstOrDefault();

                                //7. Who is the PBC Owner (Other)
                                TextItemField txtPbcOwner = item.Field<TextItemField>(q7OtherField);
                                soxTracker.PBCOwnerOther = txtPbcOwner.Value;

                                //8. Does the control need a population file request?
                                CategoryItemField catPopFile = item.Field<CategoryItemField>(q8Field);
                                IEnumerable<CategoryItemField.Option> listcatPopFile = catPopFile.Options;
                                soxTracker.PopulationFileRequest = listcatPopFile.Select(x => x.Text).FirstOrDefault();

                                //9. Is sample selection/sub-selection required?
                                CategoryItemField catSubSelection = item.Field<CategoryItemField>(q9Field);
                                IEnumerable<CategoryItemField.Option> listcatSubSelection = catSubSelection.Options;
                                soxTracker.SampleSelection = listcatSubSelection.Select(x => x.Text).FirstOrDefault();

                                //10. Does the external auditor require samples to be tested in R3 (Q4)?
                                CategoryItemField catExternalAuditor = item.Field<CategoryItemField>(q10Field);
                                IEnumerable<CategoryItemField.Option> listcatExternalAuditor = catExternalAuditor.Options;
                                soxTracker.ExternalAuditorSample = listcatExternalAuditor.Select(x => x.Text).FirstOrDefault();

                                //11. How many samples to be tested in R3(Q4)?
                                TextItemField txtSampleTestedR3 = item.Field<TextItemField>(q11Field);
                                soxTracker.R3Sample = txtSampleTestedR3.Value;


                                //12. What is the PBC status (WT)
                                AppItemField WTPBCApp = item.Field<AppItemField>(q12AField);
                                IEnumerable<Item> WTPBCAppRef = WTPBCApp.Items;
                                soxTracker.WTPBC = WTPBCAppRef.Select(x => x.Title).FirstOrDefault();

                                //12. What is the PBC status (R1)
                                AppItemField R1PBCApp = item.Field<AppItemField>(q12BField);
                                IEnumerable<Item> R1PBCAppRef = R1PBCApp.Items;
                                soxTracker.R1PBC = R1PBCAppRef.Select(x => x.Title).FirstOrDefault();

                                //12. What is the PBC status (R2)
                                AppItemField R2PBCApp = item.Field<AppItemField>(q12CField);
                                IEnumerable<Item> R2PBCAppRef = R2PBCApp.Items;
                                soxTracker.R2PBC = R2PBCAppRef.Select(x => x.Title).FirstOrDefault();

                                //12. What is the PBC status (R3)
                                AppItemField R3PBCApp = item.Field<AppItemField>(q12DField);
                                IEnumerable<Item> R3PBCAppRef = R3PBCApp.Items;
                                soxTracker.R3PBC = R3PBCAppRef.Select(x => x.Title).FirstOrDefault();



                                //13. Testing Assignment (WT)
                                AppItemField WTTesterApp = item.Field<AppItemField>(q13AField);
                                IEnumerable<Item> WTTesterAppRef = WTTesterApp.Items;
                                soxTracker.WTTester = WTTesterAppRef.Select(x => x.Title).FirstOrDefault();

                                //13. Testing Assignment (R1)
                                AppItemField R1TesterApp = item.Field<AppItemField>(q13BField);
                                IEnumerable<Item> R1TesterAppRef = R1TesterApp.Items;
                                soxTracker.R1Tester = R1TesterAppRef.Select(x => x.Title).FirstOrDefault();

                                //13. Testing Assignment (R2)
                                AppItemField R2TesterApp = item.Field<AppItemField>(q13CField);
                                IEnumerable<Item> R2TesterAppRef = R2TesterApp.Items;
                                soxTracker.R2Tester = R2TesterAppRef.Select(x => x.Title).FirstOrDefault();

                                //13. Testing Assignment (R3)
                                AppItemField R3TesterApp = item.Field<AppItemField>(q13DField);
                                IEnumerable<Item> R3TesterAppRef = R3TesterApp.Items;
                                soxTracker.R3Tester = R3TesterAppRef.Select(x => x.Title).FirstOrDefault();



                                //13. 1st Level Reviewer WT
                                AppItemField WT1LReviewerApp = item.Field<AppItemField>(q13EField);
                                IEnumerable<Item> WT1LReviewerAppRef = WT1LReviewerApp.Items;
                                soxTracker.WT1LReviewer = WT1LReviewerAppRef.Select(x => x.Title).FirstOrDefault();

                                //13. 1st Level Reviewer R1
                                AppItemField R11LReviewerApp = item.Field<AppItemField>(q13FField);
                                IEnumerable<Item> R11LReviewerAppRef = R11LReviewerApp.Items;
                                soxTracker.R11LReviewer = R11LReviewerAppRef.Select(x => x.Title).FirstOrDefault();

                                //13. 1st Level Reviewer R2
                                AppItemField R21LReviewerApp = item.Field<AppItemField>(q13GField);
                                IEnumerable<Item> R21LReviewerAppRef = R21LReviewerApp.Items;
                                soxTracker.R21LReviewer = R21LReviewerAppRef.Select(x => x.Title).FirstOrDefault();

                                //13. 1st Level Reviewer R3
                                AppItemField R31LReviewerApp = item.Field<AppItemField>(q13HField);
                                IEnumerable<Item> R31LReviewerAppRef = R31LReviewerApp.Items;
                                soxTracker.R31LReviewer = R31LReviewerAppRef.Select(x => x.Title).FirstOrDefault();



                                //13. 2nd Level Reviewer WT
                                AppItemField WT2LReviewerApp = item.Field<AppItemField>(q13IField);
                                IEnumerable<Item> WT2LReviewerAppRef = WT2LReviewerApp.Items;
                                soxTracker.WT2LReviewer = WT2LReviewerAppRef.Select(x => x.Title).FirstOrDefault();

                                //13. 2nd Level Reviewer R1
                                AppItemField R12LReviewerApp = item.Field<AppItemField>(q13JField);
                                IEnumerable<Item> R12LReviewerAppRef = R12LReviewerApp.Items;
                                soxTracker.R12LReviewer = R12LReviewerAppRef.Select(x => x.Title).FirstOrDefault();

                                //13. 2nd Level Reviewer R2
                                AppItemField R22LReviewerApp = item.Field<AppItemField>(q13KField);
                                IEnumerable<Item> R22LReviewerAppRef = R22LReviewerApp.Items;
                                soxTracker.R22LReviewer = R21LReviewerAppRef.Select(x => x.Title).FirstOrDefault();

                                //13. 2nd Level Reviewer R3
                                AppItemField R32LReviewerApp = item.Field<AppItemField>(q13LField);
                                IEnumerable<Item> R32LReviewerAppRef = R32LReviewerApp.Items;
                                soxTracker.R32LReviewer = R31LReviewerAppRef.Select(x => x.Title).FirstOrDefault();



                                //14. Reviewer checklist required?
                                CategoryItemField catRevChecklist = item.Field<CategoryItemField>(q14AField);
                                IEnumerable<CategoryItemField.Option> listcatRevChecklist = catRevChecklist.Options;

                                soxTracker.RCRWT = 0;
                                soxTracker.RCRR1 = 0;
                                soxTracker.RCRR2 = 0;
                                soxTracker.RCRR3 = 0;

                                if (listcatRevChecklist != null && listcatRevChecklist.Count() > 0)
                                {
                                    foreach (var itemCheckList in listcatRevChecklist)
                                    {

                                        switch (itemCheckList.Text)
                                        {
                                            case "Walkthrough":
                                                soxTracker.RCRWT = 1;
                                                break;
                                            case "Round 1":
                                                soxTracker.RCRR1 = 1;
                                                break;
                                            case "Round 2":
                                                soxTracker.RCRR2 = 1;
                                                break;
                                            case "Round 3":
                                                soxTracker.RCRR3 = 1;
                                                break;
                                            default:
                                                break;
                                        }

                                    }
                                }

                                //14. Testing status (WT)
                                AppItemField WTTestingStatusApp = item.Field<AppItemField>(q14BField);
                                IEnumerable<Item> WTTestingStatusAppRef = WTTestingStatusApp.Items;
                                soxTracker.WTTestingStatus = WTTestingStatusAppRef.Select(x => x.Title).FirstOrDefault();

                                //14. Testing status (R1)
                                AppItemField R1TestingStatusApp = item.Field<AppItemField>(q14CField);
                                IEnumerable<Item> R1TestingStatusAppRef = R1TestingStatusApp.Items;
                                soxTracker.R1TestingStatus = R1TestingStatusAppRef.Select(x => x.Title).FirstOrDefault();

                                //14. Testing status (R2)
                                AppItemField R2TestingStatusApp = item.Field<AppItemField>(q14DField);
                                IEnumerable<Item> R2TestingStatusAppRef = R2TestingStatusApp.Items;
                                soxTracker.R2TestingStatus = R2TestingStatusAppRef.Select(x => x.Title).FirstOrDefault();

                                //14. Testing status (R3)
                                AppItemField R3TestingStatusApp = item.Field<AppItemField>(q14EField);
                                IEnumerable<Item> R3TestingStatusAppRef = R3TestingStatusApp.Items;
                                soxTracker.R3TestingStatus = R3TestingStatusAppRef.Select(x => x.Title).FirstOrDefault();


                                //15. Key report 
                                TextItemField txtKeyReport = item.Field<TextItemField>(q15Field);
                                soxTracker.KeyReport = txtKeyReport.Value;

                                //16. Key report name
                                TextItemField txtKeyReportName = item.Field<TextItemField>(q16Field);
                                soxTracker.KeyReportName = txtKeyReportName.Value;

                                //Status
                                CategoryItemField catStatus = item.Field<CategoryItemField>(statusField);
                                IEnumerable<CategoryItemField.Option> listcatRisk = catStatus.Options;
                                soxTracker.Status = listcatRisk.Select(x => x.Text).FirstOrDefault();

                                #endregion

                                listSoxTracker.Add(soxTracker);

                            }
                            count = count * (syncDateRange.offset + 1);
                            total = count >= collection.Filtered ? collection.Filtered : count;
                            totalItems = collection.Filtered;

                        }
                            

                    }

                    if (listSoxTracker != null && listSoxTracker.Count() > 0)
                    {
                        using (var context = _soxContext.Database.BeginTransaction())
                        {

                            foreach (var item in listSoxTracker)
                            {
                                var soxTrackerCheck = _soxContext.SoxTracker.FirstOrDefault(id => id.PodioItemId == (int)item.PodioItemId);
                                if (soxTrackerCheck != null)
                                {
                                    soxTrackerCheck.FY = item.FY;
                                    soxTrackerCheck.ClientName = item.ClientName;
                                    soxTrackerCheck.ClientCode = item.ClientCode;
                                    soxTrackerCheck.ClientItemId = item.ClientItemId;
                                    soxTrackerCheck.Process = item.Process;
                                    soxTrackerCheck.Subprocess = item.Subprocess;
                                    soxTrackerCheck.ControlId = item.ControlId;
                                    soxTrackerCheck.PBC = item.PBC;
                                    soxTrackerCheck.PBCOwner = item.PBCOwner;
                                    soxTrackerCheck.PBCOwnerOther = item.PBCOwnerOther;
                                    soxTrackerCheck.PopulationFileRequest = item.PopulationFileRequest;
                                    soxTrackerCheck.SampleSelection = item.SampleSelection;
                                    soxTrackerCheck.R3Sample = item.R3Sample;
                                    soxTrackerCheck.WTPBC = item.WTPBC;
                                    soxTrackerCheck.R1PBC = item.R1PBC;
                                    soxTrackerCheck.R2PBC = item.R2PBC;
                                    soxTrackerCheck.R3PBC = item.R3PBC;
                                    soxTrackerCheck.WTTester = item.WTTester;
                                    soxTrackerCheck.WT1LReviewer = item.WT1LReviewer;
                                    soxTrackerCheck.WT2LReviewer = item.WT2LReviewer;
                                    soxTrackerCheck.R1Tester = item.R1Tester;
                                    soxTrackerCheck.R11LReviewer = item.R11LReviewer;
                                    soxTrackerCheck.R12LReviewer = item.R12LReviewer;
                                    soxTrackerCheck.R2Tester = item.R2Tester;
                                    soxTrackerCheck.R21LReviewer = item.R21LReviewer;
                                    soxTrackerCheck.R22LReviewer = item.R22LReviewer;
                                    soxTrackerCheck.R3Tester = item.R3Tester;
                                    soxTrackerCheck.R31LReviewer = item.R31LReviewer;
                                    soxTrackerCheck.R32LReviewer = item.R32LReviewer;
                                    soxTrackerCheck.WTTestingStatus = item.WTTestingStatus;
                                    soxTrackerCheck.R1TestingStatus = item.R1TestingStatus;
                                    soxTrackerCheck.R2TestingStatus = item.R2TestingStatus;
                                    soxTrackerCheck.R3TestingStatus = item.R3TestingStatus;
                                    soxTrackerCheck.RCRWT = item.RCRWT;
                                    soxTrackerCheck.RCRR1 = item.RCRR1;
                                    soxTrackerCheck.RCRR2 = item.RCRR2;
                                    soxTrackerCheck.RCRR3 = item.RCRR3;
                                    soxTrackerCheck.ExternalAuditorSample = item.ExternalAuditorSample;
                                    soxTrackerCheck.Status = item.Status;
                                    soxTrackerCheck.PodioUniqueId = item.PodioUniqueId;
                                    soxTrackerCheck.PodioRevision = item.PodioRevision;
                                    soxTrackerCheck.PodioLink = item.PodioLink;
                                    soxTrackerCheck.CreatedBy = item.CreatedBy;
                                    soxTrackerCheck.KeyReport = item.KeyReport;
                                    soxTrackerCheck.KeyReportName = item.KeyReportName;
                                    listSoxTrackerUpdate.Add(soxTrackerCheck);
                                }
                                else
                                {
                                    listSoxTrackerAdd.Add(item);
                                }
                                Debug.WriteLine($"SoxTracker ItemId: {item.PodioItemId}");
                            }


                            if (listSoxTrackerAdd != null && listSoxTrackerAdd.Count() > 0)
                                _soxContext.AddRange(listSoxTrackerAdd);

                            if (listSoxTrackerUpdate != null && listSoxTrackerUpdate.Count() > 0)
                                _soxContext.UpdateRange(listSoxTrackerUpdate);

                            await _soxContext.SaveChangesAsync();
                            context.Commit();
                        }
                    }

                }
                
                string saveItems = $"{total}/{totalItems}";
                return Ok(new { syncDateRange.limit, totalItems, saveItems });

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncSoxTrackerAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncSoxTrackerAsync");
                return BadRequest(ex.ToString());
            }

        }


        #region Key Report

        //[AllowAnonymous]
        [HttpPost("keyreport/questionnaire/leadsheet/fields")]
        public async Task<IActionResult> SyncKeyReportQuestionnaireLeadsheetnFields()
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;


                PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("LeadsheetAppId").Value;
                PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("LeadsheetAppToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(int.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                if (podio.IsAuthenticated())
                {
                    Debug.WriteLine("Podio Authenticated");
                    int count = 1;
                    int position = 1;
                    var podioResult = await podio.ApplicationService.GetApp(int.Parse(PodioAppKey.AppId));

                    if (podioResult.Fields != null && podioResult.Fields.Count > 0)
                    {

                        if (podioResult.Fields.Count > 0)
                        {

                            using (var context = _soxContext.Database.BeginTransaction())
                            {
                                //Remove all options associated with app id
                                var checkQuestionOptionToDelete = _soxContext.KeyReportOption.Where(id => id.AppId.Equals(int.Parse(PodioAppKey.AppId)));
                                if (checkQuestionOptionToDelete != null && checkQuestionOptionToDelete.Count() > 0)
                                {
                                    _soxContext.RemoveRange(checkQuestionOptionToDelete);
                                    await _soxContext.SaveChangesAsync();
                                    
                                }

                                //Remove all fields associated with app id
                                var checkQuestionToDelete = _soxContext.KeyReportQuestion.Where(x => x.AppId.Equals(PodioAppKey.AppId));
                                if (checkQuestionToDelete != null && checkQuestionToDelete.Count() > 0)
                                {
                                    _soxContext.RemoveRange(checkQuestionToDelete);
                                    await _soxContext.SaveChangesAsync();
                                    
                                }
                                context.Commit();
                            }

                            foreach (var field in podioResult.Fields)
                            {
                                if (field.Status == "active")
                                {
                                    using (var context = _soxContext.Database.BeginTransaction())
                                    {
                                        //writeLog.Display(field);
                                        KeyReportQuestion keyReportField = new KeyReportQuestion();
                                        keyReportField.QuestionString = field.Label;
                                        keyReportField.Type = field.Type;
                                        keyReportField.AppId = PodioAppKey.AppId;
                                        keyReportField.FieldId = field.FieldId;
                                        keyReportField.CreatedOn = DateTime.Now;
                                        keyReportField.UpdatedOn = DateTime.Now;
                                        keyReportField.Position = position;
                                        List<KeyReportOption> listOptions = new List<KeyReportOption>();

                                        if (field.InternalConfig.Description != string.Empty)
                                        {
                                            keyReportField.Description = field.InternalConfig.Description;
                                        }

                                        if (field.Type == "category")
                                        {
                                            foreach (var option in field.InternalConfig.Settings)
                                            {
                                                if (option.Value.HasValues)
                                                {

                                                    foreach (var item in option.Value)
                                                    {
                                                        if (item["status"].ToString() == "active")
                                                        {
                                                            KeyReportOption keyReportOption = new KeyReportOption();
                                                            //Console.WriteLine($"{item["text"]}");
                                                            keyReportOption.OptionName = item["text"].ToString();
                                                            keyReportOption.CreatedOn = DateTime.Now;
                                                            keyReportOption.UpdatedOn = DateTime.Now;
                                                            keyReportOption.AppId = int.Parse(keyReportField.AppId);
                                                            keyReportOption.KeyReportQuestion = keyReportField;
                                                            keyReportOption.OptionId = $"{keyReportField.AppId}{count}";
                                                            listOptions.Add(keyReportOption);
                                                            count++;
                                                        }

                                                    }

                                                    keyReportField.Options = listOptions;
                                                }

                                            }
                                        }

                                        if (field.Type == "date")
                                        {
                                            Debug.WriteLine($"{field.InternalConfig.Settings["end"]}");
                                            if (field.InternalConfig.Settings["end"].ToString() == "enabled")
                                                keyReportField.Tag = "enabled";
                                            else
                                                keyReportField.Tag = string.Empty;

                                        }

                                        if (field.Type == "text")
                                        {
                                            if (field.InternalConfig.Settings["size"].ToString() == "large")
                                            {
                                                keyReportField.Tag = "large";
                                            }
                                        }

                                        if (field.Type == "image")
                                        {
                                            //image field 
                                        }

                                        if (field.Type == "app")
                                        {
                                            //image field 
                                        }


                                        _soxContext.Add(keyReportField);
                                        await _soxContext.SaveChangesAsync();
                                        context.Commit();
                                        position++;

                                    }
                                }
                            }

                        }

                        return Ok();
                    }

                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncKeyReportQuestionnaireLeadsheetnFields");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncKeyReportQuestionnaireLeadsheetnFields");
                return BadRequest(ex.ToString());
            }

            return NoContent();

        }

        //[AllowAnonymous]
        [HttpPost("keyreport/questionnaire/consoleformat/fields")]
        public async Task<IActionResult> SyncKeyReportQuestionnaireConsoleOrigFields()
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;


                PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("ConsolOrigFormatId").Value;
                PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("ConsolOrigFormatToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(int.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                if (podio.IsAuthenticated())
                {
                    Debug.WriteLine("Podio Authenticated");
                    int count = 1;
                    int position = 1;
                    var podioResult = await podio.ApplicationService.GetApp(int.Parse(PodioAppKey.AppId));
          
                    if (podioResult.Fields != null && podioResult.Fields.Count > 0)
                    {

                        if (podioResult.Fields.Count > 0)
                        {

                            using (var context = _soxContext.Database.BeginTransaction())
                            {
                                //Remove all options associated with app id
                                var checkQuestionOptionToDelete = _soxContext.KeyReportOption.Where(x => x.AppId == int.Parse(PodioAppKey.AppId));
                                if (checkQuestionOptionToDelete != null && checkQuestionOptionToDelete.Count() > 0)
                                {
                                    _soxContext.RemoveRange(checkQuestionOptionToDelete);
                                    await _soxContext.SaveChangesAsync();
                                }

                                //Remove all fields associated with app id
                                var checkQuestionToDelete = _soxContext.KeyReportQuestion.Where(x =>x.AppId == PodioAppKey.AppId);
                                if (checkQuestionToDelete != null && checkQuestionToDelete.Count() > 0)
                                {
                                    _soxContext.RemoveRange(checkQuestionToDelete);
                                    await _soxContext.SaveChangesAsync();  
                                }
                                context.Commit();
                            }

                            foreach (var field in podioResult.Fields)
                            {
                                if (field.Status == "active")
                                {
                                    using (var context = _soxContext.Database.BeginTransaction())
                                    {
                                        //writeLog.Display(field);
                                        KeyReportQuestion keyReportField = new KeyReportQuestion();
                                        keyReportField.QuestionString = field.Label;
                                        keyReportField.Type = field.Type;
                                        keyReportField.AppId = PodioAppKey.AppId;
                                        keyReportField.FieldId = field.FieldId;
                                        keyReportField.CreatedOn = DateTime.Now;
                                        keyReportField.UpdatedOn = DateTime.Now;
                                        keyReportField.Position = position;
                                        List<KeyReportOption> listOptions = new List<KeyReportOption>();

                                        if (field.InternalConfig.Description != string.Empty)
                                        {
                                            keyReportField.Description = field.InternalConfig.Description;
                                        }

                                        if (field.Type == "category")
                                        {
                                            foreach (var option in field.InternalConfig.Settings)
                                            {
                                                if (option.Value.HasValues)
                                                {

                                                    foreach (var item in option.Value)
                                                    {
                                                        if (item["status"].ToString() == "active")
                                                        {
                                                            KeyReportOption keyReportOption = new KeyReportOption();
                                                            //Console.WriteLine($"{item["text"]}");
                                                            keyReportOption.OptionName = item["text"].ToString();
                                                            keyReportOption.CreatedOn = DateTime.Now;
                                                            keyReportOption.UpdatedOn = DateTime.Now;
                                                            keyReportOption.AppId = int.Parse(keyReportField.AppId);
                                                            keyReportOption.KeyReportQuestion = keyReportField;
                                                            keyReportOption.OptionId = $"{keyReportField.AppId}{count}";
                                                            listOptions.Add(keyReportOption);
                                                            count++;
                                                        }

                                                    }

                                                    keyReportField.Options = listOptions;
                                                }

                                            }
                                        }

                                        if (field.Type == "date")
                                        {
                                            Debug.WriteLine($"{field.InternalConfig.Settings["end"]}");
                                            if (field.InternalConfig.Settings["end"].ToString() == "enabled")
                                                keyReportField.Tag = "enabled";
                                            else
                                                keyReportField.Tag = string.Empty;

                                        }

                                        if (field.Type == "text")
                                        {
                                            if (field.InternalConfig.Settings["size"].ToString() == "large")
                                            {
                                                keyReportField.Tag = "large";
                                            }
                                        }

                                        if (field.Type == "image")
                                        {
                                            //image field 
                                        }

                                        if (field.Type == "app")
                                        {
                                            //image field 
                                        }


                                        _soxContext.Add(keyReportField);
                                        await _soxContext.SaveChangesAsync();
                                        context.Commit();
                                        position++;

                                    }
                                }
                            }

                        }

                        return Ok();
                    }

                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncKeyReportQuestionnaireConsoleOrigFields");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncKeyReportQuestionnaireConsoleOrigFields");
                return BadRequest(ex.ToString());
            }

            return NoContent();

        }

        //[AllowAnonymous]
        [HttpPost("keyreport/questionnaire/alliuc/fields")]
        public async Task<IActionResult> SyncKeyReportQuestionnaireAllIUCFields()
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;


                PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("AllIUCId").Value;
                PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("AllIUCToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(int.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                if (podio.IsAuthenticated())
                {
                    Debug.WriteLine("Podio Authenticated");
                    int count = 1;
                    int position = 1;
                    var podioResult = await podio.ApplicationService.GetApp(int.Parse(PodioAppKey.AppId));

                    if (podioResult.Fields != null && podioResult.Fields.Count > 0)
                    {

                        if (podioResult.Fields.Count > 0)
                        {

                            using (var context = _soxContext.Database.BeginTransaction())
                            {
                                //Remove all options associated with app id
                                var checkQuestionOptionToDelete = _soxContext.KeyReportOption.Where(id => id.AppId.Equals(int.Parse(PodioAppKey.AppId)));
                                if (checkQuestionOptionToDelete != null && checkQuestionOptionToDelete.Count() > 0)
                                {
                                    _soxContext.RemoveRange(checkQuestionOptionToDelete);
                                    await _soxContext.SaveChangesAsync();
                                }

                                //Remove all fields associated with app id
                                var checkQuestionToDelete = _soxContext.KeyReportQuestion
                                    .Where(
                                        x => x.AppId.Equals(PodioAppKey.AppId));
                                if (checkQuestionToDelete != null && checkQuestionToDelete.Count() > 0)
                                {
                                    _soxContext.RemoveRange(checkQuestionToDelete);
                                    await _soxContext.SaveChangesAsync();
                                }
                                context.Commit();
                            }


                            foreach (var field in podioResult.Fields)
                            {
                                if (field.Status == "active")
                                {
                                    using (var context = _soxContext.Database.BeginTransaction())
                                    {
                                        //writeLog.Display(field);
                                        KeyReportQuestion keyReportField = new KeyReportQuestion();
                                        keyReportField.QuestionString = field.Label;
                                        keyReportField.Type = field.Type;
                                        keyReportField.AppId = PodioAppKey.AppId;
                                        keyReportField.FieldId = field.FieldId;
                                        keyReportField.CreatedOn = DateTime.Now;
                                        keyReportField.UpdatedOn = DateTime.Now;
                                        keyReportField.Position = position;
                                        List<KeyReportOption> listOptions = new List<KeyReportOption>();

                                        if (field.InternalConfig.Description != string.Empty)
                                        {
                                            keyReportField.Description = field.InternalConfig.Description;
                                        }

                                        if (field.Type == "category")
                                        {
                                            foreach (var option in field.InternalConfig.Settings)
                                            {
                                                if (option.Value.HasValues)
                                                {

                                                    foreach (var item in option.Value)
                                                    {
                                                        if (item["status"].ToString() == "active")
                                                        {
                                                            KeyReportOption keyReportOption = new KeyReportOption();
                                                            //Console.WriteLine($"{item["text"]}");
                                                            keyReportOption.OptionName = item["text"].ToString();
                                                            keyReportOption.CreatedOn = DateTime.Now;
                                                            keyReportOption.UpdatedOn = DateTime.Now;
                                                            keyReportOption.AppId = int.Parse(keyReportField.AppId);
                                                            keyReportOption.KeyReportQuestion = keyReportField;
                                                            keyReportOption.OptionId = $"{keyReportField.AppId}{count}";
                                                            listOptions.Add(keyReportOption);
                                                            count++;
                                                        }

                                                    }

                                                    keyReportField.Options = listOptions;
                                                }

                                            }
                                        }

                                        if (field.Type == "date")
                                        {
                                            Debug.WriteLine($"{field.InternalConfig.Settings["end"]}");
                                            if (field.InternalConfig.Settings["end"].ToString() == "enabled")
                                                keyReportField.Tag = "enabled";
                                            else
                                                keyReportField.Tag = string.Empty;

                                        }

                                        if (field.Type == "text")
                                        {
                                            if (field.InternalConfig.Settings["size"].ToString() == "large")
                                            {
                                                keyReportField.Tag = "large";
                                            }
                                        }

                                        if (field.Type == "image")
                                        {
                                            //image field 
                                        }

                                        if (field.Type == "app")
                                        {
                                            //image field 
                                        }


                                        _soxContext.Add(keyReportField);
                                        await _soxContext.SaveChangesAsync();
                                        context.Commit();
                                        position++;

                                    }
                                }
                            }

                        }

                        return Ok();
                    }

                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncKeyReportQuestionnaireAllIUCFields");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncKeyReportQuestionnaireAllIUCFields");
                return BadRequest(ex.ToString());
            }

            return NoContent();

        }

        //[AllowAnonymous]
        [HttpPost("keyreport/questionnaire/statustracker/fields")]
        public async Task<IActionResult> SyncKeyReportQuestionnaireStatusTrackerFields()
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;


                PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("TestStatusTrackerId").Value;
                PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("TestStatusTrackerToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(int.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                if (podio.IsAuthenticated())
                {
                    Debug.WriteLine("Podio Authenticated");
                    int count = 1;
                    int position = 1;
                    var podioResult = await podio.ApplicationService.GetApp(int.Parse(PodioAppKey.AppId));

                    if (podioResult.Fields != null && podioResult.Fields.Count > 0)
                    {

                        if (podioResult.Fields.Count > 0)
                        {

                            using (var context = _soxContext.Database.BeginTransaction())
                            {
                                //Remove all options associated with app id
                                var checkQuestionOptionToDelete = _soxContext.KeyReportOption.Where(id => id.AppId.Equals(int.Parse(PodioAppKey.AppId)));
                                if (checkQuestionOptionToDelete != null && checkQuestionOptionToDelete.Count() > 0)
                                {
                                    _soxContext.RemoveRange(checkQuestionOptionToDelete);
                                    await _soxContext.SaveChangesAsync();
                                }

                                //Remove all fields associated with app id
                                var checkQuestionToDelete = _soxContext.KeyReportQuestion
                                    .Where(
                                        x => x.AppId.Equals(PodioAppKey.AppId));
                                if (checkQuestionToDelete != null && checkQuestionToDelete.Count() > 0)
                                {
                                    _soxContext.RemoveRange(checkQuestionToDelete);
                                    await _soxContext.SaveChangesAsync();
                                }
                                context.Commit();
                            }


                            foreach (var field in podioResult.Fields)
                            {
                                if (field.Status == "active")
                                {
                                    using (var context = _soxContext.Database.BeginTransaction())
                                    {
                                        //writeLog.Display(field);
                                        KeyReportQuestion keyReportField = new KeyReportQuestion();
                                        keyReportField.QuestionString = field.Label;
                                        keyReportField.Type = field.Type;
                                        keyReportField.AppId = PodioAppKey.AppId;
                                        keyReportField.FieldId = field.FieldId;
                                        keyReportField.CreatedOn = DateTime.Now;
                                        keyReportField.UpdatedOn = DateTime.Now;
                                        keyReportField.Position = position;
                                        List<KeyReportOption> listOptions = new List<KeyReportOption>();

                                        if (field.InternalConfig.Description != string.Empty)
                                        {
                                            keyReportField.Description = field.InternalConfig.Description;
                                        }

                                        if (field.Type == "category")
                                        {
                                            foreach (var option in field.InternalConfig.Settings)
                                            {
                                                if (option.Value.HasValues)
                                                {

                                                    foreach (var item in option.Value)
                                                    {
                                                        if (item["status"].ToString() == "active")
                                                        {
                                                            KeyReportOption keyReportOption = new KeyReportOption();
                                                            //Console.WriteLine($"{item["text"]}");
                                                            keyReportOption.OptionName = item["text"].ToString();
                                                            keyReportOption.CreatedOn = DateTime.Now;
                                                            keyReportOption.UpdatedOn = DateTime.Now;
                                                            keyReportOption.AppId = int.Parse(keyReportField.AppId);
                                                            keyReportOption.KeyReportQuestion = keyReportField;
                                                            keyReportOption.OptionId = $"{keyReportField.AppId}{count}";
                                                            listOptions.Add(keyReportOption);
                                                            count++;
                                                        }

                                                    }

                                                    keyReportField.Options = listOptions;
                                                }

                                            }
                                        }

                                        if (field.Type == "date")
                                        {
                                            Debug.WriteLine($"{field.InternalConfig.Settings["end"]}");
                                            if (field.InternalConfig.Settings["end"].ToString() == "enabled")
                                                keyReportField.Tag = "enabled";
                                            else
                                                keyReportField.Tag = string.Empty;

                                        }

                                        if (field.Type == "text")
                                        {
                                            if (field.InternalConfig.Settings["size"].ToString() == "large")
                                            {
                                                keyReportField.Tag = "large";
                                            }
                                        }

                                        if (field.Type == "image")
                                        {
                                            //image field 
                                        }

                                        if (field.Type == "app")
                                        {
                                            //image field 
                                        }


                                        _soxContext.Add(keyReportField);
                                        await _soxContext.SaveChangesAsync();
                                        context.Commit();
                                        position++;

                                    }
                                }
                            }

                        }

                        return Ok();
                    }

                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncKeyReportQuestionnaireStatusTrackerFields");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncKeyReportQuestionnaireStatusTrackerFields");
                return BadRequest(ex.ToString());
            }

            return NoContent();

        }

        //[AllowAnonymous]
        [HttpPost("keyreport/questionnaire/exception/fields")]
        public async Task<IActionResult> SyncKeyReportQuestionnaireExceptionFields()
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;    


                PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("ExceptionId").Value;
                PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("ExceptionToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(int.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                if (podio.IsAuthenticated())
                {
                    Debug.WriteLine("Podio Authenticated");
                    int count = 1;
                    int position = 1;
                    var podioResult = await podio.ApplicationService.GetApp(int.Parse(PodioAppKey.AppId));

                    if (podioResult.Fields != null && podioResult.Fields.Count > 0)
                    {

                        if (podioResult.Fields.Count > 0)
                        {

                            using (var context = _soxContext.Database.BeginTransaction())
                            {
                                //Remove all options associated with app id
                                var checkQuestionOptionToDelete = _soxContext.KeyReportOption.Where(id => id.AppId.Equals(int.Parse(PodioAppKey.AppId)));
                                if (checkQuestionOptionToDelete != null && checkQuestionOptionToDelete.Count() > 0)
                                {
                                    _soxContext.RemoveRange(checkQuestionOptionToDelete);
                                    await _soxContext.SaveChangesAsync();

                                }

                                //Remove all fields associated with app id
                                var checkQuestionToDelete = _soxContext.KeyReportQuestion
                                    .Where(
                                        x => x.AppId.Equals(PodioAppKey.AppId));
                                if (checkQuestionToDelete != null && checkQuestionToDelete.Count() > 0)
                                {
                                    _soxContext.RemoveRange(checkQuestionToDelete);
                                    await _soxContext.SaveChangesAsync();

                                }
                                context.Commit();
                            }


                            foreach (var field in podioResult.Fields)
                            {
                                if (field.Status == "active")
                                {
                                    using (var context = _soxContext.Database.BeginTransaction())
                                    {
                                        //writeLog.Display(field);
                                        KeyReportQuestion keyReportField = new KeyReportQuestion();
                                        keyReportField.QuestionString = field.Label;
                                        keyReportField.Type = field.Type;
                                        keyReportField.AppId = PodioAppKey.AppId;
                                        keyReportField.FieldId = field.FieldId;
                                        keyReportField.CreatedOn = DateTime.Now;
                                        keyReportField.UpdatedOn = DateTime.Now;
                                        keyReportField.Position = position;
                                        List<KeyReportOption> listOptions = new List<KeyReportOption>();

                                        if (field.InternalConfig.Description != string.Empty)
                                        {
                                            keyReportField.Description = field.InternalConfig.Description;
                                        }

                                        if (field.Type == "category")
                                        {
                                            foreach (var option in field.InternalConfig.Settings)
                                            {
                                                if (option.Value.HasValues)
                                                {

                                                    foreach (var item in option.Value)
                                                    {
                                                        if (item["status"].ToString() == "active")
                                                        {
                                                            KeyReportOption keyReportOption = new KeyReportOption();
                                                            //Console.WriteLine($"{item["text"]}");
                                                            keyReportOption.OptionName = item["text"].ToString();
                                                            keyReportOption.CreatedOn = DateTime.Now;
                                                            keyReportOption.UpdatedOn = DateTime.Now;
                                                            keyReportOption.AppId = int.Parse(keyReportField.AppId);
                                                            keyReportOption.KeyReportQuestion = keyReportField;
                                                            keyReportOption.OptionId = $"{keyReportField.AppId}{count}";
                                                            listOptions.Add(keyReportOption);
                                                            count++;
                                                        }

                                                    }

                                                    keyReportField.Options = listOptions;
                                                }

                                            }
                                        }

                                        if (field.Type == "date")
                                        {
                                            Debug.WriteLine($"{field.InternalConfig.Settings["end"]}");
                                            if (field.InternalConfig.Settings["end"].ToString() == "enabled")
                                                keyReportField.Tag = "enabled";
                                            else
                                                keyReportField.Tag = string.Empty;

                                        }

                                        if (field.Type == "text")
                                        {
                                            if (field.InternalConfig.Settings["size"].ToString() == "large")
                                            {
                                                keyReportField.Tag = "large";
                                            }
                                        }

                                        if (field.Type == "image")
                                        {
                                            //image field 
                                        }

                                        if (field.Type == "app")
                                        {
                                            //image field 
                                        }


                                        _soxContext.Add(keyReportField);
                                        await _soxContext.SaveChangesAsync();
                                        context.Commit();
                                        position++;

                                    }
                                }
                            }

                        }

                        return Ok();
                    }

                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncKeyReportQuestionnaireExceptionFields");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncKeyReportQuestionnaireExceptionFields");
                return BadRequest(ex.ToString());
            }

            return NoContent();

        }


        //[AllowAnonymous]
        [HttpPost("keyreport/consoleformat")]
        public async Task<IActionResult> SyncKeyReportQuestionnaireConsoleOrigs([FromBody] SyncDateRange syncDateRange)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();

                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("ConsolOrigFormatId").Value;
                PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("ConsolOrigFormatToken").Value;

                var response = await PullPodioKeyReport(PodioKey, PodioAppKey, syncDateRange);
                bool isValid = response.Item1;
                int totalItems = response.Item2;
                int total = response.Item3;
                if (isValid)
                {
                    //int total = syncDateRange.limit * (syncDateRange.offset + 1);
                    //string saveItems = $"{total}/{totalItems}";
                    //return Ok(new { syncDateRange.limit, totalItems, saveItems });
                    string saveItems = $"{total}/{totalItems}";
                    return Ok(new { syncDateRange.limit, totalItems, saveItems });
                }

                return NoContent();

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncKeyReportQuestionnaireConsoleOrigs");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncKeyReportQuestionnaireConsoleOrigs");
                return BadRequest(ex.ToString());
            }

        }

        //[AllowAnonymous]
        [HttpPost("keyreport/alliuc")]
        public async Task<IActionResult> SyncKeyReportQuestionnaireAllIUC([FromBody] SyncDateRange syncDateRange)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();

                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("AllIUCId").Value;
                PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("AllIUCToken").Value;

                //bool isValid = await PullPodioKeyReport(PodioKey, PodioAppKey, syncDateRange);
                //if(isValid)
                //{
                //    return Ok();
                //}

                var response = await PullPodioKeyReport(PodioKey, PodioAppKey, syncDateRange);
                bool isValid = response.Item1;
                int totalItems = response.Item2;
                int total = response.Item3;
                if (isValid)
                {
                    //int total = syncDateRange.limit * (syncDateRange.offset + 1);
                    //string saveItems = $"{total}/{totalItems}";
                    //return Ok(new { syncDateRange.limit, totalItems, saveItems });
                    string saveItems = $"{total}/{totalItems}";
                    return Ok(new { syncDateRange.limit, totalItems, saveItems });
                }

                return NoContent();

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncKeyReportQuestionnaireAllIUC");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncKeyReportQuestionnaireAllIUC");
                return BadRequest(ex.ToString());
            }
        }

        //[AllowAnonymous]
        [HttpPost("keyreport/statustracker")]
        public async Task<IActionResult> SyncKeyReportQuestionnaireStatusTracker([FromBody] SyncDateRange syncDateRange)
        {
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();

                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("TestStatusTrackerId").Value;
                PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("TestStatusTrackerToken").Value;

                //bool isValid = await PullPodioKeyReport(PodioKey, PodioAppKey, syncDateRange);
                //if (isValid)
                //{
                //    return Ok();
                //}

                var response = await PullPodioKeyReport(PodioKey, PodioAppKey, syncDateRange);
                bool isValid = response.Item1;
                int totalItems = response.Item2;
                int total = response.Item3;
                if (isValid)
                {
                    //int total = syncDateRange.limit * (syncDateRange.offset + 1);
                    //string saveItems = $"{total}/{totalItems}";
                    //return Ok(new { syncDateRange.limit, totalItems, saveItems });
                    string saveItems = $"{total}/{totalItems}";
                    return Ok(new { syncDateRange.limit, totalItems, saveItems });
                }

                return NoContent();

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncKeyReportQuestionnaireStatusTracker");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncKeyReportQuestionnaireStatusTracker");
                return BadRequest(ex.ToString());
            }

        }

        //[AllowAnonymous]
        [HttpPost("keyreport/exception")]
        public async Task<IActionResult> SyncKeyReportQuestionnaireException([FromBody] SyncDateRange syncDateRange)
        {
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();

                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("ExceptionId").Value;
                PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("ExceptionToken").Value;

                //bool isValid = await PullPodioKeyReport(PodioKey, PodioAppKey, syncDateRange);
                //if (isValid)
                //{
                //    return Ok();
                //}

                var response = await PullPodioKeyReport(PodioKey, PodioAppKey, syncDateRange);
                bool isValid = response.Item1;
                int totalItems = response.Item2;
                int total = response.Item3;
                if (isValid)
                {
                    //int total = syncDateRange.limit * (syncDateRange.offset + 1);
                    //string saveItems = $"{total}/{totalItems}";
                    //return Ok(new { syncDateRange.limit, totalItems, saveItems });
                    string saveItems = $"{total}/{totalItems}";
                    return Ok(new { syncDateRange.limit, totalItems, saveItems });
                }

                return NoContent();

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncKeyReportQuestionnaireException");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncKeyReportQuestionnaireException");
                return BadRequest(ex.ToString());
            }

        }

        //Key report podio references
        //[AllowAnonymous]
        [HttpPost("keyreport/KeyReportControlActivity")]
        public async Task<IActionResult> SyncKeyReportControlActivityAsync([FromBody] SyncDateRange syncDateRange)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("KeyReportControlActivityId").Value;
                PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("KeyReportControlActivityToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);
                FormatService txtFormat = new FormatService();

                if (podio.IsAuthenticated())
                {

                    var filter = new Dictionary<string, object>
                    {
                        {
                            "last_edit_on",
                            new  { from = syncDateRange.startDate, to = syncDateRange.endDate}
                        },
                    };

                    PodioCollection<Item> collectionCheck = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 1, 0, filters: filter, sortDesc: true);

                    if (collectionCheck.Filtered != 0)
                    {
                        int loop = (int)Math.Ceiling((double)collectionCheck.Filtered / 500);
                        int offset = 0;
                        List<KeyReportControlActivity> listKeyReportRef = new List<KeyReportControlActivity>();
                        List<KeyReportControlActivity> listKeyReportRefAdd = new List<KeyReportControlActivity>();
                        List<KeyReportControlActivity> listKeyReportRefUpdate = new List<KeyReportControlActivity>();
                        //get filtered items and stored in list
                        for (int i = 0; i < loop; i++)
                        {
                            PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 500, offset, filters: filter, sortDesc: true);

                            foreach (var item in collection.Items)
                            {
                                KeyReportControlActivity keyReportControlActivity = new KeyReportControlActivity();

                                #region Podio Item
                                keyReportControlActivity.PodioItemId = (int)item.ItemId;
                                keyReportControlActivity.PodioUniqueId = item.AppItemIdFormatted.ToString();
                                keyReportControlActivity.PodioRevision = item.CurrentRevision.Revision;
                                keyReportControlActivity.PodioLink = item.Link.ToString();
                                keyReportControlActivity.CreatedBy = item.CreatedBy.Name.ToString();
                                keyReportControlActivity.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());

                                #region Fields

                                //get process value
                                TextItemField textProcess = item.Field<TextItemField>(int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportControlActivityFieldId").GetSection("Field1").Value.ToString()));
                                keyReportControlActivity.ControlActivity = txtFormat.ReplaceTagHtmlParagraph(textProcess.Value, false);

                                #endregion

                                #endregion

                                listKeyReportRef.Add(keyReportControlActivity);
                            }

                            offset += 500;
                        }

                        if (listKeyReportRef != null && listKeyReportRef.Count() > 0)
                        {
                            using (var context = _soxContext.Database.BeginTransaction())
                            {

                                foreach (var item in listKeyReportRef)
                                {
                                    var itemCheck = _soxContext.KeyReportControlActivity.FirstOrDefault(id => id.PodioItemId == (int)item.PodioItemId);
                                    if (itemCheck != null)
                                    {

                                        itemCheck.ControlActivity = item.ControlActivity;
                                        itemCheck.PodioUniqueId = item.PodioUniqueId;
                                        itemCheck.PodioRevision = item.PodioRevision;
                                        itemCheck.PodioLink = item.PodioLink;

                                        listKeyReportRefUpdate.Add(itemCheck);
                                    }
                                    else
                                    {
                                        listKeyReportRefAdd.Add(item);
                                    }
                                }

                                if (listKeyReportRefAdd != null && listKeyReportRefAdd.Count() > 0)
                                    _soxContext.AddRange(listKeyReportRefAdd);

                                if (listKeyReportRefUpdate != null && listKeyReportRefUpdate.Count() > 0)
                                    _soxContext.UpdateRange(listKeyReportRefUpdate);

                                await _soxContext.SaveChangesAsync();
                                context.Commit();
                            }
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncKeyReportControlActivityAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncKeyReportControlActivityAsync");
                return BadRequest(ex.ToString());
            }

            return Ok();

        }


        //[AllowAnonymous]
        [HttpPost("keyreport/KeyReportKeyControl")]
        public async Task<IActionResult> SyncKeyReportKeyControlAsync([FromBody] SyncDateRange syncDateRange)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("KeyReportKeyControlId").Value;
                PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("KeyReportKeyControlToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                if (podio.IsAuthenticated())
                {

                    var filter = new Dictionary<string, object>
                    {
                        {
                            "last_edit_on",
                            new  { from = syncDateRange.startDate, to = syncDateRange.endDate}
                        },
                    };

                    PodioCollection<Item> collectionCheck = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 1, 0, filters: filter, sortDesc: true);

                    if (collectionCheck.Filtered != 0)
                    {
                        int loop = (int)Math.Ceiling((double)collectionCheck.Filtered / 500);
                        int offset = 0;
                        List<KeyReportKeyControl> listKeyReportRef = new List<KeyReportKeyControl>();
                        List<KeyReportKeyControl> listKeyReportRefAdd = new List<KeyReportKeyControl>();
                        List<KeyReportKeyControl> listKeyReportRefUpdate = new List<KeyReportKeyControl>();
                        //get filtered items and stored in list
                        for (int i = 0; i < loop; i++)
                        {
                            PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 500, offset, filters: filter, sortDesc: true);

                            foreach (var item in collection.Items)
                            {
                                KeyReportKeyControl keyReportControlActivity = new KeyReportKeyControl();

                                #region Podio Item
                                keyReportControlActivity.PodioItemId = (int)item.ItemId;
                                keyReportControlActivity.PodioUniqueId = item.AppItemIdFormatted.ToString();
                                keyReportControlActivity.PodioRevision = item.CurrentRevision.Revision;
                                keyReportControlActivity.PodioLink = item.Link.ToString();
                                keyReportControlActivity.CreatedBy = item.CreatedBy.Name.ToString();
                                keyReportControlActivity.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());

                                #region Fields

                                //get process value
                                TextItemField textProcess = item.Field<TextItemField>(int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportKeyControlFieldId").GetSection("Field1").Value.ToString()));
                                keyReportControlActivity.Key = textProcess.Value;

                                #endregion

                                #endregion

                                listKeyReportRef.Add(keyReportControlActivity);
                            }

                            offset += 500;
                        }

                        if (listKeyReportRef != null && listKeyReportRef.Count() > 0)
                        {
                            using (var context = _soxContext.Database.BeginTransaction())
                            {

                                foreach (var item in listKeyReportRef)
                                {
                                    var itemCheck = _soxContext.KeyReportKeyControl.FirstOrDefault(id => id.PodioItemId == (int)item.PodioItemId);
                                    if (itemCheck != null)
                                    {

                                        itemCheck.Key = item.Key;
                                        itemCheck.PodioUniqueId = item.PodioUniqueId;
                                        itemCheck.PodioRevision = item.PodioRevision;
                                        itemCheck.PodioLink = item.PodioLink;

                                        listKeyReportRefUpdate.Add(itemCheck);
                                    }
                                    else
                                    {
                                        listKeyReportRefAdd.Add(item);
                                    }
                                }

                                if (listKeyReportRefAdd != null && listKeyReportRefAdd.Count() > 0)
                                    _soxContext.AddRange(listKeyReportRefAdd);

                                if (listKeyReportRefUpdate != null && listKeyReportRefUpdate.Count() > 0)
                                    _soxContext.UpdateRange(listKeyReportRefUpdate);

                                await _soxContext.SaveChangesAsync();
                                context.Commit();
                            }
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncKeyReportKeyControlAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncKeyReportKeyControlAsync");
                return BadRequest(ex.ToString());
            }

            return Ok();

        }


        //[AllowAnonymous]
        [HttpPost("keyreport/KeyReportName")]
        public async Task<IActionResult> SyncKeyReportNameAsync([FromBody] SyncDateRange syncDateRange)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("KeyReportNameId").Value;
                PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("KeyReportNameToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                if (podio.IsAuthenticated())
                {

                    var filter = new Dictionary<string, object>
                    {
                        {
                            "last_edit_on",
                            new  { from = syncDateRange.startDate, to = syncDateRange.endDate}
                        },
                    };

                    PodioCollection<Item> collectionCheck = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 1, 0, filters: filter, sortDesc: true);

                    if (collectionCheck.Filtered != 0)
                    {
                        int loop = (int)Math.Ceiling((double)collectionCheck.Filtered / 500);
                        int offset = 0;
                        List<KeyReportName> listKeyReportRef = new List<KeyReportName>();
                        List<KeyReportName> listKeyReportRefAdd = new List<KeyReportName>();
                        List<KeyReportName> listKeyReportRefUpdate = new List<KeyReportName>();
                        //get filtered items and stored in list
                        for (int i = 0; i < loop; i++)
                        {
                            PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 500, offset, filters: filter, sortDesc: true);

                            foreach (var item in collection.Items)
                            {
                                KeyReportName keyReportControlActivity = new KeyReportName();

                                #region Podio Item
                                keyReportControlActivity.PodioItemId = (int)item.ItemId;
                                keyReportControlActivity.PodioUniqueId = item.AppItemIdFormatted.ToString();
                                keyReportControlActivity.PodioRevision = item.CurrentRevision.Revision;
                                keyReportControlActivity.PodioLink = item.Link.ToString();
                                keyReportControlActivity.CreatedBy = item.CreatedBy.Name.ToString();
                                keyReportControlActivity.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());

                                #region Fields

                                //get process value
                                TextItemField textProcess = item.Field<TextItemField>(int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportNameFieldId").GetSection("Field1").Value.ToString()));
                                keyReportControlActivity.Name = textProcess.Value;

                                #endregion

                                #endregion

                                listKeyReportRef.Add(keyReportControlActivity);
                            }

                            offset += 500;
                        }

                        if (listKeyReportRef != null && listKeyReportRef.Count() > 0)
                        {
                            using (var context = _soxContext.Database.BeginTransaction())
                            {

                                foreach (var item in listKeyReportRef)
                                {
                                    var itemCheck = _soxContext.KeyReportName.FirstOrDefault(id => id.PodioItemId == (int)item.PodioItemId);
                                    if (itemCheck != null)
                                    {

                                        itemCheck.Name = item.Name;
                                        itemCheck.PodioUniqueId = item.PodioUniqueId;
                                        itemCheck.PodioRevision = item.PodioRevision;
                                        itemCheck.PodioLink = item.PodioLink;

                                        listKeyReportRefUpdate.Add(itemCheck);
                                    }
                                    else
                                    {
                                        listKeyReportRefAdd.Add(item);
                                    }
                                }

                                if (listKeyReportRefAdd != null && listKeyReportRefAdd.Count() > 0)
                                    _soxContext.AddRange(listKeyReportRefAdd);

                                if (listKeyReportRefUpdate != null && listKeyReportRefUpdate.Count() > 0)
                                    _soxContext.UpdateRange(listKeyReportRefUpdate);

                                await _soxContext.SaveChangesAsync();
                                context.Commit();
                            }
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncKeyReportNameAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncKeyReportNameAsync");
                return BadRequest(ex.ToString());
            }

            return Ok();

        }


        //[AllowAnonymous]
        [HttpPost("keyreport/KeyReportSystemSource")]
        public async Task<IActionResult> SyncKeyReportSystemSourceAsync([FromBody] SyncDateRange syncDateRange)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("KeyReportSystemSourceId").Value;
                PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("KeyReportSystemSourceToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                if (podio.IsAuthenticated())
                {

                    var filter = new Dictionary<string, object>
                    {
                        {
                            "last_edit_on",
                            new  { from = syncDateRange.startDate, to = syncDateRange.endDate}
                        },
                    };

                    PodioCollection<Item> collectionCheck = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 1, 0, filters: filter, sortDesc: true);

                    if (collectionCheck.Filtered != 0)
                    {
                        int loop = (int)Math.Ceiling((double)collectionCheck.Filtered / 500);
                        int offset = 0;
                        List<KeyReportSystemSource> listKeyReportRef = new List<KeyReportSystemSource>();
                        List<KeyReportSystemSource> listKeyReportRefAdd = new List<KeyReportSystemSource>();
                        List<KeyReportSystemSource> listKeyReportRefUpdate = new List<KeyReportSystemSource>();
                        //get filtered items and stored in list
                        for (int i = 0; i < loop; i++)
                        {
                            PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 500, offset, filters: filter, sortDesc: true);

                            foreach (var item in collection.Items)
                            {
                                KeyReportSystemSource keyReportControlActivity = new KeyReportSystemSource();

                                #region Podio Item
                                keyReportControlActivity.PodioItemId = (int)item.ItemId;
                                keyReportControlActivity.PodioUniqueId = item.AppItemIdFormatted.ToString();
                                keyReportControlActivity.PodioRevision = item.CurrentRevision.Revision;
                                keyReportControlActivity.PodioLink = item.Link.ToString();
                                keyReportControlActivity.CreatedBy = item.CreatedBy.Name.ToString();
                                keyReportControlActivity.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());

                                #region Fields

                                //get process value
                                TextItemField textProcess = item.Field<TextItemField>(int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportSystemSourceFieldId").GetSection("Field1").Value.ToString()));
                                keyReportControlActivity.SystemSource = textProcess.Value;

                                CategoryItemField catStatus = item.Field<CategoryItemField>(int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportSystemSourceFieldId").GetSection("Field2").Value.ToString()));
                                IEnumerable<CategoryItemField.Option> listcatRisk = catStatus.Options;
                                keyReportControlActivity.SystemSourceCategory = listcatRisk.Select(x => x.Text).FirstOrDefault();

                                #endregion

                                #endregion

                                listKeyReportRef.Add(keyReportControlActivity);
                            }

                            offset += 500;
                        }

                        if (listKeyReportRef != null && listKeyReportRef.Count() > 0)
                        {
                            using (var context = _soxContext.Database.BeginTransaction())
                            {

                                foreach (var item in listKeyReportRef)
                                {
                                    var itemCheck = _soxContext.KeyReportSystemSource.FirstOrDefault(id => id.PodioItemId == (int)item.PodioItemId);
                                    if (itemCheck != null)
                                    {

                                        itemCheck.SystemSource = item.SystemSource;
                                        itemCheck.SystemSourceCategory = item.SystemSourceCategory;
                                        itemCheck.PodioUniqueId = item.PodioUniqueId;
                                        itemCheck.PodioRevision = item.PodioRevision;
                                        itemCheck.PodioLink = item.PodioLink;

                                        listKeyReportRefUpdate.Add(itemCheck);
                                    }
                                    else
                                    {
                                        listKeyReportRefAdd.Add(item);
                                    }
                                }

                                if (listKeyReportRefAdd != null && listKeyReportRefAdd.Count() > 0)
                                    _soxContext.AddRange(listKeyReportRefAdd);

                                if (listKeyReportRefUpdate != null && listKeyReportRefUpdate.Count() > 0)
                                    _soxContext.UpdateRange(listKeyReportRefUpdate);

                                await _soxContext.SaveChangesAsync();
                                context.Commit();
                            }
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncKeyReportSystemSourceAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncKeyReportSystemSourceAsync");
                return BadRequest(ex.ToString());
            }

            return Ok();

        }


        //[AllowAnonymous]
        [HttpPost("keyreport/KeyReportNonKeyReport")]
        public async Task<IActionResult> SyncKeyReportNonKeyReportAsync([FromBody] SyncDateRange syncDateRange)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("KeyReportNonKeyReportId").Value;
                PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("KeyReportNonKeyReportToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                if (podio.IsAuthenticated())
                {

                    var filter = new Dictionary<string, object>
                    {
                        {
                            "last_edit_on",
                            new  { from = syncDateRange.startDate, to = syncDateRange.endDate}
                        },
                    };

                    PodioCollection<Item> collectionCheck = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 1, 0, filters: filter, sortDesc: true);

                    if (collectionCheck.Filtered != 0)
                    {
                        int loop = (int)Math.Ceiling((double)collectionCheck.Filtered / 500);
                        int offset = 0;
                        List<KeyReportNonKeyReport> listKeyReportRef = new List<KeyReportNonKeyReport>();
                        List<KeyReportNonKeyReport> listKeyReportRefAdd = new List<KeyReportNonKeyReport>();
                        List<KeyReportNonKeyReport> listKeyReportRefUpdate = new List<KeyReportNonKeyReport>();
                        //get filtered items and stored in list
                        for (int i = 0; i < loop; i++)
                        {
                            PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 500, offset, filters: filter, sortDesc: true);

                            foreach (var item in collection.Items)
                            {
                                KeyReportNonKeyReport keyReportControlActivity = new KeyReportNonKeyReport();

                                #region Podio Item
                                keyReportControlActivity.PodioItemId = (int)item.ItemId;
                                keyReportControlActivity.PodioUniqueId = item.AppItemIdFormatted.ToString();
                                keyReportControlActivity.PodioRevision = item.CurrentRevision.Revision;
                                keyReportControlActivity.PodioLink = item.Link.ToString();
                                keyReportControlActivity.CreatedBy = item.CreatedBy.Name.ToString();
                                keyReportControlActivity.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());

                                #region Fields

                                //get process value
                                TextItemField textProcess = item.Field<TextItemField>(int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportNonKeyReportFieldId").GetSection("Field1").Value.ToString()));
                                keyReportControlActivity.Report = textProcess.Value;

                                #endregion

                                #endregion

                                listKeyReportRef.Add(keyReportControlActivity);
                            }

                            offset += 500;
                        }

                        if (listKeyReportRef != null && listKeyReportRef.Count() > 0)
                        {
                            using (var context = _soxContext.Database.BeginTransaction())
                            {

                                foreach (var item in listKeyReportRef)
                                {
                                    var itemCheck = _soxContext.KeyReportNonKeyReport.FirstOrDefault(id => id.PodioItemId == (int)item.PodioItemId);
                                    if (itemCheck != null)
                                    {

                                        itemCheck.Report = item.Report;
                                        itemCheck.PodioUniqueId = item.PodioUniqueId;
                                        itemCheck.PodioRevision = item.PodioRevision;
                                        itemCheck.PodioLink = item.PodioLink;

                                        listKeyReportRefUpdate.Add(itemCheck);
                                    }
                                    else
                                    {
                                        listKeyReportRefAdd.Add(item);
                                    }
                                }

                                if (listKeyReportRefAdd != null && listKeyReportRefAdd.Count() > 0)
                                    _soxContext.AddRange(listKeyReportRefAdd);

                                if (listKeyReportRefUpdate != null && listKeyReportRefUpdate.Count() > 0)
                                    _soxContext.UpdateRange(listKeyReportRefUpdate);

                                await _soxContext.SaveChangesAsync();
                                context.Commit();
                            }
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncKeyReportNonKeyReportAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncKeyReportNonKeyReportAsync");
                return BadRequest(ex.ToString());
            }

            return Ok();

        }


        //[AllowAnonymous]
        [HttpPost("keyreport/KeyReportReportCustomized")]
        public async Task<IActionResult> SyncKeyReportReportCustomizedAsync([FromBody] SyncDateRange syncDateRange)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("KeyReportReportCustomizedId").Value;
                PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("KeyReportReportCustomizedToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                if (podio.IsAuthenticated())
                {

                    var filter = new Dictionary<string, object>
                    {
                        {
                            "last_edit_on",
                            new  { from = syncDateRange.startDate, to = syncDateRange.endDate}
                        },
                    };

                    PodioCollection<Item> collectionCheck = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 1, 0, filters: filter, sortDesc: true);

                    if (collectionCheck.Filtered != 0)
                    {
                        int loop = (int)Math.Ceiling((double)collectionCheck.Filtered / 500);
                        int offset = 0;
                        List<KeyReportReportCustomized> listKeyReportRef = new List<KeyReportReportCustomized>();
                        List<KeyReportReportCustomized> listKeyReportRefAdd = new List<KeyReportReportCustomized>();
                        List<KeyReportReportCustomized> listKeyReportRefUpdate = new List<KeyReportReportCustomized>();
                        //get filtered items and stored in list
                        for (int i = 0; i < loop; i++)
                        {
                            PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 500, offset, filters: filter, sortDesc: true);

                            foreach (var item in collection.Items)
                            {
                                KeyReportReportCustomized keyReportControlActivity = new KeyReportReportCustomized();

                                #region Podio Item
                                keyReportControlActivity.PodioItemId = (int)item.ItemId;
                                keyReportControlActivity.PodioUniqueId = item.AppItemIdFormatted.ToString();
                                keyReportControlActivity.PodioRevision = item.CurrentRevision.Revision;
                                keyReportControlActivity.PodioLink = item.Link.ToString();
                                keyReportControlActivity.CreatedBy = item.CreatedBy.Name.ToString();
                                keyReportControlActivity.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());

                                #region Fields

                                //get process value
                                //TextItemField textProcess = item.Field<TextItemField>(int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportReportCustomizedFieldId").GetSection("Field1").Value.ToString()));
                                //keyReportControlActivity.ReportCustomized = textProcess.Value;

                                CategoryItemField catStatus = item.Field<CategoryItemField>(int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportReportCustomizedFieldId").GetSection("Field1").Value.ToString()));
                                IEnumerable<CategoryItemField.Option> listcatRisk = catStatus.Options;
                                keyReportControlActivity.ReportCustomized = listcatRisk.Select(x => x.Text).FirstOrDefault();

                                #endregion

                                #endregion

                                listKeyReportRef.Add(keyReportControlActivity);
                            }

                            offset += 500;
                        }

                        if (listKeyReportRef != null && listKeyReportRef.Count() > 0)
                        {
                            using (var context = _soxContext.Database.BeginTransaction())
                            {

                                foreach (var item in listKeyReportRef)
                                {
                                    var itemCheck = _soxContext.KeyReportReportCustomized.FirstOrDefault(id => id.PodioItemId == (int)item.PodioItemId);
                                    if (itemCheck != null)
                                    {

                                        itemCheck.ReportCustomized = item.ReportCustomized;
                                        itemCheck.PodioUniqueId = item.PodioUniqueId;
                                        itemCheck.PodioRevision = item.PodioRevision;
                                        itemCheck.PodioLink = item.PodioLink;

                                        listKeyReportRefUpdate.Add(itemCheck);
                                    }
                                    else
                                    {
                                        listKeyReportRefAdd.Add(item);
                                    }
                                }

                                if (listKeyReportRefAdd != null && listKeyReportRefAdd.Count() > 0)
                                    _soxContext.AddRange(listKeyReportRefAdd);

                                if (listKeyReportRefUpdate != null && listKeyReportRefUpdate.Count() > 0)
                                    _soxContext.UpdateRange(listKeyReportRefUpdate);

                                await _soxContext.SaveChangesAsync();
                                context.Commit();
                            }
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncKeyReportReportCustomizedAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncKeyReportReportCustomizedAsync");
                return BadRequest(ex.ToString());
            }

            return Ok();

        }


        //[AllowAnonymous]
        [HttpPost("keyreport/KeyReportIUCType")]
        public async Task<IActionResult> SyncKeyReportIUCTypeAsync([FromBody] SyncDateRange syncDateRange)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("KeyReportIUCTypeId").Value;
                PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("KeyReportIUCTypeToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                if (podio.IsAuthenticated())
                {

                    var filter = new Dictionary<string, object>
                    {
                        {
                            "last_edit_on",
                            new  { from = syncDateRange.startDate, to = syncDateRange.endDate}
                        },
                    };

                    PodioCollection<Item> collectionCheck = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 1, 0, filters: filter, sortDesc: true);

                    if (collectionCheck.Filtered != 0)
                    {
                        int loop = (int)Math.Ceiling((double)collectionCheck.Filtered / 500);
                        int offset = 0;
                        List<KeyReportIUCType> listKeyReportRef = new List<KeyReportIUCType>();
                        List<KeyReportIUCType> listKeyReportRefAdd = new List<KeyReportIUCType>();
                        List<KeyReportIUCType> listKeyReportRefUpdate = new List<KeyReportIUCType>();
                        //get filtered items and stored in list
                        for (int i = 0; i < loop; i++)
                        {
                            PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 500, offset, filters: filter, sortDesc: true);

                            foreach (var item in collection.Items)
                            {
                                KeyReportIUCType keyReportControlActivity = new KeyReportIUCType();

                                #region Podio Item
                                keyReportControlActivity.PodioItemId = (int)item.ItemId;
                                keyReportControlActivity.PodioUniqueId = item.AppItemIdFormatted.ToString();
                                keyReportControlActivity.PodioRevision = item.CurrentRevision.Revision;
                                keyReportControlActivity.PodioLink = item.Link.ToString();
                                keyReportControlActivity.CreatedBy = item.CreatedBy.Name.ToString();
                                keyReportControlActivity.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());

                                #region Fields

                                //get process value
                                TextItemField textProcess = item.Field<TextItemField>(int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportIUCTypeFieldId").GetSection("Field1").Value.ToString()));
                                keyReportControlActivity.IUCType = textProcess.Value;

                                #endregion

                                #endregion

                                listKeyReportRef.Add(keyReportControlActivity);
                            }

                            offset += 500;
                        }

                        if (listKeyReportRef != null && listKeyReportRef.Count() > 0)
                        {
                            using (var context = _soxContext.Database.BeginTransaction())
                            {

                                foreach (var item in listKeyReportRef)
                                {
                                    var itemCheck = _soxContext.KeyReportIUCType.FirstOrDefault(id => id.PodioItemId == (int)item.PodioItemId);
                                    if (itemCheck != null)
                                    {

                                        itemCheck.IUCType = item.IUCType;
                                        itemCheck.PodioUniqueId = item.PodioUniqueId;
                                        itemCheck.PodioRevision = item.PodioRevision;
                                        itemCheck.PodioLink = item.PodioLink;

                                        listKeyReportRefUpdate.Add(itemCheck);
                                    }
                                    else
                                    {
                                        listKeyReportRefAdd.Add(item);
                                    }
                                }

                                if (listKeyReportRefAdd != null && listKeyReportRefAdd.Count() > 0)
                                    _soxContext.AddRange(listKeyReportRefAdd);

                                if (listKeyReportRefUpdate != null && listKeyReportRefUpdate.Count() > 0)
                                    _soxContext.UpdateRange(listKeyReportRefUpdate);

                                await _soxContext.SaveChangesAsync();
                                context.Commit();
                            }
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncKeyReportIUCTypeAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncKeyReportIUCTypeAsync");
                return BadRequest(ex.ToString());
            }

            return Ok();

        }


        //[AllowAnonymous]
        [HttpPost("keyreport/KeyReportControlsRelyingIUC")]
        public async Task<IActionResult> SyncKeyReportControlsRelyingIUCAsync([FromBody] SyncDateRange syncDateRange)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("KeyReportControlsRelyingIUCId").Value;
                PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("KeyReportControlsRelyingIUCToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                if (podio.IsAuthenticated())
                {

                    var filter = new Dictionary<string, object>
                    {
                        {
                            "last_edit_on",
                            new  { from = syncDateRange.startDate, to = syncDateRange.endDate}
                        },
                    };

                    PodioCollection<Item> collectionCheck = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 1, 0, filters: filter, sortDesc: true);

                    if (collectionCheck.Filtered != 0)
                    {
                        int loop = (int)Math.Ceiling((double)collectionCheck.Filtered / 500);
                        int offset = 0;
                        List<KeyReportControlsRelyingIUC> listKeyReportRef = new List<KeyReportControlsRelyingIUC>();
                        List<KeyReportControlsRelyingIUC> listKeyReportRefAdd = new List<KeyReportControlsRelyingIUC>();
                        List<KeyReportControlsRelyingIUC> listKeyReportRefUpdate = new List<KeyReportControlsRelyingIUC>();
                        //get filtered items and stored in list
                        for (int i = 0; i < loop; i++)
                        {
                            PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 500, offset, filters: filter, sortDesc: true);

                            foreach (var item in collection.Items)
                            {
                                KeyReportControlsRelyingIUC keyReportControlActivity = new KeyReportControlsRelyingIUC();

                                #region Podio Item
                                keyReportControlActivity.PodioItemId = (int)item.ItemId;
                                keyReportControlActivity.PodioUniqueId = item.AppItemIdFormatted.ToString();
                                keyReportControlActivity.PodioRevision = item.CurrentRevision.Revision;
                                keyReportControlActivity.PodioLink = item.Link.ToString();
                                keyReportControlActivity.CreatedBy = item.CreatedBy.Name.ToString();
                                keyReportControlActivity.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());

                                #region Fields

                                //get process value
                                TextItemField textProcess = item.Field<TextItemField>(int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportControlsRelyingIUCFieldId").GetSection("Field1").Value.ToString()));
                                keyReportControlActivity.ControlsRelyingIUC = textProcess.Value;

                                #endregion

                                #endregion

                                listKeyReportRef.Add(keyReportControlActivity);
                            }

                            offset += 500;
                        }

                        if (listKeyReportRef != null && listKeyReportRef.Count() > 0)
                        {
                            using (var context = _soxContext.Database.BeginTransaction())
                            {

                                foreach (var item in listKeyReportRef)
                                {
                                    var itemCheck = _soxContext.KeyReportControlsRelyingIUC.FirstOrDefault(id => id.PodioItemId == (int)item.PodioItemId);
                                    if (itemCheck != null)
                                    {

                                        itemCheck.ControlsRelyingIUC = item.ControlsRelyingIUC;
                                        itemCheck.PodioUniqueId = item.PodioUniqueId;
                                        itemCheck.PodioRevision = item.PodioRevision;
                                        itemCheck.PodioLink = item.PodioLink;

                                        listKeyReportRefUpdate.Add(itemCheck);
                                    }
                                    else
                                    {
                                        listKeyReportRefAdd.Add(item);
                                    }
                                }

                                if (listKeyReportRefAdd != null && listKeyReportRefAdd.Count() > 0)
                                    _soxContext.AddRange(listKeyReportRefAdd);

                                if (listKeyReportRefUpdate != null && listKeyReportRefUpdate.Count() > 0)
                                    _soxContext.UpdateRange(listKeyReportRefUpdate);

                                await _soxContext.SaveChangesAsync();
                                context.Commit();
                            }
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncKeyReportControlsRelyingIUCAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncKeyReportControlsRelyingIUCAsync");
                return BadRequest(ex.ToString());
            }

            return Ok();

        }


        //[AllowAnonymous]
        [HttpPost("keyreport/KeyReportPreparer")]
        public async Task<IActionResult> SyncKeyReportPreparerAsync([FromBody] SyncDateRange syncDateRange)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("KeyReportPreparerId").Value;
                PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("KeyReportPreparerToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                if (podio.IsAuthenticated())
                {

                    var filter = new Dictionary<string, object>
                    {
                        {
                            "last_edit_on",
                            new  { from = syncDateRange.startDate, to = syncDateRange.endDate}
                        },
                    };

                    PodioCollection<Item> collectionCheck = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 1, 0, filters: filter, sortDesc: true);

                    if (collectionCheck.Filtered != 0)
                    {
                        int loop = (int)Math.Ceiling((double)collectionCheck.Filtered / 500);
                        int offset = 0;
                        List<KeyReportPreparer> listKeyReportRef = new List<KeyReportPreparer>();
                        List<KeyReportPreparer> listKeyReportRefAdd = new List<KeyReportPreparer>();
                        List<KeyReportPreparer> listKeyReportRefUpdate = new List<KeyReportPreparer>();
                        //get filtered items and stored in list
                        for (int i = 0; i < loop; i++)
                        {
                            PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 500, offset, filters: filter, sortDesc: true);

                            foreach (var item in collection.Items)
                            {
                                KeyReportPreparer keyReportControlActivity = new KeyReportPreparer();

                                #region Podio Item
                                keyReportControlActivity.PodioItemId = (int)item.ItemId;
                                keyReportControlActivity.PodioUniqueId = item.AppItemIdFormatted.ToString();
                                keyReportControlActivity.PodioRevision = item.CurrentRevision.Revision;
                                keyReportControlActivity.PodioLink = item.Link.ToString();
                                keyReportControlActivity.CreatedBy = item.CreatedBy.Name.ToString();
                                keyReportControlActivity.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());

                                #region Fields

                                //get process value
                                TextItemField textProcess = item.Field<TextItemField>(int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportPreparerFieldId").GetSection("Field1").Value.ToString()));
                                keyReportControlActivity.Preparer = textProcess.Value;

                                #endregion

                                #endregion

                                listKeyReportRef.Add(keyReportControlActivity);
                            }

                            offset += 500;
                        }

                        if (listKeyReportRef != null && listKeyReportRef.Count() > 0)
                        {
                            using (var context = _soxContext.Database.BeginTransaction())
                            {

                                foreach (var item in listKeyReportRef)
                                {
                                    var itemCheck = _soxContext.KeyReportPreparer.FirstOrDefault(id => id.PodioItemId == (int)item.PodioItemId);
                                    if (itemCheck != null)
                                    {

                                        itemCheck.Preparer = item.Preparer;
                                        itemCheck.PodioUniqueId = item.PodioUniqueId;
                                        itemCheck.PodioRevision = item.PodioRevision;
                                        itemCheck.PodioLink = item.PodioLink;

                                        listKeyReportRefUpdate.Add(itemCheck);
                                    }
                                    else
                                    {
                                        listKeyReportRefAdd.Add(item);
                                    }
                                }

                                if (listKeyReportRefAdd != null && listKeyReportRefAdd.Count() > 0)
                                    _soxContext.AddRange(listKeyReportRefAdd);

                                if (listKeyReportRefUpdate != null && listKeyReportRefUpdate.Count() > 0)
                                    _soxContext.UpdateRange(listKeyReportRefUpdate);

                                await _soxContext.SaveChangesAsync();
                                context.Commit();
                            }
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncKeyReportPreparerAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncKeyReportPreparerAsync");
                return BadRequest(ex.ToString());
            }

            return Ok();

        }



        //[AllowAnonymous]
        [HttpPost("keyreport/KeyReportUniqueKeyReport")]
        public async Task<IActionResult> SyncKeyReportUniqueKeyReportAsync([FromBody] SyncDateRange syncDateRange)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("KeyReportUniqueKeyReportId").Value;
                PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("KeyReportUniqueKeyReportToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                if (podio.IsAuthenticated())
                {

                    var filter = new Dictionary<string, object>
                    {
                        {
                            "last_edit_on",
                            new  { from = syncDateRange.startDate, to = syncDateRange.endDate}
                        },
                    };

                    PodioCollection<Item> collectionCheck = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 1, 0, filters: filter, sortDesc: true);

                    if (collectionCheck.Filtered != 0)
                    {
                        int loop = (int)Math.Ceiling((double)collectionCheck.Filtered / 500);
                        int offset = 0;
                        List<KeyReportUniqueKeyReport> listKeyReportRef = new List<KeyReportUniqueKeyReport>();
                        List<KeyReportUniqueKeyReport> listKeyReportRefAdd = new List<KeyReportUniqueKeyReport>();
                        List<KeyReportUniqueKeyReport> listKeyReportRefUpdate = new List<KeyReportUniqueKeyReport>();
                        FormatService formatService = new FormatService();
                        //get filtered items and stored in list
                        for (int i = 0; i < loop; i++)
                        {
                            PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 500, offset, filters: filter, sortDesc: true);

                            foreach (var item in collection.Items)
                            {
                                KeyReportUniqueKeyReport keyReportControlActivity = new KeyReportUniqueKeyReport();

                                #region Podio Item
                                keyReportControlActivity.PodioItemId = (int)item.ItemId;
                                keyReportControlActivity.PodioUniqueId = item.AppItemIdFormatted.ToString();
                                keyReportControlActivity.PodioRevision = item.CurrentRevision.Revision;
                                keyReportControlActivity.PodioLink = item.Link.ToString();
                                keyReportControlActivity.CreatedBy = item.CreatedBy.Name.ToString();
                                keyReportControlActivity.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());

                                #region Fields

                                //CategoryItemField catStatus = item.Field<CategoryItemField>(int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportUniqueKeyReportFieldId").GetSection("Field1").Value.ToString()));
                                //IEnumerable<CategoryItemField.Option> listcatRisk = catStatus.Options;
                                //keyReportControlActivity.UniqueKeyReport = listcatRisk.Select(x => x.Text).FirstOrDefault();

                                TextItemField textProcess = item.Field<TextItemField>(int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportUniqueKeyReportFieldId").GetSection("Field1").Value.ToString()));
                                keyReportControlActivity.UniqueKeyReport = textProcess.Value != string.Empty ? formatService.ReplaceTagHtmlParagraph(textProcess.Value, false) : string.Empty;

                                #endregion

                                #endregion

                                listKeyReportRef.Add(keyReportControlActivity);
                            }

                            offset += 500;
                        }

                        if (listKeyReportRef != null && listKeyReportRef.Count() > 0)
                        {
                            using (var context = _soxContext.Database.BeginTransaction())
                            {

                                foreach (var item in listKeyReportRef)
                                {
                                    var itemCheck = _soxContext.KeyReportUniqueKeyReport.FirstOrDefault(id => id.PodioItemId == (int)item.PodioItemId);
                                    if (itemCheck != null)
                                    {

                                        itemCheck.UniqueKeyReport = item.UniqueKeyReport;
                                        itemCheck.PodioUniqueId = item.PodioUniqueId;
                                        itemCheck.PodioRevision = item.PodioRevision;
                                        itemCheck.PodioLink = item.PodioLink;

                                        listKeyReportRefUpdate.Add(itemCheck);
                                    }
                                    else
                                    {
                                        listKeyReportRefAdd.Add(item);
                                    }
                                }

                                if (listKeyReportRefAdd != null && listKeyReportRefAdd.Count() > 0)
                                    _soxContext.AddRange(listKeyReportRefAdd);

                                if (listKeyReportRefUpdate != null && listKeyReportRefUpdate.Count() > 0)
                                    _soxContext.UpdateRange(listKeyReportRefUpdate);

                                await _soxContext.SaveChangesAsync();
                                context.Commit();
                            }
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncKeyReportUniqueKeyReportAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncKeyReportUniqueKeyReportAsync");
                return BadRequest(ex.ToString());
            }

            return Ok();

        }


        [AllowAnonymous]
        [HttpPost("keyreport/KeyReportNotes")]
        public async Task<IActionResult> SyncKeyReportNotesAsync([FromBody] SyncDateRange syncDateRange)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("KeyReportNotesId").Value;
                PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("KeyReportNotesToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                if (podio.IsAuthenticated())
                {

                    var filter = new Dictionary<string, object>
                    {
                        {
                            "last_edit_on",
                            new  { from = syncDateRange.startDate, to = syncDateRange.endDate}
                        },
                    };

                    PodioCollection<Item> collectionCheck = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 1, 0, filters: filter, sortDesc: true);

                    if (collectionCheck.Filtered != 0)
                    {
                        int loop = (int)Math.Ceiling((double)collectionCheck.Filtered / 500);
                        int offset = 0;
                        List<KeyReportNotes> listKeyReportRef = new List<KeyReportNotes>();
                        List<KeyReportNotes> listKeyReportRefAdd = new List<KeyReportNotes>();
                        List<KeyReportNotes> listKeyReportRefUpdate = new List<KeyReportNotes>();
                        //get filtered items and stored in list
                        for (int i = 0; i < loop; i++)
                        {
                            PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 500, offset, filters: filter, sortDesc: true);

                            foreach (var item in collection.Items)
                            {
                                KeyReportNotes keyReportControlActivity = new KeyReportNotes();

                                #region Podio Item
                                keyReportControlActivity.PodioItemId = (int)item.ItemId;
                                keyReportControlActivity.PodioUniqueId = item.AppItemIdFormatted.ToString();
                                keyReportControlActivity.PodioRevision = item.CurrentRevision.Revision;
                                keyReportControlActivity.PodioLink = item.Link.ToString();
                                keyReportControlActivity.CreatedBy = item.CreatedBy.Name.ToString();
                                keyReportControlActivity.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());

                                #region Fields

                                //get process value
                                TextItemField textProcess = item.Field<TextItemField>(int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportNotesFieldId").GetSection("Field1").Value.ToString()));
                                keyReportControlActivity.ReportNotes = textProcess.Value;

                                #endregion

                                #endregion

                                listKeyReportRef.Add(keyReportControlActivity);
                            }

                            offset += 500;
                        }

                        if (listKeyReportRef != null && listKeyReportRef.Count() > 0)
                        {
                            using (var context = _soxContext.Database.BeginTransaction())
                            {

                                foreach (var item in listKeyReportRef)
                                {
                                    var itemCheck = _soxContext.KeyReportNotes.FirstOrDefault(id => id.PodioItemId == (int)item.PodioItemId);
                                    if (itemCheck != null)
                                    {

                                        itemCheck.ReportNotes = item.ReportNotes;
                                        itemCheck.PodioUniqueId = item.PodioUniqueId;
                                        itemCheck.PodioRevision = item.PodioRevision;
                                        itemCheck.PodioLink = item.PodioLink;

                                        listKeyReportRefUpdate.Add(itemCheck);
                                    }
                                    else
                                    {
                                        listKeyReportRefAdd.Add(item);
                                    }
                                }

                                if (listKeyReportRefAdd != null && listKeyReportRefAdd.Count() > 0)
                                    _soxContext.AddRange(listKeyReportRefAdd);

                                if (listKeyReportRefUpdate != null && listKeyReportRefUpdate.Count() > 0)
                                    _soxContext.UpdateRange(listKeyReportRefUpdate);

                                await _soxContext.SaveChangesAsync();
                                context.Commit();
                            }
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncKeyReportNotesAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncKeyReportNotesAsync");
                return BadRequest(ex.ToString());
            }

            return Ok();

        }


        //[AllowAnonymous]
        [HttpPost("keyreport/KeyReportNumber")]
        public async Task<IActionResult> SyncKeyKeyReportNumberAsync([FromBody] SyncDateRange syncDateRange)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("KeyReportNumberId").Value;
                PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("KeyReportNumberToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                if (podio.IsAuthenticated())
                {

                    var filter = new Dictionary<string, object>
                    {
                        {
                            "last_edit_on",
                            new  { from = syncDateRange.startDate, to = syncDateRange.endDate}
                        },
                    };

                    PodioCollection<Item> collectionCheck = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 1, 0, filters: filter, sortDesc: true);

                    if (collectionCheck.Filtered != 0)
                    {
                        int loop = (int)Math.Ceiling((double)collectionCheck.Filtered / 500);
                        int offset = 0;
                        List<KeyReportNumber> listKeyReportRef = new List<KeyReportNumber>();
                        List<KeyReportNumber> listKeyReportRefAdd = new List<KeyReportNumber>();
                        List<KeyReportNumber> listKeyReportRefUpdate = new List<KeyReportNumber>();
                        //get filtered items and stored in list
                        for (int i = 0; i < loop; i++)
                        {
                            PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 500, offset, filters: filter, sortDesc: true);

                            foreach (var item in collection.Items)
                            {
                                KeyReportNumber keyReportControlActivity = new KeyReportNumber();

                                #region Podio Item
                                keyReportControlActivity.PodioItemId = (int)item.ItemId;
                                keyReportControlActivity.PodioUniqueId = item.AppItemIdFormatted.ToString();
                                keyReportControlActivity.PodioRevision = item.CurrentRevision.Revision;
                                keyReportControlActivity.PodioLink = item.Link.ToString();
                                keyReportControlActivity.CreatedBy = item.CreatedBy.Name.ToString();
                                keyReportControlActivity.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());

                                #region Fields

                                //get process value
                                TextItemField textProcess = item.Field<TextItemField>(int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportNumberFieldId").GetSection("Field1").Value.ToString()));
                                keyReportControlActivity.ReportNumber = textProcess.Value;

                                #endregion

                                #endregion

                                listKeyReportRef.Add(keyReportControlActivity);
                            }

                            offset += 500;
                        }

                        if (listKeyReportRef != null && listKeyReportRef.Count() > 0)
                        {
                            using (var context = _soxContext.Database.BeginTransaction())
                            {

                                foreach (var item in listKeyReportRef)
                                {
                                    var itemCheck = _soxContext.KeyReportNumber.FirstOrDefault(id => id.PodioItemId == (int)item.PodioItemId);
                                    if (itemCheck != null)
                                    {

                                        itemCheck.ReportNumber = item.ReportNumber;
                                        itemCheck.PodioUniqueId = item.PodioUniqueId;
                                        itemCheck.PodioRevision = item.PodioRevision;
                                        itemCheck.PodioLink = item.PodioLink;

                                        listKeyReportRefUpdate.Add(itemCheck);
                                    }
                                    else
                                    {
                                        listKeyReportRefAdd.Add(item);
                                    }
                                }

                                if (listKeyReportRefAdd != null && listKeyReportRefAdd.Count() > 0)
                                    _soxContext.AddRange(listKeyReportRefAdd);

                                if (listKeyReportRefUpdate != null && listKeyReportRefUpdate.Count() > 0)
                                    _soxContext.UpdateRange(listKeyReportRefUpdate);

                                await _soxContext.SaveChangesAsync();
                                context.Commit();
                            }
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncKeyReportNumberAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncKeyReportNumberAsync");
                return BadRequest(ex.ToString());
            }

            return Ok();

        }


        //[AllowAnonymous]
        [HttpPost("keyreport/KeyReportTester")]
        public async Task<IActionResult> SyncKeyReportTesterAsync([FromBody] SyncDateRange syncDateRange)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("KeyReportTesterId").Value;
                PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("KeyReportTesterToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                if (podio.IsAuthenticated())
                {

                    var filter = new Dictionary<string, object>
                    {
                        {
                            "last_edit_on",
                            new  { from = syncDateRange.startDate, to = syncDateRange.endDate}
                        },
                    };

                    PodioCollection<Item> collectionCheck = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 1, 0, filters: filter, sortDesc: true);

                    if (collectionCheck.Filtered != 0)
                    {
                        int loop = (int)Math.Ceiling((double)collectionCheck.Filtered / 500);
                        int offset = 0;
                        List<KeyReportTester> listKeyReportRef = new List<KeyReportTester>();
                        List<KeyReportTester> listKeyReportRefAdd = new List<KeyReportTester>();
                        List<KeyReportTester> listKeyReportRefUpdate = new List<KeyReportTester>();
                        //get filtered items and stored in list
                        for (int i = 0; i < loop; i++)
                        {
                            PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 500, offset, filters: filter, sortDesc: true);

                            foreach (var item in collection.Items)
                            {
                                KeyReportTester keyReportControlActivity = new KeyReportTester();

                                #region Podio Item
                                keyReportControlActivity.PodioItemId = (int)item.ItemId;
                                keyReportControlActivity.PodioUniqueId = item.AppItemIdFormatted.ToString();
                                keyReportControlActivity.PodioRevision = item.CurrentRevision.Revision;
                                keyReportControlActivity.PodioLink = item.Link.ToString();
                                keyReportControlActivity.CreatedBy = item.CreatedBy.Name.ToString();
                                keyReportControlActivity.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());

                                #region Fields

                                //get process value
                                TextItemField textProcess = item.Field<TextItemField>(int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportTesterFieldId").GetSection("Field1").Value.ToString()));
                                keyReportControlActivity.Tester = textProcess.Value;

                                #endregion

                                #endregion

                                listKeyReportRef.Add(keyReportControlActivity);
                            }

                            offset += 500;
                        }

                        if (listKeyReportRef != null && listKeyReportRef.Count() > 0)
                        {
                            using (var context = _soxContext.Database.BeginTransaction())
                            {

                                foreach (var item in listKeyReportRef)
                                {
                                    var itemCheck = _soxContext.KeyReportTester.FirstOrDefault(id => id.PodioItemId == (int)item.PodioItemId);
                                    if (itemCheck != null)
                                    {

                                        itemCheck.Tester = item.Tester;
                                        itemCheck.PodioUniqueId = item.PodioUniqueId;
                                        itemCheck.PodioRevision = item.PodioRevision;
                                        itemCheck.PodioLink = item.PodioLink;

                                        listKeyReportRefUpdate.Add(itemCheck);
                                    }
                                    else
                                    {
                                        listKeyReportRefAdd.Add(item);
                                    }
                                }

                                if (listKeyReportRefAdd != null && listKeyReportRefAdd.Count() > 0)
                                    _soxContext.AddRange(listKeyReportRefAdd);

                                if (listKeyReportRefUpdate != null && listKeyReportRefUpdate.Count() > 0)
                                    _soxContext.UpdateRange(listKeyReportRefUpdate);

                                await _soxContext.SaveChangesAsync();
                                context.Commit();
                            }
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncKeyReportTesterAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncKeyReportTesterAsync");
                return BadRequest(ex.ToString());
            }

            return Ok();

        }


        //[AllowAnonymous]
        [HttpPost("keyreport/KeyReportReviewer")]
        public async Task<IActionResult> SyncKeyReportReviewerAsync([FromBody] SyncDateRange syncDateRange)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("KeyReportReviewerId").Value;
                PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("KeyReportReviewerToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                if (podio.IsAuthenticated())
                {

                    var filter = new Dictionary<string, object>
                    {
                        {
                            "last_edit_on",
                            new  { from = syncDateRange.startDate, to = syncDateRange.endDate}
                        },
                    };

                    PodioCollection<Item> collectionCheck = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 1, 0, filters: filter, sortDesc: true);

                    if (collectionCheck.Filtered != 0)
                    {
                        int loop = (int)Math.Ceiling((double)collectionCheck.Filtered / 500);
                        int offset = 0;
                        List<KeyReportReviewer> listKeyReportRef = new List<KeyReportReviewer>();
                        List<KeyReportReviewer> listKeyReportRefAdd = new List<KeyReportReviewer>();
                        List<KeyReportReviewer> listKeyReportRefUpdate = new List<KeyReportReviewer>();
                        //get filtered items and stored in list
                        for (int i = 0; i < loop; i++)
                        {
                            PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 500, offset, filters: filter, sortDesc: true);

                            foreach (var item in collection.Items)
                            {
                                KeyReportReviewer keyReportControlActivity = new KeyReportReviewer();

                                #region Podio Item
                                keyReportControlActivity.PodioItemId = (int)item.ItemId;
                                keyReportControlActivity.PodioUniqueId = item.AppItemIdFormatted.ToString();
                                keyReportControlActivity.PodioRevision = item.CurrentRevision.Revision;
                                keyReportControlActivity.PodioLink = item.Link.ToString();
                                keyReportControlActivity.CreatedBy = item.CreatedBy.Name.ToString();
                                keyReportControlActivity.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());

                                #region Fields

                                //get process value
                                TextItemField textProcess = item.Field<TextItemField>(int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportReviewerFieldId").GetSection("Field1").Value.ToString()));
                                keyReportControlActivity.Reviewer = textProcess.Value;

                                #endregion

                                #endregion

                                listKeyReportRef.Add(keyReportControlActivity);
                            }

                            offset += 500;
                        }

                        if (listKeyReportRef != null && listKeyReportRef.Count() > 0)
                        {
                            using (var context = _soxContext.Database.BeginTransaction())
                            {

                                foreach (var item in listKeyReportRef)
                                {
                                    var itemCheck = _soxContext.KeyReportReviewer.FirstOrDefault(id => id.PodioItemId == (int)item.PodioItemId);
                                    if (itemCheck != null)
                                    {

                                        itemCheck.Reviewer = item.Reviewer;
                                        itemCheck.PodioUniqueId = item.PodioUniqueId;
                                        itemCheck.PodioRevision = item.PodioRevision;
                                        itemCheck.PodioLink = item.PodioLink;

                                        listKeyReportRefUpdate.Add(itemCheck);
                                    }
                                    else
                                    {
                                        listKeyReportRefAdd.Add(item);
                                    }
                                }

                                if (listKeyReportRefAdd != null && listKeyReportRefAdd.Count() > 0)
                                    _soxContext.AddRange(listKeyReportRefAdd);

                                if (listKeyReportRefUpdate != null && listKeyReportRefUpdate.Count() > 0)
                                    _soxContext.UpdateRange(listKeyReportRefUpdate);

                                await _soxContext.SaveChangesAsync();
                                context.Commit();
                            }
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncKeyReportReviewerAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncKeyReportReviewerAsync");
                return BadRequest(ex.ToString());
            }

            return Ok();

        }


        #endregion



        [HttpPost("podio/sync/questionnaire/cta306")]
        public async Task<IActionResult> SyncCta306Async(string clientName)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("QuestionnaireApp").GetSection("Cta306AppId").Value;
                PodioAppKey.AppToken = _config.GetSection("QuestionnaireApp").GetSection("Cta306AppToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(int.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                if (podio.IsAuthenticated())
                {
                    Debug.WriteLine("Podio Authenticated");
                    int count = 1;
                    var podioResult = await podio.ApplicationService.GetApp(int.Parse(PodioAppKey.AppId));
                    if (podioResult.Fields != null && podioResult.Fields.Count > 0)
                    {

                        if (podioResult.Fields.Count > 0)
                        {
                            //Remove all data first before insert
                            using (var context = _soxContext.Database.BeginTransaction())
                            {

                                _soxContext.RemoveRange(_soxContext.QuestionnaireQuestion.Where(id => id.AppId == PodioAppKey.AppId));
                                _soxContext.RemoveRange(_soxContext.QuestionnaireOption.Where(id => id.AppId == int.Parse(PodioAppKey.AppId)));
                                await _soxContext.SaveChangesAsync();
                                context.Commit();
                            }

                            foreach (var field in podioResult.Fields)
                            {
                                //Console.WriteLine($"{field.Label} : Type ({field.Type})");

                                if (field.Type != "app")
                                {

                                    using (var context = _soxContext.Database.BeginTransaction())
                                    {
                                        QuestionnaireQuestion Question = new QuestionnaireQuestion();
                                        Question.ClientName = clientName;
                                        Question.ControlName = "CTA 3.06";
                                        Question.QuestionString = field.Label;
                                        Question.Type = field.Type;
                                        Question.AppId = PodioAppKey.AppId;
                                        Question.CreatedOn = DateTime.Now;
                                        Question.UpdatedOn = DateTime.Now;
                                        List<QuestionnaireOption> listOptions = new List<QuestionnaireOption>();

                                        if (field.Type == "category")
                                        {
                                            foreach (var option in field.InternalConfig.Settings)
                                            {
                                                if (option.Value.HasValues)
                                                {

                                                    foreach (var item in option.Value)
                                                    {

                                                        QuestionnaireOption QuestionnaireOption = new QuestionnaireOption();
                                                        //Console.WriteLine($"{item["text"]}");
                                                        QuestionnaireOption.OptionName = item["text"].ToString();
                                                        QuestionnaireOption.CreatedOn = DateTime.Now;
                                                        QuestionnaireOption.UpdatedOn = DateTime.Now;
                                                        QuestionnaireOption.AppId = int.Parse(Question.AppId);
                                                        QuestionnaireOption.QuestionnaireQuestion = Question;
                                                        QuestionnaireOption.OptionId = $"{Question.AppId}{count}";
                                                        listOptions.Add(QuestionnaireOption);
                                                        count++;
                                                    }

                                                    Question.Options = listOptions;
                                                }

                                            }
                                        }


                                        _soxContext.Add(Question);
                                        _soxContext.AddRange(listOptions);
                                        await _soxContext.SaveChangesAsync();
                                        context.Commit();
                                    }

                                }


                            }

                        }


                    }

                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncCta306Async");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncCta306Async");
                return BadRequest(ex.ToString());
            }

            return Ok();

        }

        /// <summary>
        /// Sync Questionnaire Podio Fields
        /// </summary>
        /// <param name="questionnaireParam"></param>
        /// <returns>Success (200)</returns>
        [HttpPost("podio/sync/questionnaire/fields")]
        public async Task<IActionResult> SyncQuestionnaireFields([FromBody] QuestionnaireFieldParam questionnaireParam)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = questionnaireParam.AppKey.AppId;
                PodioAppKey.AppToken = questionnaireParam.AppKey.AppToken;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(int.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                if (podio.IsAuthenticated())
                {
                    Debug.WriteLine("Podio Authenticated");
                    int count = 1;
                    int position = 1;
                    var podioResult = await podio.ApplicationService.GetApp(int.Parse(PodioAppKey.AppId));
                    if (podioResult.Fields != null && podioResult.Fields.Count > 0)
                    {

                        if (podioResult.Fields.Count > 0)
                        {
                            //Check if client name and control id exists
                            using (var context = _soxContext.Database.BeginTransaction())
                            {
                                var checkField = _soxContext.QuestionnaireQuestion.Where(x => x.ClientName.Equals(questionnaireParam.ClientName) && x.ControlName.Equals(questionnaireParam.ControlName)).FirstOrDefault();
                                if (checkField != null)
                                {
                                    _soxContext.RemoveRange(_soxContext.QuestionnaireOption.Where(id => id.AppId == int.Parse(checkField.AppId)));
                                    _soxContext.RemoveRange(_soxContext.QuestionnaireQuestion.Where(id => id.AppId == checkField.AppId));
                                    await _soxContext.SaveChangesAsync();
                                    context.Commit();
                                }
                            }

                            //Save app id and token
                            using (var context = _soxContext.Database.BeginTransaction())
                            {
                                PodioAppKey podioApp = new PodioAppKey();
                                var podioKey = _soxContext.PodioAppKey.Where(x => x.AppId.Equals(questionnaireParam.AppKey.AppId)).FirstOrDefault();
                                if (podioKey != null)
                                {
                                    //udpate app token in table PodioAppKey
                                    podioApp = podioKey;
                                    podioKey.AppToken = questionnaireParam.AppKey.AppToken;
                                    _soxContext.Update(podioKey);
                                    await _soxContext.SaveChangesAsync();
                                    context.Commit();

                                }
                                else
                                {
                                    //Create app id and app token in table PodioAppKey
                                    podioApp.AppId = questionnaireParam.AppKey.AppId;
                                    podioApp.AppToken = questionnaireParam.AppKey.AppToken;
                                    _soxContext.Add(podioApp);
                                    await _soxContext.SaveChangesAsync();
                                    context.Commit();

                                }
                            }

                            foreach (var field in podioResult.Fields)
                            {
                                if (field.Status == "active")
                                {
                                    using (var context = _soxContext.Database.BeginTransaction())
                                    {
                                        QuestionnaireQuestion Question = new QuestionnaireQuestion();
                                        Question.ClientName = questionnaireParam.ClientName;
                                        Question.ControlName = questionnaireParam.ControlName;
                                        Question.QuestionString = field.Label;
                                        Question.Type = field.Type;
                                        Question.AppId = PodioAppKey.AppId;
                                        Question.FieldId = field.FieldId;
                                        Question.CreatedOn = DateTime.Now;
                                        Question.UpdatedOn = DateTime.Now;
                                        Question.Position = position;
                                        List<QuestionnaireOption> listOptions = new List<QuestionnaireOption>();

                                        if (field.InternalConfig.Description != string.Empty)
                                        {
                                            Question.Description = field.InternalConfig.Description;
                                        }

                                        if (field.Type == "category")
                                        {
                                            foreach (var option in field.InternalConfig.Settings)
                                            {
                                                if (option.Value.HasValues)
                                                {

                                                    foreach (var item in option.Value)
                                                    {
                                                        if (item["status"].ToString() == "active")
                                                        {
                                                            QuestionnaireOption QuestionnaireOption = new QuestionnaireOption();
                                                            //Console.WriteLine($"{item["text"]}");
                                                            QuestionnaireOption.OptionName = item["text"].ToString();
                                                            QuestionnaireOption.CreatedOn = DateTime.Now;
                                                            QuestionnaireOption.UpdatedOn = DateTime.Now;
                                                            QuestionnaireOption.AppId = int.Parse(Question.AppId);
                                                            QuestionnaireOption.QuestionnaireQuestion = Question;
                                                            QuestionnaireOption.OptionId = $"{Question.AppId}{count}";
                                                            listOptions.Add(QuestionnaireOption);
                                                            count++;
                                                        }
                                                            
                                                    }

                                                    Question.Options = listOptions;
                                                }

                                            }
                                        }

                                        if (field.Type == "date")
                                        {
                                            Debug.WriteLine($"{field.InternalConfig.Settings["end"]}");
                                            if (field.InternalConfig.Settings["end"].ToString() == "enabled")
                                                Question.DtEndRequire = "enabled";
                                            else
                                                Question.DtEndRequire = string.Empty;

                                        }

                                        if (field.Type == "text")
                                        {
                                            if (field.InternalConfig.Settings["size"].ToString() == "large")
                                            {
                                                Question.DtEndRequire = "large";
                                            }
                                        }

                                        if(field.Type == "image")
                                        {
                                            //image field 
                                        }

                                        _soxContext.Add(Question);
                                        _soxContext.AddRange(listOptions);
                                        await _soxContext.SaveChangesAsync();
                                        context.Commit();
                                        position++;
                                    }

                                }

                            }

                        }


                    }

                }

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncQuestionnaireFields");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncQuestionnaireFields");
                return BadRequest(ex.ToString());
            }

            return Ok();

        }


        [HttpPost("iuc/systemgen/fields")]
        public async Task<IActionResult> SyncIUCSystemGenFields([FromBody] QuestionnaireFieldParam questionnaireParam)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                //PodioAppKey.AppId = questionnaireParam.AppId;
                //PodioAppKey.AppToken = questionnaireParam.AppToken;
                PodioAppKey.AppId = questionnaireParam.AppKey.AppId;
                PodioAppKey.AppToken = questionnaireParam.AppKey.AppToken;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(int.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                if (podio.IsAuthenticated())
                {
                    //Console.WriteLine("Podio Authenticated");
                    int position = 1;
                    var podioResult = await podio.ApplicationService.GetApp(int.Parse(PodioAppKey.AppId));
                    if (podioResult.Fields != null && podioResult.Fields.Count > 0)
                    {

                        if (podioResult.Fields.Count > 0)
                        {
                            //Check if exists and remove
                            using (var context = _soxContext.Database.BeginTransaction())
                            {
                                var checkField = _soxContext.IUCSystemGen
                                    .Where(x => x.AppId.Equals(questionnaireParam.AppKey.AppId))
                                    .Include(x => x.ListQuestionAnswer)
                                    .FirstOrDefault();
                                if (checkField != null)
                                {
                                    var listQandA = _soxContext.QandA.Where(x => EF.Property<int>(x, "IUCSystemGenId") == checkField.Id);
                                    if (listQandA.Count() > 0)
                                    {
                                        foreach (var item in listQandA)
                                        {
                                            _soxContext.Remove(item);
                                        }
                                    }
                                    _soxContext.Remove(checkField);
                                    await _soxContext.SaveChangesAsync();
                                    context.Commit();
                                }


                            }

                            //Save app id and token
                            using (var context = _soxContext.Database.BeginTransaction())
                            {
                                PodioAppKey podioApp = new PodioAppKey();
                                var podioKey = _soxContext.PodioAppKey.Where(x => x.AppId.Equals(questionnaireParam.AppKey.AppId)).FirstOrDefault();
                                if (podioKey != null)
                                {
                                    //udpate app token in table PodioAppKey
                                    podioApp = podioKey;
                                    podioKey.AppToken = questionnaireParam.AppKey.AppToken;
                                    await _soxContext.SaveChangesAsync();
                                    context.Commit();
                                    _soxContext.Update(podioKey);
                                }
                                else
                                {
                                    //Create app id and app token in table PodioAppKey
                                    podioApp.AppId = questionnaireParam.AppKey.AppId;
                                    podioApp.AppToken = questionnaireParam.AppKey.AppToken;
                                    await _soxContext.SaveChangesAsync();
                                    context.Commit();
                                    _soxContext.Add(podioApp);
                                }
                            }

                            //Save client control
                            using (var context = _soxContext.Database.BeginTransaction())
                            {
                                ClientControl clientControl = new ClientControl();
                                var checkClientControl = _soxContext.ClientControl.Where(x => x.ClientName.Equals(questionnaireParam.ClientName) && x.ControlName.Equals(questionnaireParam.ControlName)).FirstOrDefault();
                                if (checkClientControl == null)
                                {
                                    clientControl.ClientName = questionnaireParam.ClientName;
                                    clientControl.ControlName = questionnaireParam.ControlName;
                                    await _soxContext.SaveChangesAsync();
                                    context.Commit();
                                    _soxContext.Add(clientControl);
                                }

                            }

                            using (var context = _soxContext.Database.BeginTransaction())
                            {
                                IUCSystemGen IUCSystemGen = new IUCSystemGen();
                                IUCSystemGen.AppId = questionnaireParam.AppKey.AppId;
                                List<QandA> listQandA = new List<QandA>();

                                foreach (var field in podioResult.Fields)
                                {
                                    if (field.Status == "active")
                                    {

                                        QandA qandA = new QandA();
                                        qandA.Question = field.Label;
                                        qandA.Type = field.Type;
                                        qandA.AppId = PodioAppKey.AppId;
                                        qandA.FieldId = field.FieldId;
                                        qandA.Position = position;
                                        qandA.UniqueId = "0";
                                        if (field.Type == "text")
                                        {
                                            if (field.InternalConfig.Settings["size"].ToString() == "large")
                                            {
                                                qandA.Options = "large";
                                            }
                                        }
                                        listQandA.Add(qandA);
                                        position++;
                                    }

                                }

                                if (listQandA.Count > 0)
                                {
                                    IUCSystemGen.ListQuestionAnswer = listQandA;
                                }

                                _soxContext.Add(IUCSystemGen);
                                //_soxContext.AddRange(listQandA);

                                await _soxContext.SaveChangesAsync();
                                context.Commit();

                            }

                        }

                    }

                }

            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncIUCSystemGenFields");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncIUCSystemGenFields");
                return BadRequest(ex.ToString());
            }

            return Ok();

        }

        [HttpPost("iuc/nonsystemgen/fields")]
        public async Task<IActionResult> SyncIUCNonSystemGenFields([FromBody] QuestionnaireFieldParam questionnaireParam)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                //PodioAppKey.AppId = questionnaireParam.AppId;
                //PodioAppKey.AppToken = questionnaireParam.AppToken;
                PodioAppKey.AppId = questionnaireParam.AppKey.AppId;
                PodioAppKey.AppToken = questionnaireParam.AppKey.AppToken;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(int.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                if (podio.IsAuthenticated())
                {
                    //Console.WriteLine("Podio Authenticated");
                    int position = 1;
                    var podioResult = await podio.ApplicationService.GetApp(int.Parse(PodioAppKey.AppId));
                    if (podioResult.Fields != null && podioResult.Fields.Count > 0)
                    {

                        if (podioResult.Fields.Count > 0)
                        {
                            //Check if exists and remove
                            using (var context = _soxContext.Database.BeginTransaction())
                            {
                                var checkField = _soxContext.IUCNonSystemGen
                                    .Where(x => x.AppId.Equals(questionnaireParam.AppKey.AppId))
                                    .Include(x => x.ListQuestionAnswer)
                                    .FirstOrDefault();
                                if (checkField != null)
                                {
                                    var listQandA = _soxContext.QandA.Where(x => EF.Property<int>(x, "IUCNonSystemGenId") == checkField.Id);
                                    if (listQandA.Count() > 0)
                                    {
                                        foreach (var item in listQandA)
                                        {
                                            _soxContext.Remove(item);
                                        }
                                    }
                                    _soxContext.Remove(checkField);
                                    await _soxContext.SaveChangesAsync();
                                    context.Commit();
                                }


                            }

                            //Save app id and token
                            using (var context = _soxContext.Database.BeginTransaction())
                            {
                                PodioAppKey podioApp = new PodioAppKey();
                                var podioKey = _soxContext.PodioAppKey.Where(x => x.AppId.Equals(questionnaireParam.AppKey.AppId)).FirstOrDefault();
                                if (podioKey != null)
                                {
                                    //udpate app token in table PodioAppKey
                                    podioApp = podioKey;
                                    podioKey.AppToken = questionnaireParam.AppKey.AppToken;
                                    _soxContext.Update(podioKey);
                                    await _soxContext.SaveChangesAsync();
                                    context.Commit();

                                }
                                else
                                {
                                    //Create app id and app token in table PodioAppKey
                                    podioApp.AppId = questionnaireParam.AppKey.AppId;
                                    podioApp.AppToken = questionnaireParam.AppKey.AppToken;
                                    _soxContext.Add(podioApp);
                                    await _soxContext.SaveChangesAsync();
                                    context.Commit();

                                }
                            }

                            using (var context = _soxContext.Database.BeginTransaction())
                            {
                                IUCNonSystemGen IUCNonSystemGen = new IUCNonSystemGen();
                                IUCNonSystemGen.AppId = questionnaireParam.AppKey.AppId;
                                List<QandA> listQandA = new List<QandA>();

                                foreach (var field in podioResult.Fields)
                                {
                                    if (field.Status == "active")
                                    {

                                        QandA qandA = new QandA();
                                        qandA.Question = field.Label;
                                        qandA.Type = field.Type;
                                        qandA.AppId = PodioAppKey.AppId;
                                        qandA.FieldId = field.FieldId;
                                        qandA.Position = position;
                                        qandA.UniqueId = "0";
                                        if (field.Type == "text")
                                        {
                                            if (field.InternalConfig.Settings["size"].ToString() == "large")
                                            {
                                                qandA.Options = "large";
                                            }
                                        }
                                        listQandA.Add(qandA);
                                        position++;
                                    }

                                }

                                if (listQandA.Count > 0)
                                {
                                    IUCNonSystemGen.ListQuestionAnswer = listQandA;
                                }

                                _soxContext.Add(IUCNonSystemGen);
                                //_soxContext.AddRange(listQandA);

                                await _soxContext.SaveChangesAsync();
                                context.Commit();

                            }

                        }

                    }

                }

            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncIUCNonSystemGenFields");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncIUCNonSystemGenFields");
                return BadRequest(ex.ToString());
            }

            return Ok();

        }

        private async Task<(bool, int, int)> PullPodioKeyReport(PodioApiKey PodioKey, PodioAppKey PodioAppKey, SyncDateRange syncDateRange)
        {
            //WriteLog writeLog = new WriteLog();
            FormatService textFormat = new FormatService();
            var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            int totalItems = 0;
            int total = 0;
            int count = 0;
            string link = string.Empty;
            await podio.AuthenticateWithApp(int.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

            if (podio.IsAuthenticated())
            {
                Debug.WriteLine("Podio Authenticated");

                var filter = new Dictionary<string, object>
                    {
                        {
                            "last_edit_on",
                            new  { from = syncDateRange.startDate, to = syncDateRange.endDate}
                        },
                    };

                PodioCollection<Item> collectionCheck = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 1, 0, filters: filter, sortDesc: true);
                List<KeyReportUserInput> listKeyReportRef = new List<KeyReportUserInput>();
                List<KeyReportUserInput> listKeyReportRefAdd = new List<KeyReportUserInput>();
                List<KeyReportUserInput> listKeyReportRefUpdate = new List<KeyReportUserInput>();

                if (syncDateRange.limit == 0 && syncDateRange.offset == 0)
                {
                    
                    if (collectionCheck.Filtered != 0)
                    {
                        totalItems = collectionCheck.Filtered;
                        int loop = (int)Math.Ceiling((double)collectionCheck.Filtered / 500);
                        int offset = 0;
                        int itemId = 0;
                        DateTime createdOn;
                        
                        //get filtered items and stored in list
                        for (int i = 0; i < loop; i++)
                        {
                            PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 500, offset, filters: filter, sortDesc: true);

                            foreach (var item in collection.Items)
                            {


                                #region Podio Item
                                itemId = (int)item.ItemId;
                                createdOn = DateTime.Parse(item.CreatedOn.ToString());
                                link = item.Link;
                                count++;
                                #region Fields

                                //Debug.WriteLine($"------------------------------------------------");
                                //Debug.WriteLine($"ITEM");
                                //Debug.WriteLine($"------------------------------------------------");
                                //writeLog.Display(item);
                                //Debug.WriteLine($"------------------------------------------------");

                                foreach (var itemField in item.Fields)
                                {

                                    #region sample return values
                                    //------------------------------------------------
                                    //CATEGORY
                                    //------------------------------------------------
                                    //[{ "value":{ "status":"active","text":"FY20","id":1,"color":"DCEBD8"} }]
                                    //------------------------------------------------

                                    //------------------------------------------------
                                    //APP
                                    //------------------------------------------------
                                    //[{ "value":{ "files":[],"sharefile_vault_url":null,"space":null,"app_item_id":9,"app":{ "status":"active","sharefile_vault_url":null,"name":"Client (SS)","space_id":6923963,"default_view_id":null,"url_add":"https://podio.com/admin-work/sample-selection-024l4l/apps/client-ss/items/new","icon_id":250,"link_add":"https://podio.com/admin-work/sample-selection-024l4l/apps/client-ss/items/new","app_id":24339083,"current_revision":11,"is_default":false,"item_name":"Client","link":"https://podio.com/admin-work/sample-selection-024l4l/apps/client-ss","url":"https://podio.com/admin-work/sample-selection-024l4l/apps/client-ss","url_label":"client-ss","config":{ "item_name":"Client","icon_id":250,"type":"standard","name":"Client (SS)","icon":"250.png"},"item_accounting_info":null,"icon":"250.png"},"title":"Techpoint","initial_revision":{ "item_revision_id":5473928398,"created_via":{ "url":null,"auth_client_id":1,"display":false,"name":"Podio","id":1},"created_by":{ "user_id":3922787,"name":"Levin jay Tagapan","url":"https://podio.com/users/3922787","type":"user","image":{ "hosted_by":"podio","hosted_by_humanized_name":"Podio","thumbnail_link":"https://d2cmuesa4snpwn.cloudfront.net/public/300573459","link":"https://d2cmuesa4snpwn.cloudfront.net/public/300573459","file_id":300573459,"external_file_id":null,"link_target":"_blank"},"avatar_type":"file","avatar":300573459,"id":3922787,"avatar_id":300573459,"last_seen_on":"2021-02-02 19:21:28"},"created_on":"2020-10-13 19:23:25","user":{ "user_id":3922787,"name":"Levin jay Tagapan","url":"https://podio.com/users/3922787","type":"user","image":{ "hosted_by":"podio","hosted_by_humanized_name":"Podio","thumbnail_link":"https://d2cmuesa4snpwn.cloudfront.net/public/300573459","link":"https://d2cmuesa4snpwn.cloudfront.net/public/300573459","file_id":300573459,"external_file_id":null,"link_target":"_blank"},"avatar_type":"file","avatar":300573459,"id":3922787,"avatar_id":300573459,"last_seen_on":"2021-02-02 19:21:28"},"type":"creation","revision":0},"created_via":{ "url":null,"auth_client_id":1,"display":false,"name":"Podio","id":1},"created_by":{ "user_id":3922787,"name":"Levin jay Tagapan","url":"https://podio.com/users/3922787","type":"user","image":{ "hosted_by":"podio","hosted_by_humanized_name":"Podio","thumbnail_link":"https://d2cmuesa4snpwn.cloudfront.net/public/300573459","link":"https://d2cmuesa4snpwn.cloudfront.net/public/300573459","file_id":300573459,"external_file_id":null,"link_target":"_blank"},"avatar_type":"file","avatar":300573459,"id":3922787,"avatar_id":300573459,"last_seen_on":"2021-02-02 19:21:28"},"created_on":"2020-10-13 19:23:25","link":"https://podio.com/admin-work/sample-selection-024l4l/apps/client-ss/items/9","item_id":1544830187,"sharefile_vault_folder_id":null,"revision":7} }]
                                    //------------------------------------------------

                                    //------------------------------------------------
                                    //TEXT
                                    //------------------------------------------------
                                    //[{ "value":"<p>TAX</p>"}]
                                    //------------------------------------------------

                                    //------------------------------------------------
                                    //IMAGE
                                    //------------------------------------------------
                                    //[{ "value":{ "mimetype":"image/jpeg","perma_link":null,"hosted_by":"podio","description":null,"hosted_by_humanized_name":"Podio","size":327680,"thumbnail_link":"https://files.podio.com/1248512808","link":"https://files.podio.com/1248512808","file_id":1248512808,"external_file_id":null,"link_target":"_blank","name":"Q_29fd9957-3034-4377-98b0-e53f0dc3eb91.jpg"} }]
                                    //------------------------------------------------

                                    //------------------------------------------------
                                    //DATE
                                    //------------------------------------------------
                                    //[{ "start_date_utc":"2021-02-02","end":"2021-02-03 00:00:00","end_date":"2021-02-03","end_date_utc":"2021-02-03","start_time_utc":null,"start_time":null,"start_date":"2021-02-02","start":"2021-02-02 00:00:00","end_time":null,"end_time_utc":null,"end_utc":"2021-02-03","start_utc":"2021-02-02"}]
                                    //------------------------------------------------

                                    #endregion

                                    KeyReportUserInput keyReportItem = new KeyReportUserInput();
                                    keyReportItem.ItemId = itemId;
                                    keyReportItem.CreatedOn = createdOn;
                                    keyReportItem.FieldId = itemField.FieldId;
                                    keyReportItem.Link = link;
                                    switch (itemField.Type)
                                    {
                                        case "category":
                                            //category field 
                                            keyReportItem.StrAnswer = itemField.Values[0]["value"]["text"] != null ? itemField.Values[0]["value"]["text"].ToString() : string.Empty;
                                            break;
                                        case "date":
                                            //date field 
                                            keyReportItem.StrAnswer = itemField.Values[0]["start_date"] != null ? itemField.Values[0]["start_date"].ToString() : string.Empty;
                                            keyReportItem.StrAnswer2 = itemField.Values[0]["end_date"] != null ? itemField.Values[0]["end_date"].ToString() : string.Empty;
                                            break;
                                        case "text":
                                            //text field 
                                            //get config for the text field, if its large, we replace <p> to newline
                                            string textSize = itemField.Config.Settings["size"] != null ? itemField.Config.Settings["size"].ToString() : string.Empty;
                                            bool isLarge = textSize == "large" ? true : false;
                                            keyReportItem.StrAnswer = itemField.Values[0]["value"] != null ? textFormat.ReplaceTagHtmlParagraph(itemField.Values[0]["value"].ToString(), isLarge) : string.Empty;
                                            break;
                                        case "image":
                                            //image field 
                                            keyReportItem.StrAnswer = itemField.Values[0]["value"]["name"] != null ? itemField.Values[0]["value"]["name"].ToString() : string.Empty;
                                            break;
                                        case "app":
                                            //app field 
                                            keyReportItem.StrAnswer = itemField.Values[0]["value"]["title"] != null ? itemField.Values[0]["value"]["title"].ToString() : string.Empty;
                                            break;
                                        default:
                                            keyReportItem.StrAnswer = string.Empty;
                                            keyReportItem.StrAnswer2 = string.Empty;
                                            break;
                                    }

                                    listKeyReportRef.Add(keyReportItem);

                                }

                                #endregion

                                #endregion

                            }

                            offset += 500;
                        }

                        total = count >= collectionCheck.Filtered ? collectionCheck.Filtered : count * (syncDateRange.offset + 1);
                        totalItems = collectionCheck.Filtered;
                        await SavePodioKeyReport(listKeyReportRef, PodioAppKey.AppId);

                    }

                    
                }

                else
                {
                    if(collectionCheck.Filtered != 0)
                    {
                        int offset = syncDateRange.offset;
                        int itemId = 0;
                        DateTime createdOn;

                        PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), syncDateRange.limit, offset, filters: filter, sortDesc: true);

                        foreach (var item in collection.Items)
                        {


                            #region Podio Item
                            itemId = (int)item.ItemId;
                            createdOn = DateTime.Parse(item.CreatedOn.ToString());
                            link = item.Link;
                            count++;
                            #region Fields

                            //Debug.WriteLine($"------------------------------------------------");
                            //Debug.WriteLine($"ITEM");
                            //Debug.WriteLine($"------------------------------------------------");
                            //writeLog.Display(item);
                            //Debug.WriteLine($"------------------------------------------------");

                            foreach (var itemField in item.Fields)
                            {

                                KeyReportUserInput keyReportItem = new KeyReportUserInput();
                                keyReportItem.ItemId = itemId;
                                keyReportItem.CreatedOn = createdOn;
                                keyReportItem.FieldId = itemField.FieldId;
                                keyReportItem.StrQuestion = itemField.Label;
                                keyReportItem.Link = link;

                                switch (itemField.Type)
                                {
                                    case "category":
                                        //category field 
                                        keyReportItem.StrAnswer = itemField.Values[0]["value"]["text"] != null ? itemField.Values[0]["value"]["text"].ToString() : string.Empty;
                                        break;
                                    case "date":
                                        //date field 
                                        keyReportItem.StrAnswer = itemField.Values[0]["start_date"] != null ? itemField.Values[0]["start_date"].ToString() : string.Empty;
                                        keyReportItem.StrAnswer2 = itemField.Values[0]["end_date"] != null ? itemField.Values[0]["end_date"].ToString() : string.Empty;
                                        break;
                                    case "text":
                                        //text field 
                                        //get config for the text field, if its large, we replace <p> to newline
                                        string textSize = itemField.Config.Settings["size"] != null ? itemField.Config.Settings["size"].ToString() : string.Empty;
                                        bool isLarge = textSize == "large" ? true : false;
                                        keyReportItem.StrAnswer = itemField.Values[0]["value"] != null ? textFormat.ReplaceTagHtmlParagraph(itemField.Values[0]["value"].ToString(), isLarge) : string.Empty;
                                        break;
                                    case "image":
                                        //image field 
                                        keyReportItem.StrAnswer = itemField.Values[0]["value"]["name"] != null ? itemField.Values[0]["value"]["name"].ToString() : string.Empty;
                                        break;
                                    case "app":
                                        //app field 
                                        keyReportItem.StrAnswer = itemField.Values[0]["value"]["title"] != null ? itemField.Values[0]["value"]["title"].ToString() : string.Empty;
                                        break;
                                    default:
                                        keyReportItem.StrAnswer = string.Empty;
                                        keyReportItem.StrAnswer2 = string.Empty;
                                        break;
                                }

                                listKeyReportRef.Add(keyReportItem);

                            }

                            #endregion

                            #endregion

                        }

                        count = count * (syncDateRange.offset + 1);
                        total = count >= collection.Filtered ? collection.Filtered : count;
                        totalItems = collection.Filtered;

                        await SavePodioKeyReport(listKeyReportRef, PodioAppKey.AppId);
                    }
                        
                    
                }


            }


            return (true, totalItems, total);
        }

        private async Task<bool> SavePodioKeyReport(List<KeyReportUserInput> listKeyReportRef, string appId)
        {
            if (listKeyReportRef != null && listKeyReportRef.Count > 0)
            {
                //Debug.WriteLine($"-------------------------------------");
                //writeLog.Display(listKeyReportRef);
                //Debug.WriteLine($"-------------------------------------");

                //extract itemid distinct
                var listItemId = listKeyReportRef.Select(x => x.ItemId).Distinct().OrderBy(x => x);
                //Debug.WriteLine($"-------------------------------------");
                //writeLog.Display(listItemId);
                //Debug.WriteLine($"-------------------------------------");

                //Get KeyReportQuestion 
                //here we normalize the items base on the question field since return item from podio filter only contains item with value
                //item's that doesn't have a value is not return and we cannot check the image position
                List<KeyReportUserInput> listUserInput = new List<KeyReportUserInput>();
                var keyReportQuestion = _soxContext.KeyReportQuestion.Where(x => x.AppId.Equals(appId)).AsNoTracking();
                if (keyReportQuestion != null)
                {
                    foreach (var id in listItemId)
                    {
                        //var Fy = listKeyReportRef.Where(x => x.ItemId.Equals(id) && x.StrQuestion.ToLower().Contains("1. what is the fy?")).Select(x => x.StrAnswer).FirstOrDefault();
                        //var clientName = listKeyReportRef.Where(x => x.ItemId.Equals(id) && x.StrQuestion.ToLower().Contains("2. client name")).Select(x => x.StrAnswer).FirstOrDefault();
                        //var reportName = listKeyReportRef.Where(x => x.ItemId.Equals(id) && x.StrQuestion.ToLower().Contains("key control id")).Select(x => x.StrAnswer).FirstOrDefault();
                        //var controlId = listKeyReportRef.Where(x => x.ItemId.Equals(id) && x.StrQuestion.ToLower().Contains("name of key report/iuc")).Select(x => x.StrAnswer).FirstOrDefault();

                        foreach (var item in keyReportQuestion)
                        {
                            KeyReportUserInput userInput = new KeyReportUserInput();
                            userInput.AppId = item.AppId;
                            userInput.StrQuestion = item.QuestionString;
                            userInput.Type = item.Type;
                            userInput.Position = item.Position;
                            userInput.FieldId = item.FieldId;
                            userInput.Description = item.Description;
                            userInput.CreatedOn = item.CreatedOn;
                            userInput.Tag = item.Tag;
                            userInput.ItemId = id;
                            userInput.StrAnswer = listKeyReportRef.Where(x => x.ItemId.Equals(id) && x.FieldId.Equals(userInput.FieldId)).Select(x => x.StrAnswer).FirstOrDefault();
                            userInput.StrAnswer2 = listKeyReportRef.Where(x => x.ItemId.Equals(id) && x.FieldId.Equals(userInput.FieldId)).Select(x => x.StrAnswer2).FirstOrDefault();
                            userInput.Link = listKeyReportRef.Where(x => x.ItemId.Equals(id)).Select(x => x.Link).FirstOrDefault();
                            if (userInput.StrQuestion.ToLower().Contains("1. what is the fy"))
                                userInput.TagFY = userInput.StrAnswer;
                            if (userInput.StrQuestion.ToLower().Contains("2. client name"))
                                userInput.TagClientName = userInput.StrAnswer; 
                            if (userInput.StrQuestion.ToLower().Contains("key control using iuc"))
                                userInput.TagControlId = userInput.StrAnswer; 
                            if (userInput.StrQuestion.ToLower().Contains("name of iuc"))
                                userInput.TagReportName = userInput.StrAnswer;
                            if (userInput.StrQuestion.Equals("Status [System]"))
                                userInput.TagStatus = userInput.StrAnswer;
                            listUserInput.Add(userInput);
                        }

                    }
                }

                //Debug.WriteLine($"-------------------------------------");
                //writeLog.Display(listUserInput.OrderBy(x => x.Position));
                //Debug.WriteLine($"-------------------------------------");

                //update tag fields
                foreach (var item in listUserInput.OrderBy(x => x.Position))
                {
                    //if (item.TagClientName == null && item.TagClientName == string.Empty)
                    item.TagClientName = listUserInput.Where(x => x.ItemId.Equals(item.ItemId) && x.TagClientName != null).Select(x => x.TagClientName).FirstOrDefault();
                    //if (item.TagFY == null && item.TagFY == string.Empty)
                    item.TagFY = listUserInput.Where(x => x.ItemId.Equals(item.ItemId) && x.TagFY != null).Select(x => x.TagFY).FirstOrDefault();
                    //if (item.TagReportName == null && item.TagReportName == string.Empty)
                    item.TagReportName = listUserInput.Where(x => x.ItemId.Equals(item.ItemId) && x.TagReportName != null).Select(x => x.TagReportName).FirstOrDefault();
                    //if (item.TagControlId == null && item.TagControlId == string.Empty)
                    item.TagControlId = listUserInput.Where(x => x.ItemId.Equals(item.ItemId) && x.TagControlId != null).Select(x => x.TagControlId).FirstOrDefault();
                    item.TagStatus = listUserInput.Where(x => x.ItemId.Equals(item.ItemId) && x.TagStatus != null).Select(x => x.TagStatus).FirstOrDefault();

                    var checkExists = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(item.ItemId) && x.FieldId.Equals(item.FieldId)).AsNoTracking().FirstOrDefault();
                    if (checkExists != null)
                    {
                        //Updatelk,
                        using (var context = _soxContext.Database.BeginTransaction())
                        {

                            item.Id = checkExists.Id;
                            _soxContext.Update(item);
                            //_soxContext.Entry(checkExists).CurrentValues.SetValues(item);
                            await _soxContext.SaveChangesAsync();
                            context.Commit();
                            Debug.WriteLine($"Updated Item Id : {item.ItemId}");
                            //writeLog.Display(item);
                            //Debug.WriteLine($"------------------------------------------------------");
                        }
                    }
                    else
                    {
                        //Save
                        using (var context = _soxContext.Database.BeginTransaction())
                        {
                            await _soxContext.AddAsync(item);
                            await _soxContext.SaveChangesAsync();
                            context.Commit();
                            Debug.WriteLine($"Created Item Id : {item.ItemId}");
                        }
                    }
                }
                #region 
                //WriteLog writeLog = new WriteLog();
                //foreach (var item in listUserInput.OrderBy(x => x.Position))
                //{
                //    var checkExists = _soxContext.KeyReportUserInput.Where(x => x.ItemId.Equals(item.ItemId) && x.FieldId.Equals(item.FieldId)).AsNoTracking().FirstOrDefault();
                //    if (checkExists != null)
                //    {
                //        //Updatelk,
                //        using (var context = _soxContext.Database.BeginTransaction())
                //        {

                //            item.Id = checkExists.Id;
                //            _soxContext.Update(item);
                //            //_soxContext.Entry(checkExists).CurrentValues.SetValues(item);
                //            await _soxContext.SaveChangesAsync();
                //            context.Commit();
                //            Debug.WriteLine($"Updated Item Id : {item.ItemId}");
                //            //writeLog.Display(item);
                //            //Debug.WriteLine($"------------------------------------------------------");
                //        }
                //    }
                //    else
                //    {
                //        //Save
                //        using (var context = _soxContext.Database.BeginTransaction())
                //        {
                //            await _soxContext.AddAsync(item);
                //            await _soxContext.SaveChangesAsync();
                //            context.Commit();
                //            Debug.WriteLine($"Created Item Id : {item.ItemId}");
                //        }
                //    }
                //}
                #endregion



            }

            return true;
        }

        [HttpPost("keyreport/Parameter")]
        public async Task<IActionResult> SyncParameterAsync([FromBody] SyncDateRange syncDateRange)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("KeyReportParametersId").Value;
                PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("KeyReportParametersToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);
                int totalItems = 0;
                int total = 0;
                int count = 0;
                if (podio.IsAuthenticated())
                {

                    var filter = new Dictionary<string, object>
                    {
                        {
                            "last_edit_on",
                            new  { from = syncDateRange.startDate, to = syncDateRange.endDate}
                        },
                    };

                    int clientNameField = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportParametersField").GetSection("ClientName").Value);
                    int keyReportNameField = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportParametersField").GetSection("KeyReportName").Value);
                    int methodField = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportParametersField").GetSection("Method").Value);
                    int parameterField = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportParametersField").GetSection("Parameter").Value);
                    int a1Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportParametersField").GetSection("A1").Value);
                    int a2Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportParametersField").GetSection("A2").Value);
                    int a3Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportParametersField").GetSection("A3").Value);
                    int a4Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportParametersField").GetSection("A4").Value);
                    int a5Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportParametersField").GetSection("A5").Value);
                    int a6Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportParametersField").GetSection("A6").Value);
                    int a7Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportParametersField").GetSection("A7").Value);
                    int a8Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportParametersField").GetSection("A8").Value);
                    int a9Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportParametersField").GetSection("A9").Value);
                    int a10Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportParametersField").GetSection("A10").Value);

                    int statusField = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportParametersField").GetSection("Status").Value);

                    PodioCollection<Item> collectionCheck = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 1, 0, filters: filter, sortDesc: true);

                    List<ParametersLibrary> listParameter = new List<ParametersLibrary>();
                    List<ParametersLibrary> listParameterAdd = new List<ParametersLibrary>();
                    List<ParametersLibrary> listParameterUpdate = new List<ParametersLibrary>();

                    if (syncDateRange.limit == 0 && syncDateRange.offset == 0)
                    {
                        if (collectionCheck.Filtered != 0)
                        {
                            int loop = (int)Math.Ceiling((double)collectionCheck.Filtered / 500);
                            int offset = 0;

                            //get filtered items and stored in list
                            for (int i = 0; i < loop; i++)
                            {
                                PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 500, offset, filters: filter, sortDesc: true);

                                foreach (var item in collection.Items)
                                {
                                    ParametersLibrary parametersLibrary = new ParametersLibrary();
                                    count++;

                                    parametersLibrary.PodioItemId = (int)item.ItemId;
                                    parametersLibrary.PodioUniqueId = item.AppItemIdFormatted.ToString();
                                    parametersLibrary.PodioRevision = item.CurrentRevision.Revision;
                                    parametersLibrary.PodioLink = item.Link.ToString();
                                    parametersLibrary.CreatedBy = item.CreatedBy.Name.ToString();
                                    parametersLibrary.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());


                                    //Client Reference
                                    AppItemField clientApp = item.Field<AppItemField>(clientNameField);
                                    IEnumerable<Item> clientAppRef = clientApp.Items;
                                    parametersLibrary.ClientName = clientAppRef.Select(x => x.Title).FirstOrDefault();

                                    //Key Report Name
                                    TextItemField txtKeyReportName = item.Field<TextItemField>(keyReportNameField);
                                    parametersLibrary.KeyReportName = txtKeyReportName.Value;

                                    //Method
                                    AppItemField catMethod = item.Field<AppItemField>(methodField);
                                    IEnumerable<Item> methodAppRef = catMethod.Items;
                                    parametersLibrary.Method = methodAppRef.Select(x => x.Title).FirstOrDefault();

                                    //Parameter
                                    TextItemField txtCompleteness = item.Field<TextItemField>(parameterField);
                                    parametersLibrary.Parameter = txtCompleteness.Value;

                                    //A1
                                    TextItemField txtA1 = item.Field<TextItemField>(a1Field);
                                    parametersLibrary.A1 = txtA1.Value;

                                    //A2
                                    TextItemField txtA2 = item.Field<TextItemField>(a2Field);
                                    parametersLibrary.A2 = txtA2.Value;

                                    //A3
                                    TextItemField txtA3 = item.Field<TextItemField>(a3Field);
                                    parametersLibrary.A3 = txtA3.Value;

                                    //A4
                                    TextItemField txtA4 = item.Field<TextItemField>(a4Field);
                                    parametersLibrary.A4 = txtA4.Value;

                                    //A5
                                    TextItemField txtA5 = item.Field<TextItemField>(a5Field);
                                    parametersLibrary.A5 = txtA5.Value;

                                    //A6
                                    TextItemField txtA6 = item.Field<TextItemField>(a6Field);
                                    parametersLibrary.A6 = txtA6.Value;

                                    //A7
                                    TextItemField txtA7 = item.Field<TextItemField>(a7Field);
                                    parametersLibrary.A7 = txtA7.Value;

                                    //A8
                                    TextItemField txtA8 = item.Field<TextItemField>(a8Field);
                                    parametersLibrary.A8 = txtA8.Value;

                                    //A9
                                    TextItemField txtA9 = item.Field<TextItemField>(a9Field);
                                    parametersLibrary.A9 = txtA9.Value;

                                    //A10
                                    TextItemField txtA10 = item.Field<TextItemField>(a10Field);
                                    parametersLibrary.A10 = txtA10.Value;

                                    //Status
                                    CategoryItemField catStatus = item.Field<CategoryItemField>(statusField);
                                    IEnumerable<CategoryItemField.Option> listCatStatus = catStatus.Options;
                                    parametersLibrary.status = listCatStatus.Select(x => x.Text).FirstOrDefault();

                                    listParameter.Add(parametersLibrary);
                                }

                                offset += 500;
                            }

                            total = count >= collectionCheck.Filtered ? collectionCheck.Filtered : count * (syncDateRange.offset + 1);
                            totalItems = collectionCheck.Filtered;


                        }
                    }
                    else
                    {
                        if (collectionCheck.Filtered != 0)
                        {
                            int offset = syncDateRange.offset;
                            PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), syncDateRange.limit, offset, filters: filter, sortDesc: true);

                            foreach (var item in collection.Items)
                            {
                                ParametersLibrary parametersLibrary = new ParametersLibrary();
                                count++;

                                parametersLibrary.PodioItemId = (int)item.ItemId;
                                parametersLibrary.PodioUniqueId = item.AppItemIdFormatted.ToString();
                                parametersLibrary.PodioRevision = item.CurrentRevision.Revision;
                                parametersLibrary.PodioLink = item.Link.ToString();
                                parametersLibrary.CreatedBy = item.CreatedBy.Name.ToString();
                                parametersLibrary.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());


                                //Client Reference
                                AppItemField clientApp = item.Field<AppItemField>(clientNameField);
                                IEnumerable<Item> clientAppRef = clientApp.Items;
                                parametersLibrary.ClientName = clientAppRef.Select(x => x.Title).FirstOrDefault();

                                //Key Report Name
                                TextItemField txtKeyReportName = item.Field<TextItemField>(keyReportNameField);
                                parametersLibrary.KeyReportName = txtKeyReportName.Value;

                                //Method
                                AppItemField catMethod = item.Field<AppItemField>(methodField);
                                IEnumerable<Item> methodAppRef = catMethod.Items;
                                parametersLibrary.Method = methodAppRef.Select(x => x.Title).FirstOrDefault();

                                //Parameter
                                TextItemField txtCompleteness = item.Field<TextItemField>(parameterField);
                                parametersLibrary.Parameter = txtCompleteness.Value;

                                //A1
                                TextItemField txtA1 = item.Field<TextItemField>(a1Field);
                                parametersLibrary.A1 = txtA1.Value;

                                //A2
                                TextItemField txtA2 = item.Field<TextItemField>(a2Field);
                                parametersLibrary.A2 = txtA2.Value;

                                //A3
                                TextItemField txtA3 = item.Field<TextItemField>(a3Field);
                                parametersLibrary.A3 = txtA3.Value;

                                //A4
                                TextItemField txtA4 = item.Field<TextItemField>(a4Field);
                                parametersLibrary.A4 = txtA4.Value;

                                //A5
                                TextItemField txtA5 = item.Field<TextItemField>(a5Field);
                                parametersLibrary.A5 = txtA5.Value;

                                //A6
                                TextItemField txtA6 = item.Field<TextItemField>(a6Field);
                                parametersLibrary.A6 = txtA6.Value;

                                //A7
                                TextItemField txtA7 = item.Field<TextItemField>(a7Field);
                                parametersLibrary.A7 = txtA7.Value;

                                //A8
                                TextItemField txtA8 = item.Field<TextItemField>(a8Field);
                                parametersLibrary.A8 = txtA8.Value;

                                //A9
                                TextItemField txtA9 = item.Field<TextItemField>(a9Field);
                                parametersLibrary.A9 = txtA9.Value;

                                //A10
                                TextItemField txtA10 = item.Field<TextItemField>(a10Field);
                                parametersLibrary.A10 = txtA10.Value;

                                //Status
                                CategoryItemField catStatus = item.Field<CategoryItemField>(statusField);
                                IEnumerable<CategoryItemField.Option> listCatStatus = catStatus.Options;
                                parametersLibrary.status = listCatStatus.Select(x => x.Text).FirstOrDefault();

                                listParameter.Add(parametersLibrary);
                            }

                            count = count * (syncDateRange.offset + 1);
                            total = count >= collection.Filtered ? collection.Filtered : count;
                            totalItems = collection.Filtered;
                        }
                    }

                    if (listParameter != null && listParameter.Count() > 0)
                    {
                        using (var context = _soxContext.Database.BeginTransaction())
                        {

                            foreach (var item in listParameter)
                            {
                                var parameterCheck = _soxContext.ParametersLibrary.FirstOrDefault(id => id.PodioItemId == (int)item.PodioItemId);
                                if (parameterCheck != null)
                                {
                                    parameterCheck.ClientName = item.ClientName;
                                    parameterCheck.KeyReportName = item.KeyReportName;
                                    parameterCheck.Method = item.Method;
                                    parameterCheck.Parameter = item.Parameter;
                                    parameterCheck.A1 = item.A1;
                                    parameterCheck.A2 = item.A2;
                                    parameterCheck.A3 = item.A3;
                                    parameterCheck.A4 = item.A4;
                                    parameterCheck.A5 = item.A5;
                                    parameterCheck.A6 = item.A6;
                                    parameterCheck.A7 = item.A7;
                                    parameterCheck.A8 = item.A8;
                                    parameterCheck.A9 = item.A9;
                                    parameterCheck.A10 = item.A10;
                                    parameterCheck.PodioUniqueId = item.PodioUniqueId;
                                    parameterCheck.PodioRevision = item.PodioRevision;
                                    parameterCheck.PodioLink = item.PodioLink;
                                    parameterCheck.status = item.status;

                                    listParameterUpdate.Add(parameterCheck);
                                }
                                else
                                {
                                    listParameterAdd.Add(item);
                                }
                            }

                            if (listParameterAdd != null && listParameterAdd.Count() > 0)
                                _soxContext.AddRange(listParameterAdd);

                            if (listParameterUpdate != null && listParameterUpdate.Count() > 0)
                                _soxContext.UpdateRange(listParameterUpdate);

                            await _soxContext.SaveChangesAsync();
                            context.Commit();
                        }
                    }

                }
                string saveItems = $"{total}/{totalItems}";
                return Ok(new { syncDateRange.limit, totalItems, saveItems });

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncParametersAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncParametersAsync");
                return BadRequest(ex.ToString());
            }

            return Ok();

        }

        [HttpPost("keyreport/Report")]
        public async Task<IActionResult> SyncReportAsync([FromBody] SyncDateRange syncDateRange)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("KeyReportReportsId").Value;
                PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("KeyReportReportsToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);
                int totalItems = 0;
                int total = 0;
                int count = 0;
                if (podio.IsAuthenticated())
                {

                    var filter = new Dictionary<string, object>
                    {
                        {
                            "last_edit_on",
                            new  { from = syncDateRange.startDate, to = syncDateRange.endDate}
                        },
                    };

                    int clientNameField = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportReportsField").GetSection("ClientName").Value);
                    int keyReportNameField = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportReportsField").GetSection("KeyReportName").Value);
                    int methodField = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportReportsField").GetSection("Method").Value);
                    int reportField = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportReportsField").GetSection("Report").Value);
                    int b1Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportReportsField").GetSection("B1").Value);
                    int b2Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportReportsField").GetSection("B2").Value);
                    int b3Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportReportsField").GetSection("B3").Value);
                    int b4Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportReportsField").GetSection("B4").Value);
                    int b5Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportReportsField").GetSection("B5").Value);
                    int b6Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportReportsField").GetSection("B6").Value);
                    int b7Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportReportsField").GetSection("B7").Value);
                    int b8Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportReportsField").GetSection("B8").Value);
                    int b9Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportReportsField").GetSection("B9").Value);
                    int b10Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportReportsField").GetSection("B10").Value);
                    int statusField = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportReportsField").GetSection("Status").Value);

                    PodioCollection<Item> collectionCheck = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 1, 0, filters: filter, sortDesc: true);

                    List<ReportsLibrary> listReports = new List<ReportsLibrary>();
                    List<ReportsLibrary> listReportsAdd = new List<ReportsLibrary>();
                    List<ReportsLibrary> listReportsUpdate = new List<ReportsLibrary>();

                    if (syncDateRange.limit == 0 && syncDateRange.offset == 0)
                    {
                        if (collectionCheck.Filtered != 0)
                        {
                            int loop = (int)Math.Ceiling((double)collectionCheck.Filtered / 500);
                            int offset = 0;

                            //get filtered items and stored in list
                            for (int i = 0; i < loop; i++)
                            {
                                PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 500, offset, filters: filter, sortDesc: true);

                                foreach (var item in collection.Items)
                                {
                                    ReportsLibrary reportsLibrary = new ReportsLibrary();
                                    count++;

                                    reportsLibrary.PodioItemId = (int)item.ItemId;
                                    reportsLibrary.PodioUniqueId = item.AppItemIdFormatted.ToString();
                                    reportsLibrary.PodioRevision = item.CurrentRevision.Revision;
                                    reportsLibrary.PodioLink = item.Link.ToString();
                                    reportsLibrary.CreatedBy = item.CreatedBy.Name.ToString();
                                    reportsLibrary.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());






                                    //Client Reference
                                    AppItemField clientApp = item.Field<AppItemField>(clientNameField);
                                    IEnumerable<Item> clientAppRef = clientApp.Items;
                                    reportsLibrary.ClientName = clientAppRef.Select(x => x.Title).FirstOrDefault();

                                    //Key Report Name
                                    TextItemField txtKeyReportName = item.Field<TextItemField>(keyReportNameField);
                                    reportsLibrary.KeyReportName = txtKeyReportName.Value;

                                    //Method
                                    AppItemField catMethod = item.Field<AppItemField>(methodField);
                                    IEnumerable<Item> methodAppRef = catMethod.Items;
                                    reportsLibrary.Method = methodAppRef.Select(x => x.Title).FirstOrDefault();

                                    //Report
                                    TextItemField txtReport = item.Field<TextItemField>(reportField);
                                    reportsLibrary.Report = txtReport.Value;

                                    //B1
                                    TextItemField txtB1 = item.Field<TextItemField>(b1Field);
                                    reportsLibrary.B1 = txtB1.Value;

                                    //B2
                                    TextItemField txtB2 = item.Field<TextItemField>(b2Field);
                                    reportsLibrary.B2 = txtB2.Value;

                                    //B3
                                    TextItemField txtB3 = item.Field<TextItemField>(b3Field);
                                    reportsLibrary.B3 = txtB3.Value;

                                    //B4
                                    TextItemField txtB4 = item.Field<TextItemField>(b4Field);
                                    reportsLibrary.B4 = txtB4.Value;

                                    //B5
                                    TextItemField txtB5 = item.Field<TextItemField>(b5Field);
                                    reportsLibrary.B5 = txtB5.Value;

                                    //B6
                                    TextItemField txtB6 = item.Field<TextItemField>(b6Field);
                                    reportsLibrary.B6 = txtB6.Value;

                                    //B7
                                    TextItemField txtB7 = item.Field<TextItemField>(b7Field);
                                    reportsLibrary.B7 = txtB7.Value;

                                    //B8
                                    TextItemField txtB8 = item.Field<TextItemField>(b8Field);
                                    reportsLibrary.B8 = txtB8.Value;

                                    //B9
                                    TextItemField txtB9 = item.Field<TextItemField>(b9Field);
                                    reportsLibrary.B9 = txtB9.Value;

                                    //B10
                                    TextItemField txtB10 = item.Field<TextItemField>(b10Field);
                                    reportsLibrary.B10 = txtB10.Value;

                                    //Status
                                    CategoryItemField catStatus = item.Field<CategoryItemField>(statusField);
                                    IEnumerable<CategoryItemField.Option> listCatStatus = catStatus.Options;
                                    reportsLibrary.status = listCatStatus.Select(x => x.Text).FirstOrDefault();

                                    listReports.Add(reportsLibrary);
                                }

                                offset += 500;
                            }

                            total = count >= collectionCheck.Filtered ? collectionCheck.Filtered : count * (syncDateRange.offset + 1);
                            totalItems = collectionCheck.Filtered;


                        }
                    }
                    else
                    {
                        if (collectionCheck.Filtered != 0)
                        {
                            int offset = syncDateRange.offset;
                            PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), syncDateRange.limit, offset, filters: filter, sortDesc: true);

                            foreach (var item in collection.Items)
                            {
                                ReportsLibrary reportsLibrary = new ReportsLibrary();
                                count++;

                                reportsLibrary.PodioItemId = (int)item.ItemId;
                                reportsLibrary.PodioUniqueId = item.AppItemIdFormatted.ToString();
                                reportsLibrary.PodioRevision = item.CurrentRevision.Revision;
                                reportsLibrary.PodioLink = item.Link.ToString();
                                reportsLibrary.CreatedBy = item.CreatedBy.Name.ToString();
                                reportsLibrary.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());






                                //Client Reference
                                AppItemField clientApp = item.Field<AppItemField>(clientNameField);
                                IEnumerable<Item> clientAppRef = clientApp.Items;
                                reportsLibrary.ClientName = clientAppRef.Select(x => x.Title).FirstOrDefault();

                                //Key Report Name
                                TextItemField txtKeyReportName = item.Field<TextItemField>(keyReportNameField);
                                reportsLibrary.KeyReportName = txtKeyReportName.Value;

                                //Method
                                AppItemField catMethod = item.Field<AppItemField>(methodField);
                                IEnumerable<Item> methodAppRef = catMethod.Items;
                                reportsLibrary.Method = methodAppRef.Select(x => x.Title).FirstOrDefault();

                                //Report
                                TextItemField txtReport = item.Field<TextItemField>(reportField);
                                reportsLibrary.Report = txtReport.Value;

                                //B1
                                TextItemField txtB1 = item.Field<TextItemField>(b1Field);
                                reportsLibrary.B1 = txtB1.Value;

                                //B2
                                TextItemField txtB2 = item.Field<TextItemField>(b2Field);
                                reportsLibrary.B2 = txtB2.Value;

                                //B3
                                TextItemField txtB3 = item.Field<TextItemField>(b3Field);
                                reportsLibrary.B3 = txtB3.Value;

                                //B4
                                TextItemField txtB4 = item.Field<TextItemField>(b4Field);
                                reportsLibrary.B4 = txtB4.Value;

                                //B5
                                TextItemField txtB5 = item.Field<TextItemField>(b5Field);
                                reportsLibrary.B5 = txtB5.Value;

                                //B6
                                TextItemField txtB6 = item.Field<TextItemField>(b6Field);
                                reportsLibrary.B6 = txtB6.Value;

                                //B7
                                TextItemField txtB7 = item.Field<TextItemField>(b7Field);
                                reportsLibrary.B7 = txtB7.Value;

                                //B8
                                TextItemField txtB8 = item.Field<TextItemField>(b8Field);
                                reportsLibrary.B8 = txtB8.Value;

                                //B9
                                TextItemField txtB9 = item.Field<TextItemField>(b9Field);
                                reportsLibrary.B9 = txtB9.Value;

                                //B10
                                TextItemField txtB10 = item.Field<TextItemField>(b10Field);
                                reportsLibrary.B10 = txtB10.Value;
                                //Status
                                CategoryItemField catStatus = item.Field<CategoryItemField>(statusField);
                                IEnumerable<CategoryItemField.Option> listCatStatus = catStatus.Options;
                                reportsLibrary.status = listCatStatus.Select(x => x.Text).FirstOrDefault();

                                listReports.Add(reportsLibrary);
                            }

                            count = count * (syncDateRange.offset + 1);
                            total = count >= collection.Filtered ? collection.Filtered : count;
                            totalItems = collection.Filtered;
                        }
                    }

                    if (listReports != null && listReports.Count() > 0)
                    {
                        using (var context = _soxContext.Database.BeginTransaction())
                        {

                            foreach (var item in listReports)
                            {
                                var reportsCheck = _soxContext.ReportsLibrary.FirstOrDefault(id => id.PodioItemId == (int)item.PodioItemId);
                                if (reportsCheck != null)
                                {
                                    reportsCheck.ClientName = item.ClientName;
                                    reportsCheck.KeyReportName = item.KeyReportName;
                                    reportsCheck.Method = item.Method;
                                    reportsCheck.Report = item.Report;
                                    reportsCheck.B1 = item.B1;
                                    reportsCheck.B2 = item.B2;
                                    reportsCheck.B3 = item.B3;
                                    reportsCheck.B4 = item.B4;
                                    reportsCheck.B5 = item.B5;
                                    reportsCheck.B6 = item.B6;
                                    reportsCheck.B7 = item.B7;
                                    reportsCheck.B8 = item.B8;
                                    reportsCheck.B9 = item.B9;
                                    reportsCheck.B10 = item.B10;
                                    reportsCheck.PodioUniqueId = item.PodioUniqueId;
                                    reportsCheck.PodioRevision = item.PodioRevision;
                                    reportsCheck.PodioLink = item.PodioLink;
                                    reportsCheck.status = item.status;

                                    listReportsUpdate.Add(reportsCheck);
                                }
                                else
                                {
                                    listReportsAdd.Add(item);
                                }
                            }

                            if (listReportsAdd != null && listReportsAdd.Count() > 0)
                                _soxContext.AddRange(listReportsAdd);

                            if (listReportsUpdate != null && listReportsUpdate.Count() > 0)
                                _soxContext.UpdateRange(listReportsUpdate);

                            await _soxContext.SaveChangesAsync();
                            context.Commit();
                        }
                    }

                }
                string saveItems = $"{total}/{totalItems}";
                return Ok(new { syncDateRange.limit, totalItems, saveItems });

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncReportAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncReportsAsync");
                return BadRequest(ex.ToString());
            }

            return Ok();

        }
        //[AllowAnonymous]
        [HttpPost("keyreport/Completeness")]
        public async Task<IActionResult> SyncCompletenessAsync([FromBody] SyncDateRange syncDateRange)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("KeyReportCompletenessId").Value;
                PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("KeyReportCompletenessToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);
                int totalItems = 0;
                int total = 0;
                int count = 0;
                if (podio.IsAuthenticated())
                {

                    var filter = new Dictionary<string, object>
                    {
                        {
                            "last_edit_on",
                            new  { from = syncDateRange.startDate, to = syncDateRange.endDate}
                        },
                    };

                    int clientNameField = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportCompletenessField").GetSection("ClientName").Value);
                    int keyReportNameField = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportCompletenessField").GetSection("KeyReportName").Value);
                    int methodField = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportCompletenessField").GetSection("Method").Value);
                    int completenessField = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportCompletenessField").GetSection("Completeness").Value);
                    int c1Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportCompletenessField").GetSection("C1").Value);
                    int c2Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportCompletenessField").GetSection("C2").Value);
                    int c3Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportCompletenessField").GetSection("C3").Value);
                    int c4Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportCompletenessField").GetSection("C4").Value);
                    int c5Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportCompletenessField").GetSection("C5").Value);
                    int c6Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportCompletenessField").GetSection("C6").Value);
                    int c7Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportCompletenessField").GetSection("C7").Value);
                    int c8Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportCompletenessField").GetSection("C8").Value);
                    int c9Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportCompletenessField").GetSection("C9").Value);
                    int c10Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportCompletenessField").GetSection("C10").Value);
                    int statusField = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportCompletenessField").GetSection("Status").Value);

                    PodioCollection<Item> collectionCheck = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 1, 0, filters: filter, sortDesc: true);

                    List<CompletenessLibrary> listCompleteness = new List<CompletenessLibrary>();
                    List<CompletenessLibrary> listCompletenessAdd = new List<CompletenessLibrary>();
                    List<CompletenessLibrary> listCompletenessUpdate = new List<CompletenessLibrary>();

                    if (syncDateRange.limit == 0 && syncDateRange.offset == 0)
                    {
                        if (collectionCheck.Filtered != 0)
                        {
                            int loop = (int)Math.Ceiling((double)collectionCheck.Filtered / 500);
                            int offset = 0;

                            //get filtered items and stored in list
                            for (int i = 0; i < loop; i++)
                            {
                                PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 500, offset, filters: filter, sortDesc: true);

                                foreach (var item in collection.Items)
                                {
                                    CompletenessLibrary completenessLibrary = new CompletenessLibrary();
                                    count++;
                                    
                                    completenessLibrary.PodioItemId = (int)item.ItemId;
                                    completenessLibrary.PodioUniqueId = item.AppItemIdFormatted.ToString();
                                    completenessLibrary.PodioRevision = item.CurrentRevision.Revision;
                                    completenessLibrary.PodioLink = item.Link.ToString();
                                    completenessLibrary.CreatedBy = item.CreatedBy.Name.ToString();
                                    completenessLibrary.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());


                                    

                                    
                                    
                                    //Client Reference
                                    AppItemField clientApp = item.Field<AppItemField>(clientNameField);
                                    IEnumerable<Item> clientAppRef = clientApp.Items;
                                    completenessLibrary.ClientName = clientAppRef.Select(x => x.Title).FirstOrDefault();

                                    //Key Report Name
                                    TextItemField txtKeyReportName = item.Field<TextItemField>(keyReportNameField);
                                    completenessLibrary.KeyReportName = txtKeyReportName.Value;

                                    //Method
                                    AppItemField catMethod = item.Field<AppItemField>(methodField);
                                    IEnumerable<Item> methodAppRef = catMethod.Items;
                                    completenessLibrary.Method = methodAppRef.Select(x => x.Title).FirstOrDefault();
                                    //Completeness
                                    TextItemField txtCompleteness = item.Field<TextItemField>(completenessField);
                                    completenessLibrary.Completeness = txtCompleteness.Value;

                                    //C1
                                    TextItemField txtC1 = item.Field<TextItemField>(c1Field);
                                    completenessLibrary.C1 = txtC1.Value;

                                    //C2
                                    TextItemField txtC2 = item.Field<TextItemField>(c2Field);
                                    completenessLibrary.C2 = txtC2.Value;

                                    //C3
                                    TextItemField txtC3 = item.Field<TextItemField>(c3Field);
                                    completenessLibrary.C3 = txtC3.Value;

                                    //C4
                                    TextItemField txtC4 = item.Field<TextItemField>(c4Field);
                                    completenessLibrary.C4 = txtC4.Value;

                                    //C5
                                    TextItemField txtC5 = item.Field<TextItemField>(c5Field);
                                    completenessLibrary.C5 = txtC5.Value;

                                    //C6
                                    TextItemField txtC6 = item.Field<TextItemField>(c6Field);
                                    completenessLibrary.C6 = txtC6.Value;

                                    //C7
                                    TextItemField txtC7 = item.Field<TextItemField>(c7Field);
                                    completenessLibrary.C7 = txtC7.Value;

                                    //C8
                                    TextItemField txtC8 = item.Field<TextItemField>(c8Field);
                                    completenessLibrary.C8 = txtC8.Value;

                                    //C9
                                    TextItemField txtC9 = item.Field<TextItemField>(c9Field);
                                    completenessLibrary.C9 = txtC9.Value;

                                    //C10
                                    TextItemField txtC10 = item.Field<TextItemField>(c10Field);
                                    completenessLibrary.C10 = txtC10.Value;

                                    //Status
                                    CategoryItemField catStatus = item.Field<CategoryItemField>(statusField);
                                    IEnumerable<CategoryItemField.Option> listCatStatus = catStatus.Options;
                                    completenessLibrary.status = listCatStatus.Select(x => x.Text).FirstOrDefault();

                                    listCompleteness.Add(completenessLibrary);
                                }

                                offset += 500;
                            }

                            total = count >= collectionCheck.Filtered ? collectionCheck.Filtered : count * (syncDateRange.offset + 1);
                            totalItems = collectionCheck.Filtered;


                        }
                    }
                    else
                    {
                        if (collectionCheck.Filtered != 0)
                        {
                            int offset = syncDateRange.offset;
                            PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), syncDateRange.limit, offset, filters: filter, sortDesc: true);

                            foreach (var item in collection.Items)
                            {
                                CompletenessLibrary completenessLibrary = new CompletenessLibrary();
                                count++;
                               
                                completenessLibrary.PodioItemId = (int)item.ItemId;
                                completenessLibrary.PodioUniqueId = item.AppItemIdFormatted.ToString();
                                completenessLibrary.PodioRevision = item.CurrentRevision.Revision;
                                completenessLibrary.PodioLink = item.Link.ToString();
                                completenessLibrary.CreatedBy = item.CreatedBy.Name.ToString();
                                completenessLibrary.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());


                                //Client Reference
                                AppItemField clientApp = item.Field<AppItemField>(clientNameField);
                                IEnumerable<Item> clientAppRef = clientApp.Items;
                                completenessLibrary.ClientName = clientAppRef.Select(x => x.Title).FirstOrDefault();

                                //Key Report Name
                                TextItemField txtKeyReportName = item.Field<TextItemField>(keyReportNameField);
                                completenessLibrary.KeyReportName = txtKeyReportName.Value;

                                //Method
                                AppItemField catMethod = item.Field<AppItemField>(methodField);
                                IEnumerable<Item> methodAppRef = catMethod.Items;
                                completenessLibrary.Method = methodAppRef.Select(x => x.Title).FirstOrDefault();

                                //Completeness
                                TextItemField txtCompleteness = item.Field<TextItemField>(completenessField);
                                completenessLibrary.Completeness = txtCompleteness.Value;

                                //C1
                                TextItemField txtC1 = item.Field<TextItemField>(c1Field);
                                completenessLibrary.C1 = txtC1.Value;

                                //C2
                                TextItemField txtC2 = item.Field<TextItemField>(c2Field);
                                completenessLibrary.C2 = txtC2.Value;

                                //C3
                                TextItemField txtC3 = item.Field<TextItemField>(c3Field);
                                completenessLibrary.C3 = txtC3.Value;

                                //C4
                                TextItemField txtC4 = item.Field<TextItemField>(c4Field);
                                completenessLibrary.C4 = txtC4.Value;

                                //C5
                                TextItemField txtC5 = item.Field<TextItemField>(c5Field);
                                completenessLibrary.C5 = txtC5.Value;

                                //C6
                                TextItemField txtC6 = item.Field<TextItemField>(c6Field);
                                completenessLibrary.C6 = txtC6.Value;

                                //C7
                                TextItemField txtC7 = item.Field<TextItemField>(c7Field);
                                completenessLibrary.C7 = txtC7.Value;

                                //C8
                                TextItemField txtC8 = item.Field<TextItemField>(c8Field);
                                completenessLibrary.C8 = txtC8.Value;

                                //C9
                                TextItemField txtC9 = item.Field<TextItemField>(c9Field);
                                completenessLibrary.C9 = txtC9.Value;

                                //C10
                                TextItemField txtC10 = item.Field<TextItemField>(c10Field);
                                completenessLibrary.C10 = txtC10.Value;
                                //Status
                                CategoryItemField catStatus = item.Field<CategoryItemField>(statusField);
                                IEnumerable<CategoryItemField.Option> listCatStatus = catStatus.Options;
                                completenessLibrary.status = listCatStatus.Select(x => x.Text).FirstOrDefault();

                                listCompleteness.Add(completenessLibrary);
                            }

                            count = count * (syncDateRange.offset + 1);
                            total = count >= collection.Filtered ? collection.Filtered : count;
                            totalItems = collection.Filtered;
                        }
                    }

                    if (listCompleteness != null && listCompleteness.Count() > 0)
                    {
                        using (var context = _soxContext.Database.BeginTransaction())
                        {

                            foreach (var item in listCompleteness)
                            {
                                var completenessCheck = _soxContext.CompletenessLibrary.FirstOrDefault(id => id.PodioItemId == (int)item.PodioItemId);
                                if (completenessCheck != null)
                                {
                                    completenessCheck.ClientName = item.ClientName;
                                    completenessCheck.KeyReportName = item.KeyReportName;
                                    completenessCheck.Method = item.Method;
                                    completenessCheck.Completeness = item.Completeness;
                                    completenessCheck.C1 = item.C1;
                                    completenessCheck.C2 = item.C2;
                                    completenessCheck.C3 = item.C3;
                                    completenessCheck.C4 = item.C4; 
                                    completenessCheck.C5 = item.C5;
                                    completenessCheck.C6 = item.C6;
                                    completenessCheck.C7 = item.C7;
                                    completenessCheck.C8 = item.C8;
                                    completenessCheck.C9 = item.C9;
                                    completenessCheck.C10 = item.C10;
                                    completenessCheck.PodioUniqueId = item.PodioUniqueId;
                                    completenessCheck.PodioRevision = item.PodioRevision;
                                    completenessCheck.PodioLink = item.PodioLink;
                                    completenessCheck.status = item.status;

                                    listCompletenessUpdate.Add(completenessCheck);
                                }
                                else
                                {
                                    listCompletenessAdd.Add(item);
                                }
                            }

                            if (listCompletenessAdd != null && listCompletenessAdd.Count() > 0)
                                _soxContext.AddRange(listCompletenessAdd);

                            if (listCompletenessUpdate != null && listCompletenessUpdate.Count() > 0)
                                _soxContext.UpdateRange(listCompletenessUpdate);

                            await _soxContext.SaveChangesAsync();
                            context.Commit();
                        }
                    }

                }
                string saveItems = $"{total}/{totalItems}";
                return Ok(new { syncDateRange.limit, totalItems, saveItems });

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncCompletenessAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncCompletenessAsync");
                return BadRequest(ex.ToString());
            }

            return Ok();

        }

        //[AllowAnonymous]
        [HttpPost("keyreport/Accuracy")]
        public async Task<IActionResult> SyncAccuracyAsync([FromBody] SyncDateRange syncDateRange)
        {
            //Console.WriteLine("Triggered process");
            try
            {

                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("KeyReportAccuracyId").Value;
                PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("KeyReportAccuracyToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);
                int totalItems = 0;
                int total = 0;
                int count = 0;
                if (podio.IsAuthenticated())
                {

                    var filter = new Dictionary<string, object>
                    {
                        {
                            "last_edit_on",
                            new  { from = syncDateRange.startDate, to = syncDateRange.endDate}
                        },
                    };

                    int clientNameField = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportAccuracyField").GetSection("ClientName").Value);
                    int keyReportNameField = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportAccuracyField").GetSection("KeyReportName").Value);
                    int methodField = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportAccuracyField").GetSection("Method").Value);
                    int accuracyField = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportAccuracyField").GetSection("Accuracy").Value);
                    int d1Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportAccuracyField").GetSection("D1").Value);
                    int d2Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportAccuracyField").GetSection("D2").Value);
                    int d3Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportAccuracyField").GetSection("D3").Value);
                    int d4Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportAccuracyField").GetSection("D4").Value);
                    int d5Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportAccuracyField").GetSection("D5").Value);
                    int d6Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportAccuracyField").GetSection("D6").Value);
                    int d7Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportAccuracyField").GetSection("D7").Value);
                    int d8Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportAccuracyField").GetSection("D8").Value);
                    int d9Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportAccuracyField").GetSection("D9").Value);
                    int d10Field = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportAccuracyField").GetSection("D10").Value);
                    int statusField = int.Parse(_config.GetSection("KeyReportApp").GetSection("KeyReportAccuracyField").GetSection("Status").Value);

                    PodioCollection<Item> collectionCheck = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 1, 0, filters: filter, sortDesc: true);

                    List<AccuracyLibrary> listAccuracy = new List<AccuracyLibrary>();
                    List<AccuracyLibrary> listAccuracyAdd = new List<AccuracyLibrary>();
                    List<AccuracyLibrary> listAccuracyUpdate = new List<AccuracyLibrary>();

                    if (syncDateRange.limit == 0 && syncDateRange.offset == 0)
                    {
                        if (collectionCheck.Filtered != 0)
                        {
                            int loop = (int)Math.Ceiling((double)collectionCheck.Filtered / 500);
                            int offset = 0;

                            //get filtered items and stored in list
                            for (int i = 0; i < loop; i++)
                            {
                                PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 500, offset, filters: filter, sortDesc: true);

                                foreach (var item in collection.Items)
                                {
                                    AccuracyLibrary accuracyLibrary = new AccuracyLibrary();
                                    count++;

                                    accuracyLibrary.PodioItemId = (int)item.ItemId;
                                    accuracyLibrary.PodioUniqueId = item.AppItemIdFormatted.ToString();
                                    accuracyLibrary.PodioRevision = item.CurrentRevision.Revision;
                                    accuracyLibrary.PodioLink = item.Link.ToString();
                                    accuracyLibrary.CreatedBy = item.CreatedBy.Name.ToString();
                                    accuracyLibrary.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());


                                    //Client Reference
                                    AppItemField clientApp = item.Field<AppItemField>(clientNameField);
                                    IEnumerable<Item> clientAppRef = clientApp.Items;
                                    accuracyLibrary.ClientName = clientAppRef.Select(x => x.Title).FirstOrDefault();
                                     
                                    
                                    //Key Report Name
                                    TextItemField txtKeyReportName = item.Field<TextItemField>(keyReportNameField);
                                    accuracyLibrary.KeyReportName = txtKeyReportName.Value;

                                    //Method
                                    AppItemField catMethod = item.Field<AppItemField>(methodField);
                                    IEnumerable<Item> methodAppRef = catMethod.Items;
                                    accuracyLibrary.Method = methodAppRef.Select(x => x.Title).FirstOrDefault();

                                    //Accuracy
                                    TextItemField txtAccuracy = item.Field<TextItemField>(accuracyField);
                                    accuracyLibrary.Accuracy = txtAccuracy.Value;

                                    //D1
                                    TextItemField txtD1 = item.Field<TextItemField>(d1Field);
                                    accuracyLibrary.D1 = txtD1.Value;

                                    //D2
                                    TextItemField txtD2 = item.Field<TextItemField>(d2Field);
                                    accuracyLibrary.D2 = txtD2.Value;

                                    //D3
                                    TextItemField txtD3 = item.Field<TextItemField>(d3Field);
                                    accuracyLibrary.D3 = txtD3.Value;

                                    //D4
                                    TextItemField txtD4 = item.Field<TextItemField>(d4Field);
                                    accuracyLibrary.D4 = txtD4.Value;

                                    //D5
                                    TextItemField txtD5 = item.Field<TextItemField>(d5Field);
                                    accuracyLibrary.D5 = txtD5.Value;

                                    //D6
                                    TextItemField txtD6 = item.Field<TextItemField>(d6Field);
                                    accuracyLibrary.D6 = txtD6.Value;

                                    //D7
                                    TextItemField txtD7 = item.Field<TextItemField>(d7Field);
                                    accuracyLibrary.D7 = txtD7.Value;

                                    //D8
                                    TextItemField txtD8 = item.Field<TextItemField>(d8Field);
                                    accuracyLibrary.D8 = txtD8.Value;

                                    //D9
                                    TextItemField txtD9 = item.Field<TextItemField>(d9Field);
                                    accuracyLibrary.D9 = txtD9.Value;

                                    //D10
                                    TextItemField txtD10 = item.Field<TextItemField>(d10Field);
                                    accuracyLibrary.D10 = txtD10.Value;

                                    //Status
                                    CategoryItemField catStatus = item.Field<CategoryItemField>(statusField);
                                    IEnumerable<CategoryItemField.Option> listCatStatus = catStatus.Options;
                                    accuracyLibrary.status = listCatStatus.Select(x => x.Text).FirstOrDefault();

                                    listAccuracy.Add(accuracyLibrary);
                                }

                                offset += 500;
                            }

                            total = count >= collectionCheck.Filtered ? collectionCheck.Filtered : count * (syncDateRange.offset + 1);
                            totalItems = collectionCheck.Filtered;


                        }
                    }
                    else
                    {
                        if (collectionCheck.Filtered != 0)
                        {
                            int offset = syncDateRange.offset;
                            PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), syncDateRange.limit, offset, filters: filter, sortDesc: true);

                            foreach (var item in collection.Items)
                            {
                                AccuracyLibrary accuracyLibrary = new AccuracyLibrary();
                                count++;

                                accuracyLibrary.PodioItemId = (int)item.ItemId;
                                accuracyLibrary.PodioUniqueId = item.AppItemIdFormatted.ToString();
                                accuracyLibrary.PodioRevision = item.CurrentRevision.Revision;
                                accuracyLibrary.PodioLink = item.Link.ToString();
                                accuracyLibrary.CreatedBy = item.CreatedBy.Name.ToString();
                                accuracyLibrary.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());


                                //Client Reference
                                AppItemField clientApp = item.Field<AppItemField>(clientNameField);
                                IEnumerable<Item> clientAppRef = clientApp.Items;
                                accuracyLibrary.ClientName = clientAppRef.Select(x => x.Title).FirstOrDefault();


                                //Key Report Name
                                TextItemField txtKeyReportName = item.Field<TextItemField>(keyReportNameField);
                                accuracyLibrary.KeyReportName = txtKeyReportName.Value;

                                //Method
                                AppItemField catMethod = item.Field<AppItemField>(methodField);
                                IEnumerable<Item> methodAppRef = catMethod.Items;
                                accuracyLibrary.Method = methodAppRef.Select(x => x.Title).FirstOrDefault();

                                //Accuracy
                                TextItemField txtAccuracy = item.Field<TextItemField>(accuracyField);
                                accuracyLibrary.Accuracy = txtAccuracy.Value;

                                //D1
                                TextItemField txtD1 = item.Field<TextItemField>(d1Field);
                                accuracyLibrary.D1 = txtD1.Value;

                                //D2
                                TextItemField txtD2 = item.Field<TextItemField>(d2Field);
                                accuracyLibrary.D2 = txtD2.Value;

                                //D3
                                TextItemField txtD3 = item.Field<TextItemField>(d3Field);
                                accuracyLibrary.D3 = txtD3.Value;

                                //D4
                                TextItemField txtD4 = item.Field<TextItemField>(d4Field);
                                accuracyLibrary.D4 = txtD4.Value;

                                //D5
                                TextItemField txtD5 = item.Field<TextItemField>(d5Field);
                                accuracyLibrary.D5 = txtD5.Value;

                                //D6
                                TextItemField txtD6 = item.Field<TextItemField>(d6Field);
                                accuracyLibrary.D6 = txtD6.Value;

                                //D7
                                TextItemField txtD7 = item.Field<TextItemField>(d7Field);
                                accuracyLibrary.D7 = txtD7.Value;

                                //D8
                                TextItemField txtD8 = item.Field<TextItemField>(d8Field);
                                accuracyLibrary.D8 = txtD8.Value;

                                //D9
                                TextItemField txtD9 = item.Field<TextItemField>(d9Field);
                                accuracyLibrary.D9 = txtD9.Value;

                                //D10
                                TextItemField txtD10 = item.Field<TextItemField>(d10Field);
                                accuracyLibrary.D10 = txtD10.Value;
                                //Status
                                CategoryItemField catStatus = item.Field<CategoryItemField>(statusField);
                                IEnumerable<CategoryItemField.Option> listCatStatus = catStatus.Options;
                                accuracyLibrary.status = listCatStatus.Select(x => x.Text).FirstOrDefault();

                                listAccuracy.Add(accuracyLibrary);
                            }

                            count = count * (syncDateRange.offset + 1);
                            total = count >= collection.Filtered ? collection.Filtered : count;
                            totalItems = collection.Filtered;
                        }
                    }

                    if (listAccuracy != null && listAccuracy.Count() > 0)
                    {
                        using (var context = _soxContext.Database.BeginTransaction())
                        {

                            foreach (var item in listAccuracy)
                            {
                                var accuracyCheck = _soxContext.AccuracyLibrary.FirstOrDefault(id => id.PodioItemId == (int)item.PodioItemId);
                                if (accuracyCheck != null)
                                {
                                    accuracyCheck.ClientName = item.ClientName;
                                    accuracyCheck.KeyReportName = item.KeyReportName;
                                    accuracyCheck.Method = item.Method;
                                    accuracyCheck.Accuracy = item.Accuracy;
                                    accuracyCheck.D1 = item.D1;
                                    accuracyCheck.D2 = item.D2;
                                    accuracyCheck.D3 = item.D3;
                                    accuracyCheck.D4 = item.D4;
                                    accuracyCheck.D5 = item.D5;
                                    accuracyCheck.D6 = item.D6;
                                    accuracyCheck.D7 = item.D7;
                                    accuracyCheck.D8 = item.D8;
                                    accuracyCheck.D9 = item.D9;
                                    accuracyCheck.D10 = item.D10;
                                    accuracyCheck.PodioUniqueId = item.PodioUniqueId;
                                    accuracyCheck.PodioRevision = item.PodioRevision;
                                    accuracyCheck.PodioLink = item.PodioLink;
                                    accuracyCheck.status = item.status;

                                    listAccuracyUpdate.Add(accuracyCheck);
                                }
                                else
                                {
                                    listAccuracyAdd.Add(item);
                                }
                            }

                            if (listAccuracyAdd != null && listAccuracyAdd.Count() > 0)
                                _soxContext.AddRange(listAccuracyAdd);

                            if (listAccuracyUpdate != null && listAccuracyUpdate.Count() > 0)
                                _soxContext.UpdateRange(listAccuracyUpdate);

                            await _soxContext.SaveChangesAsync();
                            context.Commit();
                        }
                    }

                }
                string saveItems = $"{total}/{totalItems}";
                return Ok(new { syncDateRange.limit, totalItems, saveItems });

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncCompletenessAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncCompletenessAsync");
                return BadRequest(ex.ToString());
            }

            return Ok();

        }
        //[AllowAnonymous]
        [HttpPost("keyreport/CAMethod")]
        public async Task<IActionResult> SyncCAMethodAsync([FromBody] SyncDateRange syncDateRange)
        {
            try
            {
                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("KeyReportApp").GetSection("C&A_MethodId").Value;
                PodioAppKey.AppToken = _config.GetSection("KeyReportApp").GetSection("C&A_MethodToken").Value;

                var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);
                int totalItems = 0;
                int total = 0;
                int count = 0;
                if (podio.IsAuthenticated())
                {

                    var filter = new Dictionary<string, object>
                    {
                        {
                            "last_edit_on",
                            new  { from = syncDateRange.startDate, to = syncDateRange.endDate}
                        },
                    };


                    int methodNameField = int.Parse(_config.GetSection("KeyReportApp").GetSection("C&A_MethodField").GetSection("MethodName").Value);
                    int methodTypeField = int.Parse(_config.GetSection("KeyReportApp").GetSection("C&A_MethodField").GetSection("MethodType").Value);
                    PodioCollection<Item> collectionCheck = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 1, 0, filters: filter, sortDesc: true);

                    List<CAMethodLibrary> listCAMethod = new List<CAMethodLibrary>();
                    List<CAMethodLibrary> listCAMethodAdd = new List<CAMethodLibrary>();
                    List<CAMethodLibrary> listCAMethodUpdate = new List<CAMethodLibrary>();

                    if (syncDateRange.limit == 0 && syncDateRange.offset == 0)
                    {
                        if (collectionCheck.Filtered != 0)
                        {
                            int loop = (int)Math.Ceiling((double)collectionCheck.Filtered / 500);
                            int offset = 0;

                            //get filtered items and stored in list
                            for (int i = 0; i < loop; i++)
                            {
                                PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), 500, offset, filters: filter, sortDesc: true);

                                foreach (var item in collection.Items)
                                {
                                    CAMethodLibrary CAMethodLibrary = new CAMethodLibrary();
                                    count++;

                                    CAMethodLibrary.PodioItemId = (int)item.ItemId;
                                    CAMethodLibrary.PodioUniqueId = item.AppItemIdFormatted.ToString();
                                    CAMethodLibrary.PodioRevision = item.CurrentRevision.Revision;
                                    CAMethodLibrary.PodioLink = item.Link.ToString();
                                    CAMethodLibrary.CreatedBy = item.CreatedBy.Name.ToString();
                                    CAMethodLibrary.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());

                                    
                                    //Method Type
                                    CategoryItemField methodApp = item.Field<CategoryItemField>(methodTypeField);
                                    IEnumerable<CategoryItemField.Option> methodAppRef = methodApp.Options;
                                    CAMethodLibrary.MethodType = methodAppRef.Select(x => x.Text).FirstOrDefault();

                                    //Key Report Name
                                    TextItemField txtKeyReportName = item.Field<TextItemField>(methodNameField);
                                    CAMethodLibrary.MethodName = txtKeyReportName.Value;

                                    
                                    listCAMethod.Add(CAMethodLibrary);
                                }

                                offset += 500;
                            }

                            total = count >= collectionCheck.Filtered ? collectionCheck.Filtered : count * (syncDateRange.offset + 1);
                            totalItems = collectionCheck.Filtered;


                        }
                    }
                    else
                    {
                        if (collectionCheck.Filtered != 0)
                        {
                            int offset = syncDateRange.offset;
                            PodioCollection<Item> collection = await podio.ItemService.FilterItems(appId: Int32.Parse(PodioAppKey.AppId), syncDateRange.limit, offset, filters: filter, sortDesc: true);

                            foreach (var item in collection.Items)
                            {
                                CAMethodLibrary CAMethodLibrary = new CAMethodLibrary();
                                count++;

                                CAMethodLibrary.PodioItemId = (int)item.ItemId;
                                CAMethodLibrary.PodioUniqueId = item.AppItemIdFormatted.ToString();
                                CAMethodLibrary.PodioRevision = item.CurrentRevision.Revision;
                                CAMethodLibrary.PodioLink = item.Link.ToString();
                                CAMethodLibrary.CreatedBy = item.CreatedBy.Name.ToString();
                                CAMethodLibrary.CreatedOn = DateTime.Parse(item.CreatedOn.ToString());

                                //Method Type
                                CategoryItemField methodApp = item.Field<CategoryItemField>(methodTypeField);
                                IEnumerable<CategoryItemField.Option> methodAppRef = methodApp.Options;
                                CAMethodLibrary.MethodType = methodAppRef.Select(x => x.Text).FirstOrDefault();

                                //Key Report Name
                                TextItemField txtKeyReportName = item.Field<TextItemField>(methodNameField);
                                CAMethodLibrary.MethodName = txtKeyReportName.Value;


                                listCAMethod.Add(CAMethodLibrary);
                            }

                            count = count * (syncDateRange.offset + 1);
                            total = count >= collection.Filtered ? collection.Filtered : count;
                            totalItems = collection.Filtered;
                        }
                    }

                    if (listCAMethod != null && listCAMethod.Count() > 0)
                    {
                        using (var context = _soxContext.Database.BeginTransaction())
                        {

                            foreach (var item in listCAMethod)
                            {
                                var camethodCheck = _soxContext.CAMethodLibrary.FirstOrDefault(id => id.PodioItemId == (int)item.PodioItemId);
                                if (camethodCheck != null)
                                {
                                    camethodCheck.MethodType = item.MethodType;
                                    camethodCheck.MethodName = item.MethodName;
                                    camethodCheck.PodioUniqueId = item.PodioUniqueId;
                                    camethodCheck.PodioRevision = item.PodioRevision;
                                    camethodCheck.PodioLink = item.PodioLink;

                                    listCAMethodUpdate.Add(camethodCheck);
                                }
                                else
                                {
                                    listCAMethodAdd.Add(item);
                                }
                            }

                            if (listCAMethodAdd != null && listCAMethodAdd.Count() > 0)
                                _soxContext.AddRange(listCAMethodAdd);

                            if (listCAMethodUpdate != null && listCAMethodUpdate.Count() > 0)
                                _soxContext.UpdateRange(listCAMethodUpdate);

                            await _soxContext.SaveChangesAsync();
                            context.Commit();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSyncCAMethodLibraryAsync");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SyncCAMethodsAsync");
                return BadRequest(ex.ToString());
            }
            return Ok();
        }


        private static bool TryParseJSONSkypeObj(string json, out SkypeObj skypeObj)
        {
            try
            {
                //jObject = JObject.Parse(json);
                skypeObj = JsonConvert.DeserializeObject<SkypeObj>(json);
                return true;
            }
            catch
            {
                skypeObj = null;
                return false;
            }
        }

    }

}
