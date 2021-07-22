using A2B_App.Server.Data;
//using A2B_App.Server.Log;
using A2B_App.Shared.Podio;
using A2B_App.Shared.Sox;
using A2B_App.Shared.Time;
using Microsoft.Extensions.Configuration;
using PodioAPI;
using PodioAPI.Models;
using PodioAPI.Utils.ItemFields;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Threading.Tasks;

namespace A2B_App.Server.Services
{
    public class TimeService
    {

        private TimeContext _timeContext;
        //private PodioField _podioField;
        private readonly IConfiguration _config;
        public TimeService(TimeContext timeContext, IConfiguration config)
        {
            _timeContext = timeContext;
            _config = config;
        }

        /// <summary>
        /// Save Podio Webhook
        /// </summary>
        /// <param name="podioHook"></param>
        /// <returns></returns>
        public bool SaveWebhookAsync(PodioHook podioHook)
        {

            bool result = false;
            try
            {
                podioHook.DateCreated = DateTime.Now.ToUniversalTime();
                //podioHook.DateUpdated = DateTime.Now.ToUniversalTime();
                _timeContext.Add<PodioHook>(podioHook);
                _timeContext.SaveChanges();
                result = true;
            }
            catch (Exception ex)
            {
                FileLog.Write($"Error SaveWebhookSampleSelection {ex}", "ErrorSaveWebhookSampleSelection");
            }

            return result;
        }

        /// <summary>
        /// Save Podio Time Code
        /// </summary>
        /// <param name="ItemId"></param>
        /// <returns>bool</returns>
        public async Task<bool> SavePodioTimeCodeAsync(int ItemId)
        {
            bool result = false;

            using (var context = _timeContext.Database.BeginTransaction())
            {
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

                    if (podio.IsAuthenticated())
                    {

                        var podioItem = podio.ItemService.GetItem(ItemId);
                        var rawResponse = Newtonsoft.Json.JsonConvert.SerializeObject(podioItem.Result);
                        FileLog.Write($"{rawResponse}", "SavePodioTimeCodeAsync");

                        TimeCode timeCode = new TimeCode();

                        PodioRef podioRef = new PodioRef();

                        //check if item_id already exists in the database
                        bool isExists = false;
                        var timeCodeCheck = _timeContext.PodioRef.FirstOrDefault(id => id.ItemId == (int)podioItem.Result.ItemId);
                        if (timeCodeCheck != null)
                        {
                            podioRef = timeCodeCheck;
                            isExists = true;
                        }

                        #region Podio Item
                        podioRef.ItemId = (int)podioItem.Result.ItemId;
                        podioRef.UniqueId = podioItem.Result.AppItemIdFormatted.ToString();
                        podioRef.Revision = podioItem.Result.CurrentRevision.Revision;
                        podioRef.Link = podioItem.Result.Link.ToString();
                        podioRef.CreatedBy = podioItem.Result.CreatedBy.Name.ToString();
                        podioRef.CreatedOn = DateTime.Parse(podioItem.Result.CreatedOn.ToString());

                        string fieldClientRef = _config.GetSection("TimePodioApp").GetSection("TimeCodeField").GetSection("ClientRef").Value;
                        string fieldProjectRef = _config.GetSection("TimePodioApp").GetSection("TimeCodeField").GetSection("ProjectRef").Value;
                        string fieldTaskRef = _config.GetSection("TimePodioApp").GetSection("TimeCodeField").GetSection("TaskRef").Value;
                        string fieldStatus = _config.GetSection("TimePodioApp").GetSection("TimeCodeField").GetSection("CatStatus").Value;

                        #region podio fields

                        //client reference field
                        AppItemField clientRef = podioItem.Result.Field<AppItemField>(fieldClientRef);
                        IEnumerable<Item> clientAppRef = clientRef.Items;
                        timeCode.ClientRef = clientAppRef.Select(x => x.Title).FirstOrDefault();
                        timeCode.ClientRefId = clientAppRef.Select(x => x.ItemId).FirstOrDefault();

                        //project reference field
                        AppItemField projectRef = podioItem.Result.Field<AppItemField>(fieldProjectRef);
                        IEnumerable<Item> projectAppRef = projectRef.Items;
                        timeCode.ProjectRef = projectAppRef.Select(x => x.Title).FirstOrDefault();
                        timeCode.ProjectRefId = projectAppRef.Select(x => x.ItemId).FirstOrDefault();

                        //task reference field
                        AppItemField taskRef = podioItem.Result.Field<AppItemField>(fieldTaskRef);
                        IEnumerable<Item> taskAppRef = taskRef.Items;
                        timeCode.TaskRef = taskAppRef.Select(x => x.Title).FirstOrDefault();
                        timeCode.TaskRefId = taskAppRef.Select(x => x.ItemId).FirstOrDefault();

                        //Risk field
                        CategoryItemField catStatus = podioItem.Result.Field<CategoryItemField>(fieldStatus);
                        IEnumerable<CategoryItemField.Option> listcatRisk = catStatus.Options;
                        timeCode.Status = listcatRisk.Select(x => x.Text).FirstOrDefault();

                        #endregion

                        #endregion

                        ////Check click code
                        //if (timeCode.ClientRefId != 0)
                        //{
                        //    var clientRefCheck = _timeContext.ClientReference.FirstOrDefault(id => id.PodioItemId == (int)timeCode.ClientRefId);
                        //    if (clientRefCheck != null)
                        //    {
                        //        timeCode.ClientCode = clientRefCheck.ClientCode;
                        //    }
                        //}

                        //if (isExists)
                        //{
                        //    //if exists we update database row
                        //    _timeContext.Update(timeCode);
                        //}
                        //else
                        //{
                        //    //create new item
                        //    _timeContext.Add(timeCode);
                        //}

                        //await _timeContext.SaveChangesAsync();
                        //context.Commit();
                    }
                }

                catch (Exception ex)
                {
                    await context.RollbackAsync();
                    FileLog.Write($"Error SavePodioTimeCodeAsync {ex}", "ErrorSavePodioTimeCodeAsync");
                }

            }

            result = true;
            return result;
        }

        /// <summary>
        /// Delete Time Code
        /// </summary>
        /// <param name="ItemId"></param>
        /// <returns>bool</returns>
        public async Task<bool> DeleteTimeCodeAsync(int ItemId)
        {
            bool result = false;

            using (var context = _timeContext.Database.BeginTransaction())
            {
                try
                {
                    //    var timeCodeCheck = _timeContext.TimeCode.FirstOrDefault(id => id.PodioItemId == (int)ItemId);

                    //    if (timeCodeCheck != null)
                    //    {

                    //        _timeContext.Remove(timeCodeCheck);

                    //        await _timeContext.SaveChangesAsync();
                    //        context.Commit();
                    //    }
                }
                catch (Exception ex)
                {
                    await context.RollbackAsync();
                    FileLog.Write($"Error DeleteTimeCodeAsync {ex}", "ErrorDeleteTimeCodeAsync");
                }
            }
            result = true;
            return result;
        }

        /// <summary>
        /// Save Podio Client Reference
        /// </summary>
        /// <param name="ItemId"></param>
        /// <returns>bool</returns>
        public async Task<bool> SavePodioClientRefAsync(int ItemId)
        {
            bool result = false;

            using (var context = _timeContext.Database.BeginTransaction())
            {
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

                    if (podio.IsAuthenticated())
                    {

                        //var podioItem = podio.ItemService.GetItem(ItemId);
                        //var rawResponse = Newtonsoft.Json.JsonConvert.SerializeObject(podioItem.Result);
                        //FileLog.Write($"{rawResponse}", "SavePodioClientRefAsync");

                        //ClientReference clientReference = new ClientReference();

                        ////check if item_id already exists in the database
                        //bool isExists = false;
                        //var clientRefCheck = _timeContext.ClientReference.FirstOrDefault(id => id.PodioItemId == (int)podioItem.Result.ItemId);
                        //if (clientRefCheck != null)
                        //{
                        //    clientReference = clientRefCheck;
                        //    isExists = true;
                        //}

                        //#region Podio Item
                        //clientReference.PodioItemId = (int)podioItem.Result.ItemId;
                        //clientReference.PodioUniqueId = podioItem.Result.AppItemIdFormatted.ToString();
                        //clientReference.PodioRevision = podioItem.Result.CurrentRevision.Revision;
                        //clientReference.PodioLink = podioItem.Result.Link.ToString();
                        //clientReference.CreatedBy = podioItem.Result.CreatedBy.Name.ToString();
                        //clientReference.CreatedOn = DateTime.Parse(podioItem.Result.CreatedOn.ToString());

                        //#region podio fields

                        ////Client field
                        //TextItemField txtClient = podioItem.Result.Field<TextItemField>("cliend-reference-code");
                        //clientReference.ClientRef = txtClient.Value;

                        ////Client code field
                        //TextItemField txtClientCode = podioItem.Result.Field<TextItemField>("unique-identifier-code");
                        //clientReference.ClientCode = txtClientCode.Value;

                        //#endregion

                        //#endregion

                        //if (isExists)
                        //{
                        //    //if exists we update database row
                        //    _timeContext.Update(clientReference);
                        //}
                        //else
                        //{
                        //    //create new item
                        //    _timeContext.Add(clientReference);
                        //}


                        //await _timeContext.SaveChangesAsync();
                        //context.Commit();
                    }
                }

                catch (Exception ex)
                {
                    await context.RollbackAsync();
                    FileLog.Write($"Error SavePodioClientRefAsync {ex}", "ErrorSavePodioClientRefAsync");
                }

            }

            result = true;

            return result;
        }

        /// <summary>
        /// Delete Client Reference
        /// </summary>
        /// <param name="ItemId"></param>
        /// <returns>bool</returns>
        public async Task<bool> DeleteClientRefAsync(int ItemId)
        {
            bool result = false;

            using (var context = _timeContext.Database.BeginTransaction())
            {
                try
                {
                    //var clientRefCheck = _timeContext.ClientReference.FirstOrDefault(id => id.PodioItemId == (int)ItemId);

                    //if (clientRefCheck != null)
                    //{

                    //    _timeContext.Remove(clientRefCheck);

                    //    await _timeContext.SaveChangesAsync();
                    //    context.Commit();
                    //}



                }

                catch (Exception ex)
                {
                    await context.RollbackAsync();
                    FileLog.Write($"Error DeleteClientRefAsync {ex}", "ErrorDeleteClientRefAsync");
                }

            }

            result = true;

            return result;
        }

        /// <summary>
        /// Save Podio Project Reference
        /// </summary>
        /// <param name="ItemId"></param>
        /// <returns>bool</returns>
        public async Task<bool> SavePodioProjectRefAsync(int ItemId)
        {
            bool result = false;

            using (var context = _timeContext.Database.BeginTransaction())
            {
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

                    if (podio.IsAuthenticated())
                    {

                        //var podioItem = podio.ItemService.GetItem(ItemId);
                        //var rawResponse = Newtonsoft.Json.JsonConvert.SerializeObject(podioItem.Result);
                        //FileLog.Write($"{rawResponse}", "SavePodioProjectRefAsync");

                        //ProjectReference projectReference = new ProjectReference();

                        ////check if item_id already exists in the database
                        //bool isExists = false;
                        //var projectRefCheck = _timeContext.ProjectReference.FirstOrDefault(id => id.PodioItemId == (int)podioItem.Result.ItemId);
                        //if (projectRefCheck != null)
                        //{
                        //    projectReference = projectRefCheck;
                        //    isExists = true;
                        //}

                        //#region Podio Item
                        //projectReference.PodioItemId = (int)podioItem.Result.ItemId;
                        //projectReference.PodioUniqueId = podioItem.Result.AppItemIdFormatted.ToString();
                        //projectReference.PodioRevision = podioItem.Result.CurrentRevision.Revision;
                        //projectReference.PodioLink = podioItem.Result.Link.ToString();
                        //projectReference.CreatedBy = podioItem.Result.CreatedBy.Name.ToString();
                        //projectReference.CreatedOn = DateTime.Parse(podioItem.Result.CreatedOn.ToString());

                        //#region podio fields

                        ////Task field
                        //TextItemField txtProject = podioItem.Result.Field<TextItemField>("title");
                        //projectReference.ProjectRef = txtProject.Value;

                        //#endregion

                        //#endregion

                        //if (isExists)
                        //{
                        //    //if exists we update database row
                        //    _timeContext.Update(projectReference);
                        //}
                        //else
                        //{
                        //    //create new item
                        //    _timeContext.Add(projectReference);
                        //}


                        //await _timeContext.SaveChangesAsync();
                        //context.Commit();
                    }
                }

                catch (Exception ex)
                {
                    await context.RollbackAsync();
                    FileLog.Write($"Error SavePodioProjectRefAsync {ex}", "ErrorSavePodioProjectRefAsync");
                }

            }

            result = true;

            return result;
        }

        /// <summary>
        /// Delete Project Reference
        /// </summary>
        /// <param name="ItemId"></param>
        /// <returns>bool</returns>
        public async Task<bool> DeleteProjectRefAsync(int ItemId)
        {
            bool result = false;

            using (var context = _timeContext.Database.BeginTransaction())
            {
                try
                {
                    //var projectRefCheck = _timeContext.ProjectReference.FirstOrDefault(id => id.PodioItemId == (int)ItemId);

                    //if (projectRefCheck != null)
                    //{

                    //    _timeContext.Remove(projectRefCheck);

                    //    await _timeContext.SaveChangesAsync();
                    //    context.Commit();
                    //}
                }
                catch (Exception ex)
                {
                    await context.RollbackAsync();
                    FileLog.Write($"Error DeleteProjectRefAsync {ex}", "ErrorDeleteProjectRefAsync");
                }
            }

            result = true;
            return result;
        }

        /// <summary>
        /// Save Podio Task Reference
        /// </summary>
        /// <param name="ItemId"></param>
        /// <returns>bool</returns>
        public async Task<bool> SavePodioTaskRefAsync(int ItemId)
        {
            bool result = false;

            using (var context = _timeContext.Database.BeginTransaction())
            {
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

                        //var podioItem = podio.ItemService.GetItem(ItemId);
                        //var rawResponse = Newtonsoft.Json.JsonConvert.SerializeObject(podioItem.Result);
                        //FileLog.Write($"{rawResponse}", "SavePodioTaskRefAsync");

                        //TaskReference taskReference = new TaskReference();

                        ////check if item_id already exists in the database
                        //bool isExists = false;
                        //var taskRefCheck = _timeContext.TaskReference.FirstOrDefault(id => id.PodioItemId == (int)podioItem.Result.ItemId);
                        //if (taskRefCheck != null)
                        //{
                        //    taskReference = taskRefCheck;
                        //    isExists = true;
                        //}

                        //#region Podio Item
                        //taskReference.PodioItemId = (int)podioItem.Result.ItemId;
                        //taskReference.PodioUniqueId = podioItem.Result.AppItemIdFormatted.ToString();
                        //taskReference.PodioRevision = podioItem.Result.CurrentRevision.Revision;
                        //taskReference.PodioLink = podioItem.Result.Link.ToString();
                        //taskReference.CreatedBy = podioItem.Result.CreatedBy.Name.ToString();
                        //taskReference.CreatedOn = DateTime.Parse(podioItem.Result.CreatedOn.ToString());

                        //#region podio fields

                        ////Task field
                        //TextItemField txtTask = podioItem.Result.Field<TextItemField>("title");
                        //taskReference.TaskRef = txtTask.Value;

                        //#endregion

                        //#endregion

                        //if (isExists)
                        //{
                        //    //if exists we update database row
                        //    _timeContext.Update(taskReference);
                        //}
                        //else
                        //{
                        //    //create new item
                        //    _timeContext.Add(taskReference);
                        //}
                        //await _timeContext.SaveChangesAsync();
                        //context.Commit();
                    }
                }

                catch (Exception ex)
                {
                    await context.RollbackAsync();
                    FileLog.Write($"Error SavePodioTaskRefAsync {ex}", "ErrorSavePodioTaskRefAsync");
                }

            }

            result = true;

            return result;
        }

        /// <summary>
        /// Delete Task Reference
        /// </summary>
        /// <param name="ItemId"></param>
        /// <returns>bool</returns>
        public async Task<bool> DeleteTaskRefAsync(int ItemId)
        {
            bool result = false;
            using (var context = _timeContext.Database.BeginTransaction())
            {
                try
                {
                    //var taskCodeCheck = _timeContext.TaskReference.FirstOrDefault(id => id.PodioItemId == (int)ItemId);

                    //if (taskCodeCheck != null)
                    //{

                    //    _timeContext.Remove(taskCodeCheck);

                    //    await _timeContext.SaveChangesAsync();
                    //    context.Commit();
                    //}

                }
                catch (Exception ex)
                {
                    await context.RollbackAsync();
                    FileLog.Write($"Error DeleteTaskRefAsync {ex}", "ErrorDeleteTaskRefAsync");
                }
            }
            result = true;
            return result;
        }

    }
}
