using A2B_App.Server.Data;
using A2B_App.Server.Services;
using A2B_App.Shared.Podio;
using A2B_App.Shared.Sox;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using MySql.Data.MySqlClient;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using PodioAPI;
using PodioAPI.Models;
using PodioAPI.Utils;
using PodioAPI.Utils.ItemFields;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading.Tasks;


namespace A2B_App.Server.Controllers
{
    [Authorize]
    [Route("api/[controller]")]
    [ApiController]
    public class QuestionnaireController : ControllerBase
    {

        private readonly IConfiguration _config;
        private readonly ILogger<QuestionnaireController> _logger;
        private readonly IWebHostEnvironment _environment;
        private readonly SoxContext _soxContext;

        public QuestionnaireController(IConfiguration config,
            ILogger<QuestionnaireController> logger,
            IWebHostEnvironment environment,
            SoxContext soxContext)
        {
            _config = config;
            _logger = logger;
            _environment = environment;
            _soxContext = soxContext;
        }

        [AllowAnonymous]
        [HttpGet("download/{filename}")]
        public async Task<IActionResult> GetDownloadAsync(string filename)
        {
            //test link = https://localhost:44344/SampleSelection/download/Draft-TestRound-20200408_183306.xlsx

            try
            {
                string startupPath = Directory.GetCurrentDirectory();
                //string path = Path.Combine(startupPath, "include", "questionnaire", "download", filename); 032321
                string path = Path.Combine(startupPath, "include", "upload", "soxquestionnaire", filename);
                var memory = new MemoryStream();
                using (var stream = new FileStream(path, FileMode.Open))
                {
                    await stream.CopyToAsync(memory);
                }
                memory.Position = 0;
                var ext = Path.GetExtension(path).ToLowerInvariant();

                return File(memory, GetMimeTypes()[ext], Path.GetFileName(path));
            }
            catch (Exception ex)
            {
                FileLog.Write($"Error GetDownloadAsync {ex}", "ErrorGetDownloadAsync");
                AdminService adminService = new AdminService(_config);
               // adminService.SendAlert(true, true, ex.ToString(), "GetDownloadAsync");
                if (_environment.IsDevelopment())
                {
                    return BadRequest("Oops something when wrong!." + ex.ToString());
                }
                else
                {
                    return BadRequest("Oops something when wrong!.");
                }

            }
        }

        //[AllowAnonymous]
        [HttpGet("allroundset/{rcmItemId:int}")]
        public IActionResult GetRoundSet(int rcmItemId)
        {

            try
            {
                var rcmCheck = _soxContext.Rcm.AsNoTracking().FirstOrDefault(x => x.PodioItemId == rcmItemId);
                if (rcmCheck != null)
                {
                    var roundSetCheck = _soxContext.QuestionnaireRoundSet
                        .OrderByDescending(x => x.Id)
                        .Where(id => id.Rcm.Id == rcmCheck.Id)
                        .AsNoTracking();
                    if (roundSetCheck != null)
                    {
                        List<QuestionnaireRoundSet> listRoundSet = new List<QuestionnaireRoundSet>();
                        foreach (var item in roundSetCheck)
                        {
                            //item.Rcm = rcmCheck;
                            listRoundSet.Add(item);
                        }
                        
                        return Ok(listRoundSet);
                    }
                }
   
            }
            catch (Exception ex)
            {
                FileLog.Write($"Error GetRoundSet {ex}", "ErrorGetRoundSet");
                AdminService adminService = new AdminService(_config);
               // adminService.SendAlert(true, true, ex.ToString(), "GetRoundSet");
                return BadRequest("Oops something when wrong!." + ex.ToString());
            }

            return BadRequest($"Bad request..");
        }

        //[AllowAnonymous]
        [HttpGet("roundset/{uniqueid}")]
        public IActionResult GetSpecificRoundSet(string uniqueid)
        {
            //test link = https://localhost:44344/SampleSelection/download/Draft-TestRound-20200408_183306.xlsx

            try
            {

                var roundSetCheck = _soxContext.QuestionnaireRoundSet
                    .Include(x => x.sampleSel1)
                    .Include(x => x.sampleSel2)
                    .Include(x => x.sampleSel3)
                    .Where(id => id.UniqueId == uniqueid)
                    .AsNoTracking()
                    .FirstOrDefault();
                if (roundSetCheck != null)
                {
                    //Get sample selection test round 1
                    if (roundSetCheck.sampleSel1 != null)
                    {

                        var resTestround1 = _soxContext.TestRounds.FromSqlRaw($"CALL `sox`.`sp_get_testround`({roundSetCheck.sampleSel1.Id});").ToList();
                        if (resTestround1.Count > 0)
                        {
                            roundSetCheck.sampleSel1.ListTestRound = resTestround1;
                        }


                        //var sampleSel1 = _soxContext.SampleSelection
                        //    .Where(x => x.Id == roundSetCheck.sampleSel1.Id)
                        //    .Include(x => x.ListTestRound)
                        //    .Select(x => x.ListTestRound)
                        //    .AsNoTracking()
                        //    .FirstOrDefault();
                        //if (sampleSel1 != null)
                        //{
                        //    roundSetCheck.sampleSel1.ListTestRound = sampleSel1;
                        //}
                    }
                    //Get sample selection test round 2
                    if (roundSetCheck.sampleSel2 != null)
                    {
                        var resTestround2 = _soxContext.TestRounds.FromSqlRaw($"CALL `sox`.`sp_get_testround`({roundSetCheck.sampleSel2.Id});").ToList();
                        if (resTestround2.Count > 0)
                        {
                            roundSetCheck.sampleSel2.ListTestRound = resTestround2;
                        }

                    }
                    //Get sample selection test round 3
                    if (roundSetCheck.sampleSel3 != null)
                    {

                        var resTestround3 = _soxContext.TestRounds.FromSqlRaw($"CALL `sox`.`sp_get_testround`({roundSetCheck.sampleSel3.Id});").ToList();
                        if (resTestround3.Count > 0)
                        {
                            roundSetCheck.sampleSel3.ListTestRound = resTestround3;
                        }

                    }


                    WriteLog writeLog = new WriteLog();

                    var resQuestionnaireUserAnswer = _soxContext.QuestionnaireUserAnswer.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_user_answer`({roundSetCheck.Id});").ToList();
                    if(resQuestionnaireUserAnswer.Any())
                    {
                        roundSetCheck.ListUserInputRound1 = resQuestionnaireUserAnswer
                            .OrderBy(x => x.Position)
                            .Where(x => x.RoundName.Equals("Round 1"))
                            .ToList();

                        roundSetCheck.ListUserInputRound2 = resQuestionnaireUserAnswer
                            .OrderBy(x => x.Position)
                            .Where(x => x.RoundName.Equals("Round 2"))
                            .ToList();

                        roundSetCheck.ListUserInputRound3 = resQuestionnaireUserAnswer
                            .OrderBy(x => x.Position)
                            .Where(x => x.RoundName.Equals("Round 3"))
                            .ToList();
                    }

                    var resNotesItem = _soxContext.NotesItem.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_notesitem`({roundSetCheck.Id});").ToList();
                    if (resNotesItem.Any())
                    {
                        roundSetCheck.ListUniqueNotes = resNotesItem
                            .OrderBy(x => x.Id)
                            .ToList();
                    }

                    var resListRound = _soxContext.RoundItem.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_rounditem`({roundSetCheck.Id});").ToList();
                    if (resListRound.Any())
                    {
                        roundSetCheck.ListRoundItem = resListRound
                            .OrderBy(x => x.Id)
                            .ToList();
                    }

                    #region IUC System Gen

                    var resIUCSystemGenAnswer1 = _soxContext.IUCSystemGenAnswer.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_iuc_system_gen`({roundSetCheck.Id},'Round 1');").ToList();
                    if (resIUCSystemGenAnswer1.Any())
                    {
                        roundSetCheck.ListIUCSystemGen1 = resIUCSystemGenAnswer1.ToList();
                        if (roundSetCheck.ListIUCSystemGen1.Count > 0)
                        {
                            foreach (var item in roundSetCheck.ListIUCSystemGen1)
                            {
                                var resIUCSystemGenQuestion1 = _soxContext.IUCQuestionUserAnswer.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_iuc_user_input`({item.Id},'1');").ToList();
                                if (resIUCSystemGenQuestion1.Count > 0)
                                {
                                    item.ListQuestionAnswer = resIUCSystemGenQuestion1;
                                }
                            }
                        }
                    }

                    var resIUCSystemGenAnswer2 = _soxContext.IUCSystemGenAnswer.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_iuc_system_gen`({roundSetCheck.Id},'Round 2');").ToList();
                    if (resIUCSystemGenAnswer2.Any())
                    {
                        roundSetCheck.ListIUCSystemGen2 = resIUCSystemGenAnswer2.ToList();
                        if (roundSetCheck.ListIUCSystemGen2.Count > 0)
                        {
                            foreach (var item in roundSetCheck.ListIUCSystemGen2)
                            {
                                var resIUCSystemGenQuestion2 = _soxContext.IUCQuestionUserAnswer.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_iuc_user_input`({item.Id},'1');").ToList();
                                if (resIUCSystemGenQuestion2.Count > 0)
                                {
                                    item.ListQuestionAnswer = resIUCSystemGenQuestion2;
                                }
                            }
                        }
                    }

                    var resIUCSystemGenAnswer3 = _soxContext.IUCSystemGenAnswer.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_iuc_system_gen`({roundSetCheck.Id},'Round 3');").ToList();
                    if (resIUCSystemGenAnswer3.Any())
                    {
                        roundSetCheck.ListIUCSystemGen3 = resIUCSystemGenAnswer3.ToList();
                        if (roundSetCheck.ListIUCSystemGen3.Count > 0)
                        {
                            foreach (var item in roundSetCheck.ListIUCSystemGen3)
                            {
                                var resIUCSystemGenQuestion3 = _soxContext.IUCQuestionUserAnswer.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_iuc_user_input`({item.Id},'1');").ToList();
                                if (resIUCSystemGenQuestion3.Count > 0)
                                {
                                    item.ListQuestionAnswer = resIUCSystemGenQuestion3;
                                }
                            }
                        }
                    }

                    #endregion

                    #region IUC Non System Gen

                    var resIUCNonSystemGenAnswer1 = _soxContext.IUCNonSystemGenAnswer.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_iuc_non_system_gen`({roundSetCheck.Id},'Round 1');").ToList();
                    if (resIUCNonSystemGenAnswer1.Any())
                    {
                        roundSetCheck.ListIUCNonSystemGen1 = resIUCNonSystemGenAnswer1.ToList();
                        if(roundSetCheck.ListIUCNonSystemGen1.Count > 0)
                        {
                            foreach (var item in roundSetCheck.ListIUCNonSystemGen1)
                            {
                                var resIUCNonSystemGenQuestion1 = _soxContext.IUCQuestionUserAnswer.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_iuc_user_input`({item.Id},'2');").ToList();
                                if (resIUCNonSystemGenQuestion1.Count > 0)
                                {
                                    item.ListQuestionAnswer = resIUCNonSystemGenQuestion1;
                                }
                            }
                        }
                    }

                    var resIUCNonSystemGenAnswer2 = _soxContext.IUCNonSystemGenAnswer.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_iuc_non_system_gen`({roundSetCheck.Id},'Round 2');").ToList();
                    if (resIUCNonSystemGenAnswer2.Any())
                    {
                        roundSetCheck.ListIUCNonSystemGen2 = resIUCNonSystemGenAnswer2.ToList();
                        if (roundSetCheck.ListIUCNonSystemGen2.Count > 0)
                        {
                            foreach (var item in roundSetCheck.ListIUCNonSystemGen2)
                            {
                                var resIUCNonSystemGenQuestion2 = _soxContext.IUCQuestionUserAnswer.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_iuc_user_input`({item.Id},'2');").ToList();
                                if (resIUCNonSystemGenQuestion2.Count > 0)
                                {
                                    item.ListQuestionAnswer = resIUCNonSystemGenQuestion2;
                                }
                            }

                        }
                    }

                    var resIUCNonSystemGenAnswer3 = _soxContext.IUCNonSystemGenAnswer.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_iuc_non_system_gen`({roundSetCheck.Id},'Round 3');").ToList();
                    if (resIUCNonSystemGenAnswer3.Any())
                    {
                        roundSetCheck.ListIUCNonSystemGen3 = resIUCNonSystemGenAnswer3.ToList();
                        if (roundSetCheck.ListIUCNonSystemGen3.Count > 0)
                        {
                            foreach (var item in roundSetCheck.ListIUCNonSystemGen3)
                            {
                                var resIUCNonSystemGenQuestion3 = _soxContext.IUCQuestionUserAnswer.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_iuc_user_input`({item.Id},'2');").ToList();
                                if (resIUCNonSystemGenQuestion3.Count > 0)
                                {
                                    item.ListQuestionAnswer = resIUCNonSystemGenQuestion3;
                                }
                            }

                        }
                    }


                    #endregion

                    var resHeaderNote = _soxContext.HeaderNote.FromSqlRaw($"CALL `sox`.`sp_get_headernote`({roundSetCheck.Id});").ToList();
                    if (resHeaderNote.Any())
                    {
                        roundSetCheck.ListHeaderNote = resHeaderNote.ToList();
                    }

                    //if (roundSetCheck.ListIUCNonSystemGen1 != null && roundSetCheck.ListIUCNonSystemGen1.Count > 0)
                    //{
                    //    foreach (var item in roundSetCheck.ListIUCNonSystemGen1)
                    //    {
                    //        var iucNonSys1 = _soxContext.IUCNonSystemGenAnswer
                    //        .Where(x => x.Id == item.Id)
                    //        .Include(x => x.ListQuestionAnswer)
                    //        .Select(x => x.ListQuestionAnswer)
                    //        .AsNoTracking()
                    //        .FirstOrDefault();
                    //        if (iucNonSys1 != null)
                    //        {
                    //            item.ListQuestionAnswer = iucNonSys1;
                    //        }
                    //    }
                    //}
                    //if (roundSetCheck.ListIUCNonSystemGen2 != null && roundSetCheck.ListIUCNonSystemGen2.Count > 0)
                    //{
                    //    foreach (var item in roundSetCheck.ListIUCNonSystemGen2)
                    //    {
                    //        var iucNonSys2 = _soxContext.IUCNonSystemGenAnswer
                    //        .Where(x => x.Id == item.Id)
                    //        .Include(x => x.ListQuestionAnswer)
                    //        .Select(x => x.ListQuestionAnswer)
                    //        .AsNoTracking()
                    //        .FirstOrDefault();
                    //        if (iucNonSys2 != null)
                    //        {
                    //            item.ListQuestionAnswer = iucNonSys2;
                    //        }
                    //    }
                    //}
                    //if (roundSetCheck.ListIUCNonSystemGen3 != null && roundSetCheck.ListIUCNonSystemGen3.Count > 0)
                    //{
                    //    foreach (var item in roundSetCheck.ListIUCNonSystemGen3)
                    //    {
                    //        var iucNonSys3 = _soxContext.IUCNonSystemGenAnswer
                    //        .Where(x => x.Id == item.Id)
                    //        .Include(x => x.ListQuestionAnswer)
                    //        .Select(x => x.ListQuestionAnswer)
                    //        .AsNoTracking()
                    //        .FirstOrDefault();
                    //        if (iucNonSys3 != null)
                    //        {
                    //            item.ListQuestionAnswer = iucNonSys3;
                    //        }
                    //    }
                    //}

                    //if (roundSetCheck.ListIUCSystemGen1 != null && roundSetCheck.ListIUCSystemGen1.Count > 0)
                    //{
                    //    foreach (var item in roundSetCheck.ListIUCSystemGen1)
                    //    {
                    //        var iucSys1 = _soxContext.IUCSystemGenAnswer
                    //        .Where(x => x.Id == item.Id)
                    //        .Include(x => x.ListQuestionAnswer)
                    //        .Select(x => x.ListQuestionAnswer)
                    //        .AsNoTracking()
                    //        .FirstOrDefault();
                    //        if (iucSys1 != null)
                    //        {
                    //            item.ListQuestionAnswer = iucSys1;
                    //        }
                    //    }
                    //}
                    //if (roundSetCheck.ListIUCSystemGen2 != null && roundSetCheck.ListIUCSystemGen2.Count > 0)
                    //{
                    //    foreach (var item in roundSetCheck.ListIUCSystemGen2)
                    //    {
                    //        var iucSys2 = _soxContext.IUCSystemGenAnswer
                    //        .Where(x => x.Id == item.Id)
                    //        .Include(x => x.ListQuestionAnswer)
                    //        .Select(x => x.ListQuestionAnswer)
                    //        .AsNoTracking()
                    //        .FirstOrDefault();
                    //        if (iucSys2 != null)
                    //        {
                    //            item.ListQuestionAnswer = iucSys2;
                    //        }
                    //    }
                    //}
                    //if (roundSetCheck.ListIUCSystemGen3 != null && roundSetCheck.ListIUCSystemGen3.Count > 0)
                    //{
                    //    foreach (var item in roundSetCheck.ListIUCSystemGen3)
                    //    {
                    //        var iucSys3 = _soxContext.IUCSystemGenAnswer
                    //        .Where(x => x.Id == item.Id)
                    //        .Include(x => x.ListQuestionAnswer)
                    //        .Select(x => x.ListQuestionAnswer)
                    //        .AsNoTracking()
                    //        .FirstOrDefault();
                    //        if (iucSys3 != null)
                    //        {
                    //            item.ListQuestionAnswer = iucSys3;
                    //        }
                    //    }
                    //}

                    return Ok(roundSetCheck);

                }
               

            }
            catch (Exception ex)
            {
                FileLog.Write($"Error GetSpecificRoundSet {ex}", "ErrorGetSpecificRoundSet");
                AdminService adminService = new AdminService(_config);
              //  adminService.SendAlert(true, true, ex.ToString(), "GetSpecificRoundSet");
                return BadRequest("Oops something when wrong!." + ex.ToString());
            }

            return BadRequest($"Bad request..");
        }

        //[AllowAnonymous]
        [HttpGet("roundset2/{uniqueid}")]
        public IActionResult GetSpecificRoundSet2(string uniqueid)
        {
            //test link = https://localhost:44344/SampleSelection/download/Draft-TestRound-20200408_183306.xlsx

            try
            {

                var roundSetCheck = _soxContext.QuestionnaireRoundSet
                    .OrderByDescending(x => x.Id)
                    .Include(x => x.ListUserInputRound1)
                    .Include(x => x.ListUserInputRound2)
                    .Include(x => x.ListUserInputRound3)
                    .Include(x => x.ListUniqueNotes)
                    .Include(x => x.ListRoundItem)
                    .Include(x => x.ListIUCSystemGen1)
                    .Include(x => x.ListIUCSystemGen2)
                    .Include(x => x.ListIUCSystemGen3)
                    .Include(x => x.ListIUCNonSystemGen1)
                    .Include(x => x.ListIUCNonSystemGen2)
                    .Include(x => x.ListIUCNonSystemGen3)
                    .Include(x => x.ListHeaderNote)
                    .Include(x => x.sampleSel1)
                    //    .ThenInclude(x => x.ListTestRound)
                    .Include(x => x.sampleSel2)
                    //    .ThenInclude(x => x.ListTestRound)
                    .Include(x => x.sampleSel3)
                    //    .ThenInclude(x => x.ListTestRound)
                    .Where(id => id.UniqueId == uniqueid)
                    .AsNoTracking()
                    .FirstOrDefault();
                if (roundSetCheck != null)
                {
                    //Get sample selection test round 1
                    if (roundSetCheck.sampleSel1 != null)
                    {
                        var sampleSel1 = _soxContext.SampleSelection
                            .Where(x => x.Id == roundSetCheck.sampleSel1.Id)
                            .Include(x => x.ListTestRound)
                            .Select(x => x.ListTestRound)
                            .AsNoTracking()
                            .FirstOrDefault();
                        if (sampleSel1 != null)
                        {
                            roundSetCheck.sampleSel1.ListTestRound = sampleSel1;
                        }
                    }
                    //Get sample selection test round 2
                    if (roundSetCheck.sampleSel2 != null)
                    {
                        var sampleSel2 = _soxContext.SampleSelection
                            .Where(x => x.Id == roundSetCheck.sampleSel2.Id)
                            .Include(x => x.ListTestRound)
                            .Select(x => x.ListTestRound)
                            .AsNoTracking()
                            .FirstOrDefault();
                        if (sampleSel2 != null)
                        {
                            roundSetCheck.sampleSel2.ListTestRound = sampleSel2;
                        }
                    }
                    //Get sample selection test round 3
                    if (roundSetCheck.sampleSel3 != null)
                    {
                        var sampleSel3 = _soxContext.SampleSelection
                            .Where(x => x.Id == roundSetCheck.sampleSel3.Id)
                            .Include(x => x.ListTestRound)
                            .Select(x => x.ListTestRound)
                            .AsNoTracking()
                            .FirstOrDefault();
                        if (sampleSel3 != null)
                        {
                            roundSetCheck.sampleSel3.ListTestRound = sampleSel3;
                        }
                    }


                    if (roundSetCheck.ListIUCNonSystemGen1 != null && roundSetCheck.ListIUCNonSystemGen1.Count > 0)
                    {
                        foreach (var item in roundSetCheck.ListIUCNonSystemGen1)
                        {
                            var iucNonSys1 = _soxContext.IUCNonSystemGenAnswer
                            .Where(x => x.Id == item.Id)
                            .Include(x => x.ListQuestionAnswer)
                            .Select(x => x.ListQuestionAnswer)
                            .AsNoTracking()
                            .FirstOrDefault();
                            if (iucNonSys1 != null)
                            {
                                item.ListQuestionAnswer = iucNonSys1;
                            }
                        }
                    }
                    if (roundSetCheck.ListIUCNonSystemGen2 != null && roundSetCheck.ListIUCNonSystemGen2.Count > 0)
                    {
                        foreach (var item in roundSetCheck.ListIUCNonSystemGen2)
                        {
                            var iucNonSys2 = _soxContext.IUCNonSystemGenAnswer
                            .Where(x => x.Id == item.Id)
                            .Include(x => x.ListQuestionAnswer)
                            .Select(x => x.ListQuestionAnswer)
                            .AsNoTracking()
                            .FirstOrDefault();
                            if (iucNonSys2 != null)
                            {
                                item.ListQuestionAnswer = iucNonSys2;
                            }
                        }
                    }
                    if (roundSetCheck.ListIUCNonSystemGen3 != null && roundSetCheck.ListIUCNonSystemGen3.Count > 0)
                    {
                        foreach (var item in roundSetCheck.ListIUCNonSystemGen3)
                        {
                            var iucNonSys3 = _soxContext.IUCNonSystemGenAnswer
                            .Where(x => x.Id == item.Id)
                            .Include(x => x.ListQuestionAnswer)
                            .Select(x => x.ListQuestionAnswer)
                            .AsNoTracking()
                            .FirstOrDefault();
                            if (iucNonSys3 != null)
                            {
                                item.ListQuestionAnswer = iucNonSys3;
                            }
                        }
                    }

                    if (roundSetCheck.ListIUCSystemGen1 != null && roundSetCheck.ListIUCSystemGen1.Count > 0)
                    {
                        foreach (var item in roundSetCheck.ListIUCSystemGen1)
                        {
                            var iucSys1 = _soxContext.IUCSystemGenAnswer
                            .Where(x => x.Id == item.Id)
                            .Include(x => x.ListQuestionAnswer)
                            .Select(x => x.ListQuestionAnswer)
                            .AsNoTracking()
                            .FirstOrDefault();
                            if (iucSys1 != null)
                            {
                                item.ListQuestionAnswer = iucSys1;
                            }
                        }
                    }
                    if (roundSetCheck.ListIUCSystemGen2 != null && roundSetCheck.ListIUCSystemGen2.Count > 0)
                    {
                        foreach (var item in roundSetCheck.ListIUCSystemGen2)
                        {
                            var iucSys2 = _soxContext.IUCSystemGenAnswer
                            .Where(x => x.Id == item.Id)
                            .Include(x => x.ListQuestionAnswer)
                            .Select(x => x.ListQuestionAnswer)
                            .AsNoTracking()
                            .FirstOrDefault();
                            if (iucSys2 != null)
                            {
                                item.ListQuestionAnswer = iucSys2;
                            }
                        }
                    }
                    if (roundSetCheck.ListIUCSystemGen3 != null && roundSetCheck.ListIUCSystemGen3.Count > 0)
                    {
                        foreach (var item in roundSetCheck.ListIUCSystemGen3)
                        {
                            var iucSys3 = _soxContext.IUCSystemGenAnswer
                            .Where(x => x.Id == item.Id)
                            .Include(x => x.ListQuestionAnswer)
                            .Select(x => x.ListQuestionAnswer)
                            .AsNoTracking()
                            .FirstOrDefault();
                            if (iucSys3 != null)
                            {
                                item.ListQuestionAnswer = iucSys3;
                            }
                        }
                    }

                    return Ok(roundSetCheck);

                }


            }
            catch (Exception ex)
            {
                FileLog.Write($"Error GetSpecificRoundSet2 {ex}", "ErrorGetSpecificRoundSet2");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetSpecificRoundSet2");
                return BadRequest("Oops something when wrong!." + ex.ToString());
            }

            return BadRequest($"Bad request..");
        }

        [AllowAnonymous]
        [HttpDelete("roundset/{uniqueid}")]
        public async Task<IActionResult> RemoveSpecificRoundSet(string uniqueid)
        {
            //test link = https://localhost:44344/SampleSelection/download/Draft-TestRound-20200408_183306.xlsx

            try
            {

                var roundSetCheck = _soxContext.QuestionnaireRoundSet
                    .Include(x => x.sampleSel1)
                    .Include(x => x.sampleSel2)
                    .Include(x => x.sampleSel3)
                    .Where(id => id.UniqueId == uniqueid)
                    .AsNoTracking()
                    .FirstOrDefault();
                if (roundSetCheck != null)
                {

                    using (var context = _soxContext.Database.BeginTransaction())
                    {
                        _soxContext.Remove(roundSetCheck);
                        await _soxContext.SaveChangesAsync();
                        context.Commit();
                        return Ok();
                    }
                }
            }
            catch (Exception ex)
            {
                FileLog.Write($"Error GetSpecificRoundSet {ex}", "ErrorGetSpecificRoundSet");
                AdminService adminService = new AdminService(_config);
                //  adminService.SendAlert(true, true, ex.ToString(), "GetSpecificRoundSet");
                return BadRequest("Oops something when wrong!." + ex.ToString());
            }

            return BadRequest($"Bad request..");
        }

        [HttpPost("create")]
        public async Task<IActionResult> CreateQuestionnaireAsync([FromBody] Questionnaire questionnaire)
        {
            bool status = false;
            using (var context = _soxContext.Database.BeginTransaction())
            {
                try
                {
                    Questionnaire _questionnaire = new Questionnaire();
                    _questionnaire = questionnaire;

                    bool isExists = false;
                    var questionnaireCheck = _soxContext.Questionnaire.FirstOrDefault(id => id.PodioItemId == _questionnaire.PodioItemId);
                    if (questionnaireCheck != null)
                    {
                        isExists = true;
                    }

                    if (isExists)
                    {
                        //if exists we update database row
                        _questionnaire.Id = questionnaireCheck.Id;
                        _soxContext.Entry(questionnaireCheck).CurrentValues.SetValues(_questionnaire);
                    }
                    else
                    {
                        //create new item
                        _soxContext.Add(_questionnaire);
                    }

                    await _soxContext.SaveChangesAsync();
                    context.Commit();
                    status = true;
                }
                catch (Exception ex)
                {
                    //Console.WriteLine(ex.ToString());
                    //ErrorLog.Write(ex);
                    await context.RollbackAsync();
                    FileLog.Write($"Error CreateQuestionnaireAsync {ex}", "ErrorCreateQuestionnaireAsync");
                    AdminService adminService = new AdminService(_config);
                    adminService.SendAlert(true, true, ex.ToString(), "CreateQuestionnaireAsync");
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

        [HttpPost("additionalInputs")]//mark
        public async Task<IActionResult> AdditionalInputs([FromBody] QuestionaireAddedInputs Inputs)
        {
            bool status = false;
            using (var context = _soxContext.Database.BeginTransaction())
            {
                try
                {
                    bool isExists = false;
                    var questionnaireCheck = _soxContext.QuestionaireAddedInputs.Where(
                        x => x.ClientName.Equals(Inputs.ClientName) &&
                        x.ClientCode.Equals(Inputs.ClientCode) &&
                        x.ClientItemId.Equals(Inputs.ClientItemId) &&
                        x.roundset.Equals(Inputs.roundset)).FirstOrDefault();

                    QuestionaireAddedInputs _Inputs = new QuestionaireAddedInputs();
                    _Inputs = Inputs;

                    if (questionnaireCheck != null)
                    {
                        isExists = true;
                    }

                    if (isExists)
                    {
                        //if exists we update database row
                        _Inputs.Id = questionnaireCheck.Id;
                        _soxContext.Entry(questionnaireCheck).CurrentValues.SetValues(_Inputs);
                    }
                    else
                    {
                        //create new item
                        _soxContext.Add(_Inputs);
                    }

                    await _soxContext.SaveChangesAsync();
                    context.Commit();
                    status = true;
                }
                catch (Exception ex)
                {
                    //Console.WriteLine(ex.ToString());
                    //ErrorLog.Write(ex);
                    await context.RollbackAsync();
                    FileLog.Write($"Error Save additional Inputs {ex}", "Error Save additional Inputs");
                    AdminService adminService = new AdminService(_config);
                    adminService.SendAlert(true, true, ex.ToString(), "Error Save additional Inputs");
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

        [HttpPost("createDraftWorkpaper")]
        public async Task<IActionResult> CreateDraftWorkpaper([FromBody] QuestionnaireTesterSet questionnaireTesterSet)
        {
            using (var context = _soxContext.Database.BeginTransaction())
            {

            
                try
                {
                
                
                    if(questionnaireTesterSet.UniqueId == string.Empty)
                    {
                        Guid guid = Guid.NewGuid();
                        questionnaireTesterSet.UniqueId = guid.ToString();
                    }
                        
                    List<QuestionnaireUserAnswer> listUserInput = new List<QuestionnaireUserAnswer>();
                    FormatService formatService = new FormatService();

                    var workpaperStatus = _soxContext.WorkpaperStatus.Where(x => x.Index.Equals(questionnaireTesterSet.WorkpaperStatus.Index)).AsNoTracking().FirstOrDefault();
                    if(workpaperStatus != null)
                    {
                        questionnaireTesterSet.WorkpaperStatus = workpaperStatus;
                        

                        var rcm = _soxContext.Rcm.Where(rcmItem => rcmItem.PodioItemId.Equals(questionnaireTesterSet.RcmItemId) && rcmItem.Status.Equals("Active")).AsNoTracking().FirstOrDefault();
                        if (rcm != null)
                        {
                            questionnaireTesterSet.Rcm = rcm;
                            var listQuestion = _soxContext.QuestionnaireQuestion
                                        .Where(x =>
                                            x.ClientName.Equals(rcm.ClientName) &&
                                            x.ControlName.Equals(rcm.ControlId))
                                        .Include(x => x.Options)
                                        .AsNoTracking()
                                        .ToList();

                            
                            List<string> TestingAttHeader = new List<string>();

                            //setup testing attributes question
                            List<RoundItem2> ListRoundItem2 = new List<RoundItem2>();
                            RoundItem2 roundItem = new RoundItem2();
                            roundItem.ListRoundQA = new List<RoundQA2>();
                            int posQAItem = 0;
                            List<RoundQA2> listRoundQA = new List<RoundQA2>();
                            string appId = string.Empty;
                            //Set question
                            if (listQuestion != null && listQuestion.Any()) //if listquestion has value
                            {
                                
                                //create default user input from list of question
                                var sortedListQuestion = listQuestion.OrderBy( x=> x.Position ).ToList();
                                foreach (var itemQuestion in sortedListQuestion)
                                {
                                    appId = itemQuestion.AppId;
                                    if (itemQuestion.QuestionString.ToLower().Contains("(rt)"))
                                    {
                                        //TestingAttHeader.Add(itemQuestion.QuestionString);
                                        RoundQA2 tempRoundQA = new RoundQA2();
                                        tempRoundQA.Position = posQAItem;
                                        tempRoundQA.Question = itemQuestion.QuestionString;
                                        tempRoundQA.DtEndRequire = itemQuestion.DtEndRequire;
                                        tempRoundQA.Type = itemQuestion.Type;
                                        tempRoundQA.Options = itemQuestion.Options;
                                        listRoundQA.Add(tempRoundQA);
                                        posQAItem++;
                                    }

                                    QuestionnaireUserAnswer userInput = new QuestionnaireUserAnswer();
                                    userInput.StrAnswer = string.Empty;
                                    userInput.StrAnswer2 = string.Empty;
                                    userInput.StrQuestion = itemQuestion.QuestionString.ToLower().Contains("(ro)") ? itemQuestion.QuestionString.Replace("(RO)", string.Empty) : itemQuestion.QuestionString;
                                    userInput.IsDisabled = itemQuestion.QuestionString.ToLower().Contains("(ro)") ? true : false;
                                    userInput.Description = itemQuestion.Description;
                                    userInput.Position = itemQuestion.Position;
                                    userInput.AppId = itemQuestion.AppId;
                                    userInput.FieldId = itemQuestion.FieldId;
                                    userInput.ItemId = 0;
                                    userInput.Type = itemQuestion.Type;
                                    userInput.DtEndRequire = itemQuestion.DtEndRequire;
                                    userInput.CreatedOn = DateTime.Now;
                                    userInput.UpdatedOn = DateTime.Now; 
                                    List<QuestionnaireOption> listOption = new List<QuestionnaireOption>();
                                    foreach (var itemOptions in itemQuestion.Options)
                                    {
                                        QuestionnaireOption questionnaireOption = new QuestionnaireOption();
                                        questionnaireOption.OptionName = itemOptions.OptionName;
                                        questionnaireOption.OptionId = itemOptions.OptionId;
                                        questionnaireOption.AppId = itemOptions.AppId;
                                        questionnaireOption.CreatedOn = DateTime.Now;
                                        listOption.Add(questionnaireOption);
                                    }


                                    #region populate answer base from RCM
                                    if(rcm != null && itemQuestion.Type != "image")
                                    {
                                        switch (itemQuestion.QuestionString.ToLower())
                                        {
                                            case string s when s.Contains("1. what is the client name?"):
                                                if (rcm.ClientName != null && rcm.ClientName != string.Empty)
                                                {
                                                    userInput.StrAnswer = formatService.ReplaceTagHtmlParagraph(rcm.ClientName, false);
                                                }
                                                break;
                                            case string s when s.Contains("what is the purpose of this control?"):
                                                if (rcm.SpecificRisk != null && rcm.SpecificRisk != string.Empty)
                                                {
                                                    userInput.StrAnswer = formatService.ReplaceTagHtmlParagraph(rcm.SpecificRisk, false);
                                                }
                                                break;
                                            case string s when s.Contains("what is the control id?"):
                                                if (rcm.ControlId != null && rcm.ControlId != string.Empty)
                                                {
                                                    userInput.StrAnswer = formatService.ReplaceTagHtmlParagraph(rcm.ControlId, false);
                                                }
                                                break;
                                            case string s when s.Contains("what is the control activity?"):
                                                if (rcm.ControlActivityFy19 != null && rcm.ControlActivityFy19 != string.Empty)
                                                {
                                                    userInput.StrAnswer = formatService.ReplaceTagHtmlParagraph(rcm.ControlActivityFy19, false);
                                                }
                                                break;
                                            case string s when s.Contains("who is the control owner?"):
                                                if (rcm.ControlOwner != null && rcm.ControlOwner != string.Empty)
                                                {
                                                    userInput.StrAnswer = formatService.ReplaceTagHtmlParagraph(rcm.ControlOwner, false);
                                                }
                                                break;
                                            case string s when
                                                s.Contains("what are the procedures to test this control") ||
                                                s.Contains("what are the procedures to test control"):
                                                if (rcm.TestProc != null && rcm.TestProc != string.Empty)
                                                {
                                                    userInput.StrAnswer = formatService.ReplaceTagHtmlParagraph(rcm.TestProc, false);
                                                }
                                                break;
                                            case string s when s.Contains("when was the control first put in place?"):
                                                if (rcm.ControlPlaceDate != null && rcm.ControlPlaceDate != string.Empty)
                                                {
                                                    userInput.StrAnswer = formatService.ReplaceTagHtmlParagraph(rcm.ControlPlaceDate, false);
                                                }
                                                break;
                                            case string s when s.Contains("what is the date range you are testing?"):
                                                if (rcm.TestingPeriod != null && rcm.TestingPeriod != string.Empty)
                                                {
                                                    userInput.StrAnswer = formatService.ReplaceTagHtmlParagraph(rcm.TestingPeriod, false);
                                                }
                                                break;
                                            case string s when s.Contains("what is the level of risk for the control?"):
                                                if (rcm.RiskLvl != null && rcm.RiskLvl != string.Empty)
                                                {
                                                    userInput.StrAnswer = formatService.ReplaceTagHtmlParagraph(rcm.RiskLvl, false);
                                                }
                                                break;
                                            case string s when s.Contains("how often does this control happen?"):
                                                if (rcm.ControlFrequency != null && rcm.ControlFrequency != string.Empty)
                                                {
                                                    userInput.StrAnswer = formatService.ReplaceTagHtmlParagraph(rcm.ControlFrequency, false);
                                                }
                                                break;
                                        }
                                        userInput.StrDefaultAnswer = userInput.StrAnswer;
                                    }
                                    #endregion


                                    userInput.Options = listOption;
                                    listUserInput.Add(userInput);

                                }
                                
                            }




                            //Set IPE notes
                            List<IPENote> listIPE;
                            if (rcm.ControlId.ToLower().Contains("itgc"))
                            {
                                listIPE = new List<IPENote>
                                {
                                    new IPENote
                                    {
                                        Name = "IPE Note",
                                        Note = "IPE Note",
                                        Description = string.Empty,
                                        Display = false
                                    }
                                };

                            }
                            else
                            {
                                
                                listIPE = new List<IPENote>
                                {
                                    new IPENote
                                    {
                                        Name = "IPE Walkthrough",
                                        Note = "IPE/Validation of PBC Reports (Completeness/Accuracy) Walkthrough",
                                        Description = string.Empty,
                                        Display = false
                                    },
                                    new IPENote
                                    {
                                        Name = "IPE Round 1",
                                        Note = "IPE/Validation of PBC Reports (Completeness/Accuracy) Round 1",
                                        Description = string.Empty,
                                        Display = false
                                    },
                                    new IPENote
                                    {
                                        Name = "IPE Round 2",
                                        Note = "IPE/Validation of PBC Reports (Completeness/Accuracy) Round 2",
                                        Description = string.Empty,
                                        Display = false
                                    },
                                        new IPENote
                                    {
                                        Name = "IPE Round 3",
                                        Note = "IPE/Validation of PBC Reports (Completeness/Accuracy) Round 3",
                                        Description = string.Empty,
                                        Display = false
                                    }
                                };
                            }



                            roundItem.RoundName = "Header";
                            roundItem.Position = 0;
                            roundItem.A2Q2Samples = "0";
                            roundItem.AppId = appId;
                            roundItem.CreatedOn = DateTime.UtcNow;
                            roundItem.CreatedBy = questionnaireTesterSet.AddedBy;
                            foreach (var itemRoundHeader in listRoundQA)
                            {
                                RoundQA2 roundQA = new RoundQA2();
                                roundQA.Question = itemRoundHeader.Question;
                                roundQA.Position = itemRoundHeader.Position;
                                roundQA.DtEndRequire = itemRoundHeader.DtEndRequire;
                                roundQA.Type = itemRoundHeader.Type;
                                //roundQA.Options = itemRoundHeader.Options;
                                List<QuestionnaireOption> listOption = new List<QuestionnaireOption>();
                                foreach (var itemOptions in itemRoundHeader.Options)
                                {
                                    QuestionnaireOption questionnaireOption = new QuestionnaireOption();
                                    questionnaireOption.OptionName = itemOptions.OptionName;
                                    questionnaireOption.OptionId = itemOptions.OptionId;
                                    questionnaireOption.AppId = itemOptions.AppId;
                                    questionnaireOption.CreatedOn = DateTime.Now;
                                    listOption.Add(questionnaireOption);
                                }

                                roundQA.Options = listOption;
                                roundItem.ListRoundQA.Add(roundQA);
                                posQAItem++;
                            }

                            ListRoundItem2.Add(roundItem);

                            questionnaireTesterSet.ListRoundItem2 = ListRoundItem2;


                            questionnaireTesterSet.ListIPENote = listIPE;
                            questionnaireTesterSet.Position = 0;
                            questionnaireTesterSet.ListUserInputRound = listUserInput.OrderBy(order => order.Position).ToList();
                            questionnaireTesterSet.CreatedOn = DateTime.UtcNow;

                        }

                        _soxContext.Entry(questionnaireTesterSet.WorkpaperStatus).State = EntityState.Unchanged;
                        _soxContext.Entry(questionnaireTesterSet.Rcm).State = EntityState.Unchanged;
                        _soxContext.Add(questionnaireTesterSet);

                        await _soxContext.SaveChangesAsync();

                        


                        context.Commit();
                        return Ok(questionnaireTesterSet);
                    }
                
                    
                }
                catch (Exception ex)
                {
                    FileLog.Write($"Error CreateDraftWorkpaperAsync {ex}", "ErrorCreateDraftWorkpaperAsync");
                    AdminService adminService = new AdminService(_config);
                    adminService.SendAlert(true, true, ex.ToString(), "CreateDraftWorkpaperAsync");
                    context.Rollback(); //rollback added data 
                    return BadRequest();
                }
            }
            return NoContent();
           
        }

        [AllowAnonymous]
        [HttpPost("getDraftWorkpaper")]
        public IActionResult GetDraftWorkpaper([FromBody] QuestionnaireDraftId draftId)
        {

            try
            {
                
                var listDraftTester = _soxContext.QuestionnaireTesterSet
                .Where(x => 
                    x.UniqueId.Equals(draftId.UniqueId) && 
                    x.RoundName.Equals(draftId.RoundName))
                .Include(sysGen => sysGen.ListIUCSystemGen)
                .Include(nonSysGen => nonSysGen.ListIUCNonSystemGen)
                .Include(notes => notes.ListUniqueNotes)       
                .Include(ipe => ipe.ListIPENote)
                .Include(header => header.ListHeaderNote2)
                .Include(stat => stat.WorkpaperStatus)
                .Include(genNote => genNote.GeneralNote)
                .Include(samp => samp.SampleSel)
                .Include(userInput => userInput.ListUserInputRound)
                    .ThenInclude(inner => inner.Options)
                .Include(round => round.ListRoundItem2)
                    .ThenInclude(qa => qa.ListRoundQA)
                        .ThenInclude(opt => opt.Options)
                .AsNoTracking()
                .ToList();
                if (listDraftTester != null && listDraftTester.Any())
                {
                    foreach (var item in listDraftTester)
                    {
                        if (item.SampleSel != null)
                        {

                            var resTestround1 = _soxContext.TestRounds.FromSqlRaw($"CALL `sox`.`sp_get_testround`({item.SampleSel.Id});").AsNoTracking().ToList();
                            if (resTestround1.Count > 0)
                            {
                                item.SampleSel.ListTestRound = resTestround1;
                            }
                        }

                        if (item.ListUserInputRound != null && item.ListUserInputRound.Any())
                        {
                            item.ListUserInputRound = item.ListUserInputRound.OrderBy(sort => sort.Position).ToList();
                        }


                        //var resQuestionnaireUserAnswer = _soxContext.QuestionnaireUserAnswer.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_user_answer`({item.Id});").AsNoTracking().ToList();
                        //if (resQuestionnaireUserAnswer.Any())
                        //{
                        //    item.ListUserInputRound = resQuestionnaireUserAnswer
                        //        .OrderBy(x => x.Position)
                        //        .Where(x => x.RoundName.Equals(item.RoundName))
                        //        .ToList();

                            //}

                            //var resNotesItem = _soxContext.NotesItem.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_notesitem`({item.Id});").AsNoTracking().ToList();
                            //if (resNotesItem.Any())
                            //{
                            //    item.ListUniqueNotes = resNotesItem
                            //        .OrderBy(x => x.Id)
                            //        .ToList();
                            //}

                            //var resListRound = _soxContext.RoundItem.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_rounditem`({item.Id});").AsNoTracking().ToList();
                            //if (resListRound.Any())
                            //{
                            //    item.ListRoundItem = resListRound
                            //        .OrderBy(x => x.Id)
                            //        .ToList();
                            //}

                            #region IUC System Gen

                            //var resIUCSystemGenAnswer1 = _soxContext.IUCSystemGenAnswer.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_iuc_system_gen`({item.Id},'{item.RoundName}');").ToList();
                            //if (resIUCSystemGenAnswer1.Any())
                            //{
                            //    item.ListIUCSystemGen = resIUCSystemGenAnswer1.ToList();
                            //    if (item.ListIUCSystemGen.Count > 0)
                            //    {
                            //        foreach (var item in item.ListIUCSystemGen)
                            //        {
                            //            var resIUCSystemGenQuestion1 = _soxContext.IUCQuestionUserAnswer.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_iuc_user_input`({item.Id},'1');").ToList();
                            //            if (resIUCSystemGenQuestion1.Count > 0)
                            //            {
                            //                item.ListQuestionAnswer = resIUCSystemGenQuestion1;
                            //            }
                            //        }
                            //    }
                            //}

                            //var resIUCSystemGenAnswer2 = _soxContext.IUCSystemGenAnswer.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_iuc_system_gen`({roundSetCheck.Id},'Round 2');").ToList();
                            //if (resIUCSystemGenAnswer2.Any())
                            //{
                            //    roundSetCheck.ListIUCSystemGen2 = resIUCSystemGenAnswer2.ToList();
                            //    if (roundSetCheck.ListIUCSystemGen2.Count > 0)
                            //    {
                            //        foreach (var item in roundSetCheck.ListIUCSystemGen2)
                            //        {
                            //            var resIUCSystemGenQuestion2 = _soxContext.IUCQuestionUserAnswer.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_iuc_user_input`({item.Id},'1');").ToList();
                            //            if (resIUCSystemGenQuestion2.Count > 0)
                            //            {
                            //                item.ListQuestionAnswer = resIUCSystemGenQuestion2;
                            //            }
                            //        }
                            //    }
                            //}

                            //var resIUCSystemGenAnswer3 = _soxContext.IUCSystemGenAnswer.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_iuc_system_gen`({roundSetCheck.Id},'Round 3');").ToList();
                            //if (resIUCSystemGenAnswer3.Any())
                            //{
                            //    roundSetCheck.ListIUCSystemGen3 = resIUCSystemGenAnswer3.ToList();
                            //    if (roundSetCheck.ListIUCSystemGen3.Count > 0)
                            //    {
                            //        foreach (var item in roundSetCheck.ListIUCSystemGen3)
                            //        {
                            //            var resIUCSystemGenQuestion3 = _soxContext.IUCQuestionUserAnswer.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_iuc_user_input`({item.Id},'1');").ToList();
                            //            if (resIUCSystemGenQuestion3.Count > 0)
                            //            {
                            //                item.ListQuestionAnswer = resIUCSystemGenQuestion3;
                            //            }
                            //        }
                            //    }
                            //}

                            #endregion

                            #region IUC Non System Gen

                            //var resIUCNonSystemGenAnswer1 = _soxContext.IUCNonSystemGenAnswer.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_iuc_non_system_gen`({roundSetCheck.Id},'Round 1');").ToList();
                            //if (resIUCNonSystemGenAnswer1.Any())
                            //{
                            //    roundSetCheck.ListIUCNonSystemGen1 = resIUCNonSystemGenAnswer1.ToList();
                            //    if (roundSetCheck.ListIUCNonSystemGen1.Count > 0)
                            //    {
                            //        foreach (var item in roundSetCheck.ListIUCNonSystemGen1)
                            //        {
                            //            var resIUCNonSystemGenQuestion1 = _soxContext.IUCQuestionUserAnswer.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_iuc_user_input`({item.Id},'2');").ToList();
                            //            if (resIUCNonSystemGenQuestion1.Count > 0)
                            //            {
                            //                item.ListQuestionAnswer = resIUCNonSystemGenQuestion1;
                            //            }
                            //        }
                            //    }
                            //}

                            //var resIUCNonSystemGenAnswer2 = _soxContext.IUCNonSystemGenAnswer.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_iuc_non_system_gen`({roundSetCheck.Id},'Round 2');").ToList();
                            //if (resIUCNonSystemGenAnswer2.Any())
                            //{
                            //    roundSetCheck.ListIUCNonSystemGen2 = resIUCNonSystemGenAnswer2.ToList();
                            //    if (roundSetCheck.ListIUCNonSystemGen2.Count > 0)
                            //    {
                            //        foreach (var item in roundSetCheck.ListIUCNonSystemGen2)
                            //        {
                            //            var resIUCNonSystemGenQuestion2 = _soxContext.IUCQuestionUserAnswer.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_iuc_user_input`({item.Id},'2');").ToList();
                            //            if (resIUCNonSystemGenQuestion2.Count > 0)
                            //            {
                            //                item.ListQuestionAnswer = resIUCNonSystemGenQuestion2;
                            //            }
                            //        }

                            //    }
                            //}

                            //var resIUCNonSystemGenAnswer3 = _soxContext.IUCNonSystemGenAnswer.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_iuc_non_system_gen`({roundSetCheck.Id},'Round 3');").ToList();
                            //if (resIUCNonSystemGenAnswer3.Any())
                            //{
                            //    roundSetCheck.ListIUCNonSystemGen3 = resIUCNonSystemGenAnswer3.ToList();
                            //    if (roundSetCheck.ListIUCNonSystemGen3.Count > 0)
                            //    {
                            //        foreach (var item in roundSetCheck.ListIUCNonSystemGen3)
                            //        {
                            //            var resIUCNonSystemGenQuestion3 = _soxContext.IUCQuestionUserAnswer.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_iuc_user_input`({item.Id},'2');").ToList();
                            //            if (resIUCNonSystemGenQuestion3.Count > 0)
                            //            {
                            //                item.ListQuestionAnswer = resIUCNonSystemGenQuestion3;
                            //            }
                            //        }

                            //    }
                            //}


                            #endregion

                            //var resHeaderNote = _soxContext.HeaderNote.FromSqlRaw($"CALL `sox`.`sp_get_headernote`({roundSetCheck.Id});").ToList();
                            //if (resHeaderNote.Any())
                            //{
                            //    roundSetCheck.ListHeaderNote = resHeaderNote.ToList();
                            //}
                    }
                    //Get sample selection test round 1

                }

                return Ok(listDraftTester);

              
            }
            catch (Exception ex)
            {
                FileLog.Write($"Error GetDraftWorkpaper {ex}", "ErrorGetDraftWorkpaper");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetDraftWorkpaper");
                return BadRequest();
            }

        }

        [AllowAnonymous]
        [HttpPost("saveDraftWorkpaper")]
        public async Task<IActionResult> SaveDraftWorkpaper([FromBody] List<QuestionnaireTesterSet> listQuestionnaireTesterSet)
        {

            FormatService formatService = new FormatService();
            QuestionnaireTesterSet questionnaireTester = new QuestionnaireTesterSet();

            //Update
            using (var context = _soxContext.Database.BeginTransaction())
            {

                try
                {
                    

                    if (listQuestionnaireTesterSet != null && listQuestionnaireTesterSet.Any())
                    {
                        
                        int index = listQuestionnaireTesterSet.Count() - 1; // get total count and save only the last index

                        //Get last item and save
                        questionnaireTester.Id = listQuestionnaireTesterSet[index].Id;
                        questionnaireTester.UniqueId = listQuestionnaireTesterSet[index].UniqueId;
                        questionnaireTester.ListUserInputRound = listQuestionnaireTesterSet[index].ListUserInputRound;
                        questionnaireTester.ListIUCSystemGen = listQuestionnaireTesterSet[index].ListIUCSystemGen;
                        questionnaireTester.ListIUCNonSystemGen = listQuestionnaireTesterSet[index].ListIUCNonSystemGen;
                        questionnaireTester.ListUniqueNotes = listQuestionnaireTesterSet[index].ListUniqueNotes;
                        questionnaireTester.ListRoundItem2 = listQuestionnaireTesterSet[index].ListRoundItem2;
                        questionnaireTester.GeneralNote = listQuestionnaireTesterSet[index].GeneralNote;
                        questionnaireTester.ListIPENote = listQuestionnaireTesterSet[index].ListIPENote;
                        questionnaireTester.SampleSel = listQuestionnaireTesterSet[index].SampleSel;
                        questionnaireTester.WorkpaperStatus = listQuestionnaireTesterSet[index].WorkpaperStatus;
                        questionnaireTester.RoundName = listQuestionnaireTesterSet[index].RoundName;
                        questionnaireTester.AddedBy = listQuestionnaireTesterSet[index].AddedBy;
                        questionnaireTester.DraftNum = listQuestionnaireTesterSet[index].DraftNum;
                        questionnaireTester.RcmItemId = listQuestionnaireTesterSet[index].RcmItemId;
                        questionnaireTester.Rcm = listQuestionnaireTesterSet[index].Rcm;
                        questionnaireTester.UserAction = listQuestionnaireTesterSet[index].UserAction;
                        questionnaireTester.Position = listQuestionnaireTesterSet[index].Position;
                        questionnaireTester.CreatedOn = listQuestionnaireTesterSet[index].CreatedOn;

                        //check if exists, then update
                        var checkQuestionnaireTester = _soxContext.QuestionnaireTesterSet.Where(id => id.Id.Equals(questionnaireTester.Id)).AsNoTracking().FirstOrDefault();
                        if (checkQuestionnaireTester != null)
                        {

                            //update reference in QuestionnaireTesterSet
                            foreach (var item in questionnaireTester.ListUserInputRound)
                            {
                                if (item.IsDisabled) // if read only and is updated by user, we set back the value to default
                                    item.StrAnswer = item.StrDefaultAnswer;

                                _soxContext.Entry(item).State = item.Id == 0 ? EntityState.Added : EntityState.Modified;
                            }

                            foreach (var item in questionnaireTester.ListRoundItem2)
                            {
                                _soxContext.Entry(item).State = item.Id == 0 ? EntityState.Added : EntityState.Modified;

                                foreach (var itemRoundQA in item.ListRoundQA)
                                {
                                    _soxContext.Entry(itemRoundQA).State = itemRoundQA.Id == 0 ? EntityState.Added : EntityState.Modified;
                                }
                            }

                            if (questionnaireTester.ListIPENote != null)
                            {
                                foreach (var item in questionnaireTester.ListIPENote)
                                {
                                    _soxContext.Entry(item).State = item.Id == 0 ? EntityState.Added : EntityState.Modified;
                                }
                            }

                            if (questionnaireTester.ListHeaderNote2 != null)
                            {
                                foreach (var item in questionnaireTester.ListHeaderNote2)
                                {
                                    _soxContext.Entry(item).State = item.Id == 0 ? EntityState.Added : EntityState.Modified;
                                }
                            }

                            if (questionnaireTester.ListUniqueNotes != null)
                            {
                                foreach (var item in questionnaireTester.ListUniqueNotes)
                                {
                                    _soxContext.Entry(item).State = item.Id == 0 ? EntityState.Added : EntityState.Modified;
                                }
                            }

                            if (questionnaireTester.SampleSel != null && questionnaireTester.SampleSel.PodioItemId != 0)
                            {
                                var sampleSelId = _soxContext.SampleSelection.Where(selId => selId.PodioItemId.Equals(questionnaireTester.SampleSel.PodioItemId)).Select(y => y.Id).FirstOrDefault();
                                if (sampleSelId != 0)
                                {
                                    questionnaireTester.SampleSel.Id = sampleSelId;
                                    _soxContext.Entry(questionnaireTester.SampleSel).State = EntityState.Modified;
                                }
                            }
                            else
                            {
                                questionnaireTester.SampleSel = null;
                            }

                            if (questionnaireTester.GeneralNote != null)
                            {
                                _soxContext.Entry(questionnaireTester.GeneralNote).State = questionnaireTester.GeneralNote.Id == 0 ? EntityState.Added : EntityState.Modified;
                            }

                            _soxContext.Update(questionnaireTester);
                            await _soxContext.SaveChangesAsync();


                            #region Old Process
                            ////get previous draft index
                            ////we will use the previous index to copy previous draft
                            //int prevDraftIndex = 0;
                            //List<QuestionnaireUserAnswer> listPrevUserInput = new List<QuestionnaireUserAnswer>();
                            //List<IPENote> listPrevIPE = new List<IPENote>();
                            //List<QuestionnaireTesterSet> listAllDraft = new List<QuestionnaireTesterSet>();
                            //GeneralNote prevGeneralNote = new GeneralNote();
                            //List<RoundItem2> ListRoundItem2 = new List<RoundItem2>();
                            //List<NotesItem2> ListUniqueNotes = new List<NotesItem2>();
                            //List<HeaderNote2> ListHeaderNote2 = new List<HeaderNote2>();
                            //RoundItem2 roundItem = new RoundItem2();
                            //roundItem.ListRoundQA = new List<RoundQA2>();
                            //int posQAItem = 0;
                            //List<RoundQA2> listRoundQA = new List<RoundQA2>();
                            //string appId = string.Empty;
                            //if (questionnaireTesterNew.DraftNum > 1 && (workpaperIndex == 5 || workpaperIndex == 0))
                            //{
                            //    listAllDraft = _soxContext.QuestionnaireTesterSet.Where(prev => prev.UniqueId.Equals(questionnaireTester.UniqueId) && prev.RcmItemId.Equals(questionnaireTester.RcmItemId)).AsNoTracking().ToList();  
                            //    if(listAllDraft.Any())
                            //    {
                            //        prevDraftIndex = listAllDraft.Count() - 2;
                            //        listPrevUserInput = listQuestionnaireTesterSet[prevDraftIndex].ListUserInputRound.ToList();
                            //        listPrevIPE = listQuestionnaireTesterSet[prevDraftIndex].ListIPENote.ToList();
                            //        prevGeneralNote = listQuestionnaireTesterSet[prevDraftIndex].GeneralNote;
                            //        ListRoundItem2 = listQuestionnaireTesterSet[prevDraftIndex].ListRoundItem2.ToList();
                            //        ListUniqueNotes = listQuestionnaireTesterSet[prevDraftIndex].ListUniqueNotes.ToList();
                            //        ListHeaderNote2 = listQuestionnaireTesterSet[prevDraftIndex].ListHeaderNote2.ToList();
                            //    }
                            //}

                            ////check if item already created
                            ////if next sequence is already created then there is no need to create
                            //var checkNextQuestionnaireTester = _soxContext.QuestionnaireTesterSet.Where(y => y.DraftNum.Equals(questionnaireTesterNew.DraftNum) && y.WorkpaperStatus.Equals(questionnaireTesterNew.WorkpaperStatus)).FirstOrDefault();
                            //if(checkNextQuestionnaireTester == null)
                            //{
                            //    questionnaireTesterNew.RoundName = questionnaireTester.RoundName;
                            //    questionnaireTesterNew.AddedBy = questionnaireTester.AddedBy;
                            //    questionnaireTesterNew.RcmItemId = questionnaireTester.RcmItemId;
                            //    questionnaireTesterNew.Rcm = questionnaireTester.Rcm;

                            //    if (questionnaireTesterNew.WorkpaperStatus.Id == 2)
                            //        questionnaireTesterNew.UserAction = "Final";

                            //    var listQuestion = _soxContext.QuestionnaireQuestion
                            //            .Where(x =>
                            //                x.ClientName.Equals(questionnaireTesterNew.Rcm.ClientName) &&
                            //                x.ControlName.Equals(questionnaireTesterNew.Rcm.ControlId))
                            //            .Include(x => x.Options)
                            //            .AsNoTracking()
                            //            .ToList();


                            //    if (listQuestion != null && listQuestion.Any()) //if listquestion has value
                            //    {

                            //        //create default user input from list of question
                            //        var sortedListQuestion = listQuestion.OrderBy(x => x.Position).ToList();
                            //        foreach (var itemQuestion in sortedListQuestion)
                            //        {

                            //            appId = itemQuestion.AppId;
                            //            if (itemQuestion.QuestionString.ToLower().Contains("(rt)"))
                            //            {
                            //                //TestingAttHeader.Add(itemQuestion.QuestionString);
                            //                RoundQA2 tempRoundQA = new RoundQA2();
                            //                tempRoundQA.Position = posQAItem;
                            //                tempRoundQA.Question = itemQuestion.QuestionString;
                            //                tempRoundQA.DtEndRequire = itemQuestion.DtEndRequire;
                            //                tempRoundQA.Type = itemQuestion.Type;
                            //                tempRoundQA.Options = itemQuestion.Options;
                            //                listRoundQA.Add(tempRoundQA);
                            //                posQAItem++;
                            //            }

                            //            QuestionnaireUserAnswer userInput = new QuestionnaireUserAnswer();
                            //            userInput.StrQuestion = itemQuestion.QuestionString;
                            //            userInput.Description = itemQuestion.Description;
                            //            userInput.Position = itemQuestion.Position;
                            //            userInput.AppId = itemQuestion.AppId;
                            //            userInput.FieldId = itemQuestion.FieldId;
                            //            userInput.ItemId = 0;
                            //            userInput.Type = itemQuestion.Type;
                            //            userInput.DtEndRequire = itemQuestion.DtEndRequire;
                            //            userInput.CreatedOn = DateTime.Now;
                            //            userInput.UpdatedOn = DateTime.Now;
                            //            List<QuestionnaireOption> listOption = new List<QuestionnaireOption>();
                            //            foreach (var itemOptions in itemQuestion.Options)
                            //            {
                            //                QuestionnaireOption questionnaireOption = new QuestionnaireOption();
                            //                questionnaireOption.OptionName = itemOptions.OptionName;
                            //                questionnaireOption.OptionId = itemOptions.OptionId;
                            //                questionnaireOption.AppId = itemOptions.AppId;
                            //                questionnaireOption.CreatedOn = DateTime.Now;
                            //                listOption.Add(questionnaireOption);
                            //            }



                            //            //copy answer from previous draft
                            //            if (workpaperIndex == 5 || workpaperIndex == 0)
                            //            {
                            //                var prevAnswer = listPrevUserInput.Where(z => z.FieldId.Equals(userInput.FieldId) && z.StrQuestion.Equals(userInput.StrQuestion)).Select(ans => new { ans.StrAnswer, ans.StrAnswer2 });
                            //                userInput.StrAnswer = prevAnswer.Select(ans1 => ans1.StrAnswer).FirstOrDefault();
                            //                userInput.StrAnswer2 = prevAnswer.Select(ans2 => ans2.StrAnswer2).FirstOrDefault();

                            //            }
                            //            else 
                            //            {
                            //                //Set answer if new
                            //                userInput.StrAnswer = string.Empty;
                            //                userInput.StrAnswer2 = string.Empty;

                            //                #region populate answer base from RCM
                            //                if (questionnaireTesterNew.Rcm != null && itemQuestion.Type != "image")
                            //                {
                            //                    switch (itemQuestion.QuestionString.ToLower())
                            //                    {
                            //                        case string s when s.Contains("1. what is the client name?"):
                            //                            if (questionnaireTesterNew.Rcm.ClientName != null && questionnaireTesterNew.Rcm.ClientName != string.Empty)
                            //                            {
                            //                                userInput.StrAnswer = formatService.ReplaceTagHtmlParagraph(questionnaireTesterNew.Rcm.ClientName, false);
                            //                            }
                            //                            break;
                            //                        case string s when s.Contains("what is the purpose of this control?"):
                            //                            if (questionnaireTesterNew.Rcm.SpecificRisk != null && questionnaireTesterNew.Rcm.SpecificRisk != string.Empty)
                            //                            {
                            //                                userInput.StrAnswer = formatService.ReplaceTagHtmlParagraph(questionnaireTesterNew.Rcm.SpecificRisk, false);
                            //                            }
                            //                            break;
                            //                        case string s when s.Contains("what is the control id?"):
                            //                            if (questionnaireTesterNew.Rcm.ControlId != null && questionnaireTesterNew.Rcm.ControlId != string.Empty)
                            //                            {
                            //                                userInput.StrAnswer = formatService.ReplaceTagHtmlParagraph(questionnaireTesterNew.Rcm.ControlId, false);
                            //                            }
                            //                            break;
                            //                        case string s when s.Contains("what is the control activity?"):
                            //                            if (questionnaireTesterNew.Rcm.ControlActivityFy19 != null && questionnaireTesterNew.Rcm.ControlActivityFy19 != string.Empty)
                            //                            {
                            //                                userInput.StrAnswer = formatService.ReplaceTagHtmlParagraph(questionnaireTesterNew.Rcm.ControlActivityFy19, false);
                            //                            }
                            //                            break;
                            //                        case string s when s.Contains("who is the control owner?"):
                            //                            if (questionnaireTesterNew.Rcm.ControlOwner != null && questionnaireTesterNew.Rcm.ControlOwner != string.Empty)
                            //                            {
                            //                                userInput.StrAnswer = formatService.ReplaceTagHtmlParagraph(questionnaireTesterNew.Rcm.ControlOwner, false);
                            //                            }
                            //                            break;
                            //                        case string s when
                            //                            s.Contains("what are the procedures to test this control") ||
                            //                            s.Contains("what are the procedures to test control"):
                            //                            if (questionnaireTesterNew.Rcm.TestProc != null && questionnaireTesterNew.Rcm.TestProc != string.Empty)
                            //                            {
                            //                                userInput.StrAnswer = formatService.ReplaceTagHtmlParagraph(questionnaireTesterNew.Rcm.TestProc, false);
                            //                            }
                            //                            break;
                            //                        case string s when s.Contains("when was the control first put in place?"):
                            //                            if (questionnaireTesterNew.Rcm.ControlPlaceDate != null && questionnaireTesterNew.Rcm.ControlPlaceDate != string.Empty)
                            //                            {
                            //                                userInput.StrAnswer = formatService.ReplaceTagHtmlParagraph(questionnaireTesterNew.Rcm.ControlPlaceDate, false);
                            //                            }
                            //                            break;
                            //                        case string s when s.Contains("what is the date range you are testing?"):
                            //                            if (questionnaireTesterNew.Rcm.TestingPeriod != null && questionnaireTesterNew.Rcm.TestingPeriod != string.Empty)
                            //                            {
                            //                                userInput.StrAnswer = formatService.ReplaceTagHtmlParagraph(questionnaireTesterNew.Rcm.TestingPeriod, false);
                            //                            }
                            //                            break;
                            //                        case string s when s.Contains("what is the level of risk for the control?"):
                            //                            if (questionnaireTesterNew.Rcm.RiskLvl != null && questionnaireTesterNew.Rcm.RiskLvl != string.Empty)
                            //                            {
                            //                                userInput.StrAnswer = formatService.ReplaceTagHtmlParagraph(questionnaireTesterNew.Rcm.RiskLvl, false);
                            //                            }
                            //                            break;
                            //                        case string s when s.Contains("how often does this control happen?"):
                            //                            if (questionnaireTesterNew.Rcm.ControlFrequency != null && questionnaireTesterNew.Rcm.ControlFrequency != string.Empty)
                            //                            {
                            //                                userInput.StrAnswer = formatService.ReplaceTagHtmlParagraph(questionnaireTesterNew.Rcm.ControlFrequency, false);
                            //                            }
                            //                            break;
                            //                    }
                            //                }
                            //                #endregion


                            //            }

                            //            userInput.Options = listOption;
                            //            listUserInput.Add(userInput);

                            //        }

                            //    }


                            //    //Set IPE notes
                            //    List<IPENote> listIPE = new List<IPENote>();

                            //    //copy answer from previous draft
                            //    if ((workpaperIndex == 5 || workpaperIndex == 0))
                            //    {         
                            //        //copy previous IPE
                            //        if(listPrevIPE != null && listPrevIPE.Any())
                            //        {
                            //            foreach (var prevIPE in listPrevIPE)
                            //            {
                            //                prevIPE.Id = 0;
                            //                listIPE.Add(prevIPE);
                            //            }
                            //        }

                            //        if(prevGeneralNote != null)
                            //        {
                            //            prevGeneralNote.Id = 0;
                            //            questionnaireTesterNew.GeneralNote = prevGeneralNote;
                            //        }

                            //        questionnaireTesterNew.ListRoundItem2 = ListRoundItem2;
                            //        questionnaireTesterNew.ListUniqueNotes = ListUniqueNotes;
                            //        questionnaireTesterNew.ListHeaderNote2 = ListHeaderNote2;
                            //    }
                            //    else
                            //    {
                            //        //Set new IPE
                            //        if (questionnaireTester.Rcm.ControlId.ToLower().Contains("itgc"))
                            //        {
                            //            listIPE = new List<IPENote>
                            //            {
                            //                new IPENote
                            //                {
                            //                    Name = "IPE Note",
                            //                    Note = "IPE Note",
                            //                    Description = string.Empty,
                            //                    Display = false
                            //                }
                            //            };

                            //        }
                            //        else
                            //        {

                            //            listIPE = new List<IPENote>
                            //            {
                            //                new IPENote
                            //                {
                            //                    Name = "IPE Walkthrough",
                            //                    Note = "IPE/Validation of PBC Reports (Completeness/Accuracy) Walkthrough",
                            //                    Description = string.Empty,
                            //                    Display = false
                            //                },
                            //                new IPENote
                            //                {
                            //                    Name = "IPE Round 1",
                            //                    Note = "IPE/Validation of PBC Reports (Completeness/Accuracy) Round 1",
                            //                    Description = string.Empty,
                            //                    Display = false
                            //                },
                            //                new IPENote
                            //                {
                            //                    Name = "IPE Round 2",
                            //                    Note = "IPE/Validation of PBC Reports (Completeness/Accuracy) Round 2",
                            //                    Description = string.Empty,
                            //                    Display = false
                            //                },
                            //                    new IPENote
                            //                {
                            //                    Name = "IPE Round 3",
                            //                    Note = "IPE/Validation of PBC Reports (Completeness/Accuracy) Round 3",
                            //                    Description = string.Empty,
                            //                    Display = false
                            //                }
                            //            };
                            //        }


                            //        questionnaireTesterNew.ListRoundItem2 = questionnaireTester.ListRoundItem2;
                            //        questionnaireTesterNew.ListUniqueNotes = questionnaireTester.ListUniqueNotes;
                            //        questionnaireTesterNew.ListHeaderNote2 = questionnaireTester.ListHeaderNote2;
                            //    }



                            //    questionnaireTesterNew.ListIPENote = listIPE;
                            //    questionnaireTesterNew.Position = checkQuestionnaireTester.Position + 1;
                            //    questionnaireTesterNew.ListUserInputRound = listUserInput;

                            //    questionnaireTesterNew.SampleSel = questionnaireTester.SampleSel;


                            //    //foreach (var item in questionnaireTesterNew.ListUserInputRound)
                            //    //{
                            //    //    if(item.Options.Any())
                            //    //    {
                            //    //        foreach (var itemOption in item.Options)
                            //    //        {
                            //    //            _soxContext.Entry(itemOption).State = EntityState.Unchanged;
                            //    //        }

                            //    //    }

                            //    //}

                            //    _soxContext.Entry(questionnaireTesterNew.WorkpaperStatus).State = EntityState.Unchanged;
                            //    _soxContext.Entry(questionnaireTesterNew.Rcm).State = EntityState.Unchanged;

                            //    _soxContext.Add(questionnaireTesterNew);
                            //    await _soxContext.SaveChangesAsync();


                            //}

                            #endregion



                        }

                        context.Commit();
                    }

                    
                }
                
                catch (Exception ex)
                {
                    FileLog.Write($"Error SaveDraftWorkpaper {ex}", "ErrorSaveDraftWorkpaper");
                    AdminService adminService = new AdminService(_config);
                    adminService.SendAlert(true, true, ex.ToString(), "SaveDraftWorkpaper");
                    context.Rollback(); //rollback added data 
                    Console.WriteLine(ex.ToString());
                    return BadRequest();
                }
            }


            //Create new
            using (var context = _soxContext.Database.BeginTransaction())
            {

                try
                {
                    if (listQuestionnaireTesterSet != null && listQuestionnaireTesterSet.Any() && questionnaireTester != null)
                    {

                        //create new draft for tester or reviewer
                        //increment DraftNum
                        List<QuestionnaireUserAnswer> listUserInput = new List<QuestionnaireUserAnswer>();
                        QuestionnaireTesterSet questionnaireTesterNew = new QuestionnaireTesterSet();

                        questionnaireTesterNew.UniqueId = questionnaireTester.UniqueId;
                        int workpaperIndex = 0;
                        switch (questionnaireTester.UserAction)
                        {
                            case "Submit":
                                workpaperIndex = 3; //for review
                                questionnaireTesterNew.DraftNum = questionnaireTester.DraftNum;
                                break;
                            case "With Comments":
                                workpaperIndex = 5; //updated
                                questionnaireTesterNew.DraftNum = questionnaireTester.DraftNum + 1;
                                break;
                            case "Approved":
                                workpaperIndex = 0; //final
                                questionnaireTesterNew.DraftNum = questionnaireTester.DraftNum + 1;
                                questionnaireTesterNew.UserAction = "Final";
                                break;
                        }


                        //get workpaper status
                        var checkWorkPaperStatus = _soxContext.WorkpaperStatus.Where(index => index.Index.Equals(workpaperIndex)).AsNoTracking().FirstOrDefault();
                        if (checkWorkPaperStatus != null)
                        {
                            questionnaireTesterNew.WorkpaperStatus = checkWorkPaperStatus;
                        }



                        //check if item already created
                        //if next sequence is already created then there is no need to create
                        var checkNextQuestionnaireTester = _soxContext.QuestionnaireTesterSet
                            .Where(y => 
                                y.DraftNum.Equals(questionnaireTesterNew.DraftNum) && 
                                y.WorkpaperStatus.Equals(questionnaireTesterNew.WorkpaperStatus) &&
                                y.RoundName.Equals(questionnaireTesterNew.RoundName))
                            .FirstOrDefault();
                        if (checkNextQuestionnaireTester == null)
                        {
                            List<QuestionnaireTesterSet> listAllDraft = new List<QuestionnaireTesterSet>();
                            List<QuestionnaireUserAnswer> prevListUserInputRound = new List<QuestionnaireUserAnswer>();
                            List<IUCSystemGenAnswer> prevListIUCSystemGen = new List<IUCSystemGenAnswer>();
                            List<IUCNonSystemGenAnswer> prevListIUCNonSystemGen = new List<IUCNonSystemGenAnswer>();
                            List<NotesItem2> prevListUniqueNotes = new List<NotesItem2>();
                            List<RoundItem2> prevListRoundItem2 = new List<RoundItem2>();
                            List<HeaderNote2> prevListHeaderNote2 = new List<HeaderNote2>();
                            List<IPENote> prevListIPENote = new List<IPENote>();
                            SampleSelection prevSampleSel = new SampleSelection();
                            QuestionnaireTesterSet prevQuestionnaireTesterSet = new QuestionnaireTesterSet();
                            Rcm prevRcm = new Rcm();
                            int prevDraftIndex = 0;

                            if (questionnaireTesterNew.DraftNum > 1 && (workpaperIndex == 5 || workpaperIndex == 0))
                            {
                                listAllDraft = _soxContext.QuestionnaireTesterSet
                                    .Where(prev => prev.UniqueId.Equals(questionnaireTester.UniqueId) && prev.RcmItemId.Equals(questionnaireTester.RcmItemId))
                                    .AsNoTracking()
                                    .ToList();
                                if (listAllDraft.Any())
                                {
                                    prevDraftIndex = listAllDraft.Count() - 2;
                                    int prevId = listAllDraft[prevDraftIndex].Id;

                                    var listPrevDraft = _soxContext.QuestionnaireTesterSet
                                        .Where(testId => testId.Id.Equals(prevId))
                                        .Include(sysGen => sysGen.ListIUCSystemGen)
                                        .Include(nonSysGen => nonSysGen.ListIUCNonSystemGen)
                                        .Include(notes => notes.ListUniqueNotes)
                                        .Include(ipe => ipe.ListIPENote)
                                        .Include(header => header.ListHeaderNote2)
                                        .Include(stat => stat.WorkpaperStatus)
                                        .Include(genNote => genNote.GeneralNote)
                                        .Include(samp => samp.SampleSel)
                                        .Include(userInput => userInput.ListUserInputRound)
                                            .ThenInclude(inner => inner.Options)
                                        .Include(round => round.ListRoundItem2)
                                            .ThenInclude(qa => qa.ListRoundQA)
                                                .ThenInclude(opt => opt.Options)
                                        .AsNoTracking()
                                    .FirstOrDefault();

                                    if(listPrevDraft != null)
                                    {
                                        prevListUserInputRound = listPrevDraft.ListUserInputRound.ToList();
                                        prevListIUCSystemGen = listPrevDraft.ListIUCSystemGen.ToList();
                                        prevListIUCNonSystemGen = listPrevDraft.ListIUCNonSystemGen.ToList();
                                        prevListUniqueNotes = listPrevDraft.ListUniqueNotes.ToList();
                                        prevListRoundItem2 = listPrevDraft.ListRoundItem2.ToList();
                                        prevListHeaderNote2 = listPrevDraft.ListHeaderNote2.ToList();
                                        prevListIPENote = listPrevDraft.ListIPENote.ToList();
                                        prevSampleSel = listPrevDraft.SampleSel;
                                        prevQuestionnaireTesterSet = listPrevDraft;
                                        prevRcm = listPrevDraft.Rcm;

                                        //prevListUserInputRound = listQuestionnaireTesterSet[prevDraftIndex].ListUserInputRound.ToList();
                                        //prevListIUCSystemGen = listQuestionnaireTesterSet[prevDraftIndex].ListIUCSystemGen.ToList();
                                        //prevListIUCNonSystemGen = listQuestionnaireTesterSet[prevDraftIndex].ListIUCNonSystemGen.ToList();
                                        //prevListUniqueNotes = listQuestionnaireTesterSet[prevDraftIndex].ListUniqueNotes.ToList();
                                        //prevListRoundItem2 = listQuestionnaireTesterSet[prevDraftIndex].ListRoundItem2.ToList();
                                        //prevListHeaderNote2 = listQuestionnaireTesterSet[prevDraftIndex].ListHeaderNote2.ToList();
                                        //prevListIPENote = listQuestionnaireTesterSet[prevDraftIndex].ListIPENote.ToList();
                                        //prevSampleSel = listQuestionnaireTesterSet[prevDraftIndex].SampleSel;
                                        //prevQuestionnaireTesterSet = listQuestionnaireTesterSet[prevDraftIndex];
                                        //prevRcm = listQuestionnaireTesterSet[prevDraftIndex].Rcm;
                                    }

                                    #region Copy from previous tester draft and create new draft
                                    //Copy from previous tester draft and create new draft
                                    if (prevListUserInputRound != null && prevListUserInputRound.Any())
                                    {
                                        List<QuestionnaireUserAnswer> listUserAnswer = new List<QuestionnaireUserAnswer>();
                                        foreach (var item in prevListUserInputRound)
                                        {
                                            QuestionnaireUserAnswer userInput = new QuestionnaireUserAnswer();
                                            userInput.StrAnswer = item.StrAnswer;
                                            userInput.StrAnswer2 = item.StrAnswer2;
                                            userInput.StrQuestion = item.StrQuestion;
                                            userInput.IsDisabled = item.IsDisabled;
                                            userInput.StrDefaultAnswer = item.StrDefaultAnswer;
                                            userInput.Description = item.Description;
                                            userInput.Position = item.Position;
                                            userInput.AppId = item.AppId;
                                            userInput.FieldId = item.FieldId;
                                            userInput.ItemId = item.ItemId;
                                            userInput.Type = item.Type;
                                            userInput.DtEndRequire = item.DtEndRequire;
                                            userInput.RoundName = item.RoundName;
                                            userInput.CreatedOn = DateTime.UtcNow;

                                            if (item.Options.Any())
                                            {
                                                List<QuestionnaireOption> listOption = new List<QuestionnaireOption>();
                                                foreach (var itemOption in item.Options)
                                                {
                                                    QuestionnaireOption option = new QuestionnaireOption();
                                                    option.OptionName = itemOption.OptionName;
                                                    option.OptionId = itemOption.OptionId;
                                                    option.AppId = itemOption.AppId;
                                                    option.CreatedOn = DateTime.UtcNow;
                                                    listOption.Add(option);
                                                }
                                                userInput.Options = listOption;
                                            }


                                            listUserAnswer.Add(userInput);
                                        }
                                        questionnaireTesterNew.ListUserInputRound = listUserAnswer;
                                    }
                                    else
                                        questionnaireTesterNew.ListUserInputRound = null;


                                    if (prevListUniqueNotes != null && prevListUniqueNotes.Any())
                                    {

                                        List<NotesItem2> listNotesItem = new List<NotesItem2>();
                                        foreach (var item in prevListUniqueNotes)
                                        {
                                            NotesItem2 notesItem = new NotesItem2();
                                            notesItem.Position = item.PodioItemId;
                                            notesItem.Notes = item.Notes;
                                            notesItem.Description = item.Description;
                                            notesItem.PodioItemId = item.PodioItemId;
                                            listNotesItem.Add(notesItem);
                                        }
                                        questionnaireTesterNew.ListUniqueNotes = listNotesItem;
                                    }
                                    else
                                        questionnaireTesterNew.ListUniqueNotes = null;


                                    if (prevListRoundItem2 != null && prevListRoundItem2.Any())
                                    {

                                        List<RoundItem2> listRoundItem = new List<RoundItem2>();
                                        foreach (var item in prevListRoundItem2)
                                        {
                                            RoundItem2 roundItem = new RoundItem2();
                                            roundItem.ItemID = item.ItemID;
                                            roundItem.AppId = item.AppId;
                                            roundItem.RoundName = item.RoundName;
                                            roundItem.Position = item.Position;
                                            roundItem.A2Q2Samples = item.A2Q2Samples;
                                            roundItem.PodioItemId = item.PodioItemId;
                                            roundItem.PodioUniqueId = item.PodioUniqueId;
                                            roundItem.PodioLink = item.PodioLink;
                                            roundItem.CreatedOn = DateTime.UtcNow;

                                            if (item.ListRoundQA.Any())
                                            {
                                                List<RoundQA2> listRoundQA = new List<RoundQA2>();
                                                foreach (var itemQA in item.ListRoundQA)
                                                {
                                                    RoundQA2 roundQA = new RoundQA2();
                                                    roundQA.Position = itemQA.Position;
                                                    roundQA.Question = itemQA.Question;
                                                    roundQA.Answer = itemQA.Answer;
                                                    roundQA.Answer2 = itemQA.Answer2;
                                                    roundQA.Note = itemQA.Note;
                                                    roundQA.Type = itemQA.Type;
                                                    roundQA.DtEndRequire = itemQA.DtEndRequire;

                                                    if (itemQA.Options.Any())
                                                    {
                                                        List<QuestionnaireOption> listOptions = new List<QuestionnaireOption>();
                                                        foreach (var itemOptions in itemQA.Options)
                                                        {
                                                            QuestionnaireOption option = new QuestionnaireOption();
                                                            option.OptionName = itemOptions.OptionName;
                                                            option.OptionId = itemOptions.OptionId;
                                                            option.AppId = itemOptions.AppId;
                                                            option.CreatedOn = DateTime.UtcNow;
                                                            listOptions.Add(option);
                                                        }
                                                        roundQA.Options = listOptions;
                                                    }



                                                    listRoundQA.Add(roundQA);
                                                }
                                                roundItem.ListRoundQA = listRoundQA;
                                            }

                                            listRoundItem.Add(roundItem);
                                        }


                                        questionnaireTesterNew.ListRoundItem2 = listRoundItem;

                                    }
                                    else
                                        questionnaireTesterNew.ListRoundItem2 = null;


                                    if (prevQuestionnaireTesterSet.GeneralNote != null)
                                    {
                                        GeneralNote genNote = new GeneralNote();
                                        genNote.GeneralNoteText = prevQuestionnaireTesterSet.GeneralNote.GeneralNoteText;
                                        genNote.Description = prevQuestionnaireTesterSet.GeneralNote.Description;
                                        genNote.Display = prevQuestionnaireTesterSet.GeneralNote.Display;
                                        questionnaireTesterNew.GeneralNote = genNote;
                                    }
                                    else
                                        questionnaireTesterNew.GeneralNote = null;


                                    if (prevListIPENote != null && prevListIPENote.Any())
                                    {
                                        List<IPENote> listIPE = new List<IPENote>();
                                        foreach (var item in prevListIPENote)
                                        {
                                            IPENote ipeNote = new IPENote();
                                            ipeNote.Name = item.Name;
                                            ipeNote.Note = item.Note;
                                            ipeNote.Description = item.Description;
                                            ipeNote.Display = item.Display;
                                            listIPE.Add(ipeNote);
                                        }
                                        questionnaireTesterNew.ListIPENote = listIPE;
                                    }
                                    else
                                        questionnaireTesterNew.ListIPENote = null;


                                    if (questionnaireTester.SampleSel != null)
                                    {
                                        questionnaireTesterNew.SampleSel = prevQuestionnaireTesterSet.SampleSel;
                                    }
                                    else
                                        questionnaireTesterNew.SampleSel = null;


                                    if (questionnaireTester.ListIUCNonSystemGen != null)
                                    {
                                        questionnaireTesterNew.ListIUCNonSystemGen = prevQuestionnaireTesterSet.ListIUCNonSystemGen;
                                        //questionnaireTesterNew.SampleSel.Id = 0;
                                    }

                                    if (questionnaireTester.ListIUCSystemGen != null)
                                    {
                                        questionnaireTesterNew.ListIUCSystemGen = prevQuestionnaireTesterSet.ListIUCSystemGen;
                                        //questionnaireTesterNew.SampleSel.Id = 0;
                                    }

                                    #endregion

                                }
                            }
                            else
                            {
                                //Create blank item
                                if (questionnaireTester.ListUserInputRound != null && questionnaireTester.ListUserInputRound.Any())
                                {
                                    List<QuestionnaireUserAnswer> listUserAnswer = new List<QuestionnaireUserAnswer>();
                                    foreach (var item in questionnaireTester.ListUserInputRound)
                                    {
                                        QuestionnaireUserAnswer userInput = new QuestionnaireUserAnswer();
                                        userInput.StrAnswer = string.Empty;
                                        userInput.StrAnswer2 = string.Empty;
                                        userInput.StrQuestion = item.StrQuestion;
                                        userInput.IsDisabled = item.IsDisabled;
                                        userInput.StrDefaultAnswer = item.StrDefaultAnswer;
                                        userInput.Description = item.Description;
                                        userInput.Position = item.Position;
                                        userInput.AppId = item.AppId;
                                        userInput.FieldId = item.FieldId;
                                        userInput.ItemId = item.ItemId;
                                        userInput.Type = item.Type;
                                        userInput.DtEndRequire = item.DtEndRequire;
                                        userInput.RoundName = item.RoundName;
                                        userInput.CreatedOn = DateTime.UtcNow;

                                        if (item.Options.Any())
                                        {
                                            List<QuestionnaireOption> listOption = new List<QuestionnaireOption>();
                                            foreach (var itemOption in item.Options)
                                            {
                                                QuestionnaireOption option = new QuestionnaireOption();
                                                option.OptionName = itemOption.OptionName;
                                                option.OptionId = itemOption.OptionId;
                                                option.AppId = itemOption.AppId;
                                                option.CreatedOn = DateTime.UtcNow;
                                                listOption.Add(option);
                                            }
                                            userInput.Options = listOption;
                                        }


                                        listUserAnswer.Add(userInput);
                                    }
                                    questionnaireTesterNew.ListUserInputRound = listUserAnswer;
                                }
                                else
                                    questionnaireTesterNew.ListUserInputRound = null;


                                if (questionnaireTester.ListUniqueNotes != null && questionnaireTester.ListUniqueNotes.Any())
                                {

                                    List<NotesItem2> listNotesItem = new List<NotesItem2>();
                                    foreach (var item in questionnaireTester.ListUniqueNotes)
                                    {
                                        NotesItem2 notesItem = new NotesItem2();
                                        notesItem.Position = item.PodioItemId;
                                        notesItem.Notes = item.Notes;
                                        notesItem.Description = string.Empty;
                                        notesItem.PodioItemId = item.PodioItemId;
                                        listNotesItem.Add(notesItem);
                                    }
                                    questionnaireTesterNew.ListUniqueNotes = listNotesItem;
                                }
                                else
                                    questionnaireTesterNew.ListUniqueNotes = null;


                                if (questionnaireTester.ListRoundItem2 != null && questionnaireTester.ListRoundItem2.Any())
                                {

                                    List<RoundItem2> listRoundItem = new List<RoundItem2>();
                                    foreach (var item in questionnaireTester.ListRoundItem2)
                                    {
                                        RoundItem2 roundItem = new RoundItem2();
                                        roundItem.ItemID = item.ItemID;
                                        roundItem.AppId = item.AppId;
                                        roundItem.RoundName = item.RoundName;
                                        roundItem.Position = item.Position;
                                        roundItem.A2Q2Samples = item.A2Q2Samples;
                                        roundItem.PodioItemId = item.PodioItemId;
                                        roundItem.PodioUniqueId = item.PodioUniqueId;
                                        roundItem.PodioLink = item.PodioLink;
                                        roundItem.CreatedOn = DateTime.UtcNow;

                                        if (item.ListRoundQA.Any())
                                        {
                                            List<RoundQA2> listRoundQA = new List<RoundQA2>();
                                            foreach (var itemQA in item.ListRoundQA)
                                            {
                                                RoundQA2 roundQA = new RoundQA2();
                                                roundQA.Position = itemQA.Position;
                                                roundQA.Question = itemQA.Question;
                                                roundQA.Answer = string.Empty;
                                                roundQA.Answer2 = string.Empty;
                                                roundQA.Note = itemQA.Note;
                                                roundQA.Type = itemQA.Type;
                                                roundQA.DtEndRequire = itemQA.DtEndRequire;

                                                if (itemQA.Options.Any())
                                                {
                                                    List<QuestionnaireOption> listOptions = new List<QuestionnaireOption>();
                                                    foreach (var itemOptions in itemQA.Options)
                                                    {
                                                        QuestionnaireOption option = new QuestionnaireOption();
                                                        option.OptionName = itemOptions.OptionName;
                                                        option.OptionId = itemOptions.OptionId;
                                                        option.AppId = itemOptions.AppId;
                                                        option.CreatedOn = DateTime.UtcNow;
                                                        listOptions.Add(option);
                                                    }
                                                    roundQA.Options = listOptions;
                                                }



                                                listRoundQA.Add(roundQA);
                                            }
                                            roundItem.ListRoundQA = listRoundQA;
                                        }

                                        listRoundItem.Add(roundItem);
                                    }


                                    questionnaireTesterNew.ListRoundItem2 = listRoundItem;

                                }
                                else
                                    questionnaireTesterNew.ListRoundItem2 = null;


                                if (questionnaireTester.GeneralNote != null)
                                {
                                    GeneralNote genNote = new GeneralNote();
                                    genNote.GeneralNoteText = questionnaireTester.GeneralNote.GeneralNoteText;
                                    genNote.Description = string.Empty;
                                    genNote.Display = questionnaireTester.GeneralNote.Display;
                                    questionnaireTesterNew.GeneralNote = genNote;
                                }
                                else
                                    questionnaireTesterNew.GeneralNote = null;


                                if (questionnaireTester.ListIPENote != null && questionnaireTester.ListIPENote.Any())
                                {
                                    List<IPENote> listIPE = new List<IPENote>();
                                    foreach (var item in questionnaireTester.ListIPENote)
                                    {
                                        IPENote ipeNote = new IPENote();
                                        ipeNote.Name = item.Name;
                                        ipeNote.Note = item.Note;
                                        ipeNote.Description = string.Empty;
                                        ipeNote.Display = item.Display;
                                        listIPE.Add(ipeNote);
                                    }
                                    questionnaireTesterNew.ListIPENote = listIPE;
                                }
                                else
                                    questionnaireTesterNew.ListIPENote = null;



                                if (questionnaireTester.SampleSel != null)
                                {
                                    questionnaireTesterNew.SampleSel = questionnaireTester.SampleSel;
                                }
                                else
                                    questionnaireTesterNew.SampleSel = null;



                                if (questionnaireTester.ListIUCNonSystemGen != null)
                                {
                                    questionnaireTesterNew.ListIUCNonSystemGen = questionnaireTester.ListIUCNonSystemGen;
                                    //questionnaireTesterNew.SampleSel.Id = 0;
                                }

                                if (questionnaireTester.ListIUCSystemGen != null)
                                {
                                    questionnaireTesterNew.ListIUCSystemGen = questionnaireTester.ListIUCSystemGen;
                                    //questionnaireTesterNew.SampleSel.Id = 0;
                                }


                            }



                            questionnaireTesterNew.RoundName = questionnaireTester.RoundName;
                            questionnaireTesterNew.AddedBy = questionnaireTester.AddedBy;
                            questionnaireTesterNew.RcmItemId = questionnaireTester.RcmItemId;
                            questionnaireTesterNew.Rcm = questionnaireTester.Rcm;
                            questionnaireTesterNew.Position = questionnaireTester.Position + 1;
                            questionnaireTesterNew.CreatedOn = DateTime.UtcNow;

                            _soxContext.Entry(questionnaireTesterNew.WorkpaperStatus).State = EntityState.Unchanged;
                            _soxContext.Entry(questionnaireTesterNew.Rcm).State = EntityState.Unchanged;

                            if (questionnaireTesterNew.SampleSel != null)
                                _soxContext.Entry(questionnaireTesterNew.SampleSel).State = EntityState.Unchanged;

                            _soxContext.Add(questionnaireTesterNew);
                            await _soxContext.SaveChangesAsync();

                            context.Commit();
                        }
                    }

                    return Ok();

                }
                catch (Exception ex)
                {
                    FileLog.Write($"Error SaveDraftWorkpaper {ex}", "ErrorSaveDraftWorkpaper");
                    AdminService adminService = new AdminService(_config);
                    adminService.SendAlert(true, true, ex.ToString(), "SaveDraftWorkpaper");
                    context.Rollback(); //rollback added data 
                    Console.WriteLine(ex.ToString());
                    return BadRequest();
                }
            }
        
        
        }

        [AllowAnonymous]
        [HttpGet("getAllTesterWorkpaper")]
        public IActionResult GetTesterSet(int rcmItemId)
        {
            try
            {
                List<QuestionnaireTesterSet> listTesterSet = new List<QuestionnaireTesterSet>();
                var rcmCheck = _soxContext.Rcm.AsNoTracking().FirstOrDefault(x => x.PodioItemId == rcmItemId);
                if (rcmCheck != null)
                {
                    var testerSetCheck = _soxContext.QuestionnaireTesterSet
                        .Include(y => y.WorkpaperStatus)
                        .OrderBy(x => x.UniqueId).ThenByDescending(y => y.Position)
                        .Where(id => id.Rcm.Id == rcmCheck.Id)
                        .AsNoTracking()
                        .ToList();
                    if (testerSetCheck != null && testerSetCheck.Any())
                    {  

                        foreach (var item in testerSetCheck)
                        {
                            //listTesterSet.Add(item);

                            if (item != null)
                            {
                                var prevItem = listTesterSet.Where(dup => dup.UniqueId.Equals(item.UniqueId)).ToList();
                                //if item is not yet added
                                if(prevItem == null)
                                {
                                    listTesterSet.Add(item);
                                }
                                //else if unique id exists and round does not exists
                                else if (prevItem != null && prevItem.Where(x => x.RoundName.Equals(item.RoundName)).FirstOrDefault() == null)
                                {
                                    listTesterSet.Add(item);
                                }

                            }

                        }


                    }
                }

                return Ok(listTesterSet);
            }
            catch (Exception ex)
            {
                FileLog.Write($"Error GetTesterSet {ex}", "ErrorGetTesterSet");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetTesterSet");
                Console.WriteLine(ex.ToString());
                return BadRequest();
            }
        }

        [AllowAnonymous]
        [HttpPost("getFinalWorkpaper")]
        public IActionResult GetWorkpaperFinalSet([FromBody] QuestionnaireDraftId draftId)
        {
            try
            {
                List<QuestionnaireTesterSet> listFinalTesterSet = new List<QuestionnaireTesterSet>();


                var listDraftTester1 = _soxContext.QuestionnaireTesterSet
                .Where(x =>
                    x.UniqueId.Equals(draftId.UniqueId) &&
                    x.RcmItemId.Equals(draftId.RcmItemId) &&
                    x.RoundName.Equals("Round 1"))
                .Include(sysGen => sysGen.ListIUCSystemGen)
                .Include(nonSysGen => nonSysGen.ListIUCNonSystemGen)
                .Include(notes => notes.ListUniqueNotes)
                .Include(ipe => ipe.ListIPENote)
                .Include(header => header.ListHeaderNote2)
                .Include(stat => stat.WorkpaperStatus)
                .Include(genNote => genNote.GeneralNote)
                .Include(samp => samp.SampleSel)
                .Include(userInput => userInput.ListUserInputRound)
                    .ThenInclude(inner => inner.Options)
                .Include(round => round.ListRoundItem2)
                    .ThenInclude(qa => qa.ListRoundQA)
                        .ThenInclude(opt => opt.Options)
                .AsNoTracking()
                .OrderByDescending(dt => dt.DraftNum).ThenByDescending(upd => upd.UpdatedOn)
                .FirstOrDefault();
                //.ToList();
                if (listDraftTester1 != null)
                {

                    if (listDraftTester1.SampleSel != null)
                    {

                        var resTestround1 = _soxContext.TestRounds.FromSqlRaw($"CALL `sox`.`sp_get_testround`({listDraftTester1.SampleSel.Id});").AsNoTracking().ToList();
                        if (resTestround1.Count > 0)
                        {
                            listDraftTester1.SampleSel.ListTestRound = resTestround1;
                        }
                    }

                    if (listDraftTester1.ListUserInputRound != null && listDraftTester1.ListUserInputRound.Any())
                    {
                        listDraftTester1.ListUserInputRound = listDraftTester1.ListUserInputRound.OrderBy(sort => sort.Position).ToList();
                    }

                    listFinalTesterSet.Add(listDraftTester1);

                }
                else
                {
                    QuestionnaireTesterSet testerSetRound1 = new QuestionnaireTesterSet();
                    testerSetRound1.Id = 0;
                    testerSetRound1.UniqueId = draftId.UniqueId;
                    testerSetRound1.UniqueId = draftId.UniqueId;
                    testerSetRound1.RoundName = "Round 1";
                    listFinalTesterSet.Add(listDraftTester1);

                }

                var listDraftTester2 = _soxContext.QuestionnaireTesterSet
                .Where(x =>
                    x.UniqueId.Equals(draftId.UniqueId) &&
                    x.RcmItemId.Equals(draftId.RcmItemId) &&
                    x.RoundName.Equals("Round 2"))
                .Include(sysGen => sysGen.ListIUCSystemGen)
                .Include(nonSysGen => nonSysGen.ListIUCNonSystemGen)
                .Include(notes => notes.ListUniqueNotes)
                .Include(ipe => ipe.ListIPENote)
                .Include(header => header.ListHeaderNote2)
                .Include(stat => stat.WorkpaperStatus)
                .Include(genNote => genNote.GeneralNote)
                .Include(samp => samp.SampleSel)
                .Include(userInput => userInput.ListUserInputRound)
                    .ThenInclude(inner => inner.Options)
                .Include(round => round.ListRoundItem2)
                    .ThenInclude(qa => qa.ListRoundQA)
                        .ThenInclude(opt => opt.Options)
                .AsNoTracking()
                .OrderByDescending(dt => dt.DraftNum).ThenByDescending(upd => upd.UpdatedOn)
                .FirstOrDefault();
                //.ToList();
                if (listDraftTester2 != null)
                {

                    if (listDraftTester2.SampleSel != null)
                    {

                        var resTestround2 = _soxContext.TestRounds.FromSqlRaw($"CALL `sox`.`sp_get_testround`({listDraftTester2.SampleSel.Id});").AsNoTracking().ToList();
                        if (resTestround2.Count > 0)
                        {
                            listDraftTester2.SampleSel.ListTestRound = resTestround2;
                        }
                    }

                    if (listDraftTester2.ListUserInputRound != null && listDraftTester2.ListUserInputRound.Any())
                    {
                        listDraftTester2.ListUserInputRound = listDraftTester2.ListUserInputRound.OrderBy(sort => sort.Position).ToList();
                    }

                    listFinalTesterSet.Add(listDraftTester2);

                }
                else
                {
                    QuestionnaireTesterSet testerSetRound2 = new QuestionnaireTesterSet();
                    testerSetRound2.Id = 0;
                    testerSetRound2.UniqueId = draftId.UniqueId;
                    testerSetRound2.RcmItemId = draftId.RcmItemId;
                    testerSetRound2.RoundName = "Round 2";
                    listFinalTesterSet.Add(testerSetRound2);
                }


                var listDraftTester3 = _soxContext.QuestionnaireTesterSet
                .Where(x =>
                    x.UniqueId.Equals(draftId.UniqueId) &&
                    x.RcmItemId.Equals(draftId.RcmItemId) &&
                    x.RoundName.Equals("Round 3"))
                .Include(sysGen => sysGen.ListIUCSystemGen)
                .Include(nonSysGen => nonSysGen.ListIUCNonSystemGen)
                .Include(notes => notes.ListUniqueNotes)
                .Include(ipe => ipe.ListIPENote)
                .Include(header => header.ListHeaderNote2)
                .Include(stat => stat.WorkpaperStatus)
                .Include(genNote => genNote.GeneralNote)
                .Include(samp => samp.SampleSel)
                .Include(userInput => userInput.ListUserInputRound)
                    .ThenInclude(inner => inner.Options)
                .Include(round => round.ListRoundItem2)
                    .ThenInclude(qa => qa.ListRoundQA)
                        .ThenInclude(opt => opt.Options)
                .AsNoTracking()
                .OrderByDescending(dt => dt.DraftNum).ThenByDescending(upd => upd.UpdatedOn)
                .FirstOrDefault();
                //.ToList();
                if (listDraftTester3 != null)
                {

                    if (listDraftTester3.SampleSel != null)
                    {

                        var resTestround2 = _soxContext.TestRounds.FromSqlRaw($"CALL `sox`.`sp_get_testround`({listDraftTester3.SampleSel.Id});").AsNoTracking().ToList();
                        if (resTestround2.Count > 0)
                        {
                            listDraftTester3.SampleSel.ListTestRound = resTestround2;
                        }
                    }

                    if (listDraftTester3.ListUserInputRound != null && listDraftTester3.ListUserInputRound.Any())
                    {
                        listDraftTester3.ListUserInputRound = listDraftTester3.ListUserInputRound.OrderBy(sort => sort.Position).ToList();
                    }

                    listFinalTesterSet.Add(listDraftTester3);

                }
                else
                {
                    QuestionnaireTesterSet testerSetRound3 = new QuestionnaireTesterSet();
                    testerSetRound3.Id = 0;
                    testerSetRound3.UniqueId = draftId.UniqueId;
                    testerSetRound3.RcmItemId = draftId.RcmItemId;
                    testerSetRound3.RoundName = "Round 3";
                    listFinalTesterSet.Add(testerSetRound3);
                }


                return Ok(listFinalTesterSet);
            }
            catch (Exception ex)
            {
                FileLog.Write($"Error GetWorkpaperFinalSet {ex}", "ErrorGetWorkpaperFinalSet");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetWorkpaperFinalSet");
                Console.WriteLine(ex.ToString());
                return BadRequest();
            }
        }


        [HttpPost("excel/old/create")]
        public string CreateExcel([FromBody] Questionnaire questionnaire)
        {
            //List<string> excelFilename = new List<string>();
            string excelFilename = string.Empty;

            try
            {

                //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage xls = new ExcelPackage())
                {

                    var ws = xls.Workbook.Worksheets.Add(questionnaire.ControlId);
                    ws.View.ShowGridLines = false;

                    ExcelService xlsService = new ExcelService();
                    ws.Column(1).Width = 30;
                    ws.Column(2).Width = 30;
                    ws.Column(3).Width = 30;
                    ws.Column(4).Width = 30;
                    ws.Column(5).Width = 30;
                    ws.Column(6).Width = 30;
                    ws.Column(7).Width = 30;
                    ws.Column(8).Width = 30;

                    ws.Cells["A6"].Value = "Process:";
                    ws.Cells["C6"].Value = questionnaire.Process != null ? questionnaire.Process : string.Empty;
                    ws.Cells["A6:B6"].Merge = true;
                    ws.Cells["C6:H6"].Merge = true;
                    xlsService.ExcelWrapText(ws, "A6:H6");
                    xlsService.ExcelSetBorder(ws, "A6:H6");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A6:B6");
                    xlsService.ExcelSetArialSize10(ws, "A6:H6");
                    xlsService.ExcelSetFontBold(ws, "A6");

                    ws.Cells["A7"].Value = "Control Owner:";
                    ws.Cells["C7"].Value = questionnaire.ControlOwner != null ? questionnaire.ControlOwner : string.Empty;
                    ws.Cells["A7:B7"].Merge = true;
                    ws.Cells["C7:H7"].Merge = true;
                    xlsService.ExcelWrapText(ws, "A7:H7");
                    xlsService.ExcelSetBorder(ws, "A7:H7");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A7:B7");
                    xlsService.ExcelSetArialSize10(ws, "A7:H7");
                    xlsService.ExcelSetFontBold(ws, "A7");

                    ws.Cells["A8"].Value = "Control Id:";
                    ws.Cells["C8"].Value = questionnaire.ControlId != null ? questionnaire.ControlId : string.Empty;
                    ws.Cells["A8:B8"].Merge = true;
                    ws.Cells["C8:H8"].Merge = true;
                    xlsService.ExcelWrapText(ws, "A8:H8");
                    xlsService.ExcelSetBorder(ws, "A8:H8");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A8:B8");
                    xlsService.ExcelSetArialSize10(ws, "A8:H8");
                    xlsService.ExcelSetFontBold(ws, "A8");

                    ws.Cells["A9"].Value = "Control Description:";
                    ws.Cells["C9"].Value = questionnaire.ControlActivity != null ? questionnaire.ControlActivity : string.Empty;
                    ws.Cells["A9:B9"].Merge = true;
                    ws.Cells["C9:H9"].Merge = true;
                    xlsService.ExcelWrapText(ws, "A9:H9");
                    xlsService.ExcelSetBorder(ws, "A9:H9");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A9:B9");
                    xlsService.ExcelSetArialSize10(ws, "A9:H9");
                    xlsService.ExcelSetFontBold(ws, "A9");
                    xlsService.ExcelSetVerticalAlignCenter(ws, "A9:H9");
                    ws.Row(9).Height = 60;

                    ws.Cells["A11"].Value = "Test Validation Approach:";
                    ws.Cells["C11"].Value = questionnaire.TestValidation != null ? questionnaire.TestValidation : string.Empty;
                    ws.Cells["A11:B11"].Merge = true;
                    ws.Cells["C11:H11"].Merge = true;
                    xlsService.ExcelWrapText(ws, "A11:H11");
                    xlsService.ExcelSetBorder(ws, "A11:H11");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A11:B11");
                    xlsService.ExcelSetArialSize10(ws, "A11:H11");
                    xlsService.ExcelSetFontBold(ws, "A11");
                    xlsService.ExcelSetVerticalAlignCenter(ws, "A11:H11");
                    ws.Row(11).Height = 60;

                    ws.Cells["A12"].Value = "Test Method Used:";
                    ws.Cells["C12"].Value = questionnaire.MethodUsed != null ? questionnaire.MethodUsed : string.Empty;
                    ws.Cells["A12:B12"].Merge = true;
                    ws.Cells["C12:H12"].Merge = true;
                    xlsService.ExcelWrapText(ws, "A12:H12");
                    xlsService.ExcelSetBorder(ws, "A12:H12");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A12:B12");
                    xlsService.ExcelSetArialSize10(ws, "A12:H12");
                    xlsService.ExcelSetFontBold(ws, "A12");


                    ws.Cells["A13"].Value = "Control Frequency:";
                    ws.Cells["C13"].Value = questionnaire.ControlFrequency != null ? questionnaire.ControlFrequency : string.Empty;
                    ws.Cells["A13:B13"].Merge = true;
                    ws.Cells["C13:H13"].Merge = true;
                    xlsService.ExcelWrapText(ws, "A13:H13");
                    xlsService.ExcelSetBorder(ws, "A13:H13");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A13:B13");
                    xlsService.ExcelSetArialSize10(ws, "A13:H13");
                    xlsService.ExcelSetFontBold(ws, "A13");


                    ws.Cells["A14"].Value = "Control In Place Date:";
                    ws.Cells["C14"].Value = questionnaire.ControlPlaceDate != null ? questionnaire.ControlPlaceDate : string.Empty;
                    ws.Cells["A14:B14"].Merge = true;
                    ws.Cells["C14:H14"].Merge = true;
                    xlsService.ExcelWrapText(ws, "A14:H14");
                    xlsService.ExcelSetBorder(ws, "A14:H14");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A14:B14");
                    xlsService.ExcelSetArialSize10(ws, "A14:H14");
                    xlsService.ExcelSetFontBold(ws, "A14");


                    ws.Cells["A15"].Value = "Risk Assessment:";
                    ws.Cells["C15"].Value = questionnaire.RiskAssessment != null ? questionnaire.RiskAssessment : string.Empty;
                    ws.Cells["A15:B15"].Merge = true;
                    ws.Cells["C15:H15"].Merge = true;
                    xlsService.ExcelWrapText(ws, "A15:H15");
                    xlsService.ExcelSetBorder(ws, "A15:H15");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A15:B15");
                    xlsService.ExcelSetArialSize10(ws, "A15:H15");
                    xlsService.ExcelSetFontBold(ws, "A15");


                    ws.Cells["A16"].Value = "Sample Period:";
                    ws.Cells["C16"].Value = questionnaire.SamplePeriod != null ? questionnaire.SamplePeriod : string.Empty;
                    ws.Cells["A16:B16"].Merge = true;
                    ws.Cells["C16:H16"].Merge = true;
                    xlsService.ExcelWrapText(ws, "A16:H16");
                    xlsService.ExcelSetBorder(ws, "A16:H16");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A16:B16");
                    xlsService.ExcelSetArialSize10(ws, "A16:H16");
                    xlsService.ExcelSetFontBold(ws, "A16");


                    ws.Cells["A17"].Value = "Population Size:";
                    ws.Cells["C17"].Value = questionnaire.PopulationSize != null ? questionnaire.PopulationSize : string.Empty;
                    ws.Cells["A17:B17"].Merge = true;
                    ws.Cells["C17:H17"].Merge = true;
                    xlsService.ExcelWrapText(ws, "A17:H17");
                    xlsService.ExcelSetBorder(ws, "A17:H17");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A17:B17");
                    xlsService.ExcelSetArialSize10(ws, "A17:H17");
                    xlsService.ExcelSetFontBold(ws, "A17");


                    ws.Cells["A18"].Value = "Sample Derivation:";
                    ws.Cells["C18"].Value = questionnaire.SampleSizeDerivation != null ? questionnaire.SampleSizeDerivation : string.Empty;
                    ws.Cells["A18:B18"].Merge = true;
                    ws.Cells["C18:H18"].Merge = true;
                    xlsService.ExcelWrapText(ws, "A18:H18");
                    xlsService.ExcelSetBorder(ws, "A18:H18");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A18:B18");
                    xlsService.ExcelSetArialSize10(ws, "A18:H18");
                    xlsService.ExcelSetFontBold(ws, "A18");


                    ws.Cells["A19"].Value = "Electronic Audit Evidence:";
                    ws.Cells["C19"].Value = questionnaire.IpeInformation != null ? questionnaire.IpeInformation : string.Empty;
                    ws.Cells["A19:B19"].Merge = true;
                    ws.Cells["C19:H19"].Merge = true;
                    xlsService.ExcelWrapText(ws, "A19:H19");
                    xlsService.ExcelSetBorder(ws, "A19:H19");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A19:B19");
                    xlsService.ExcelSetArialSize10(ws, "A19:H19");
                    xlsService.ExcelSetFontBold(ws, "A19");


                    #region Round Header
                    ws.Cells["C21"].Value = "Round 1";
                    ws.Cells["C21:D21"].Merge = true;
                    xlsService.ExcelWrapText(ws, "C21:D21");
                    xlsService.ExcelSetBorder(ws, "C21:D21");

                    ws.Cells["E21"].Value = "Round 2";
                    ws.Cells["E21:F21"].Merge = true;
                    xlsService.ExcelWrapText(ws, "E21:F21");
                    xlsService.ExcelSetBorder(ws, "E21:F21");

                    ws.Cells["G21"].Value = "Round 3";
                    ws.Cells["G21:H21"].Merge = true;
                    xlsService.ExcelWrapText(ws, "G21:H21");
                    xlsService.ExcelSetBorder(ws, "G21:H21");

                    xlsService.ExcelSetBackgroundColorGray(ws, "C21:H21");
                    xlsService.ExcelSetArialSize10(ws, "C21:H21");
                    xlsService.ExcelSetFontBold(ws, "C21:H21");
                    xlsService.ExcelSetVerticalAlignCenter(ws, "C21:H21");
                    xlsService.ExcelSetHorizontalAlignCenter(ws, "C21:H21");
                    #endregion

                    ws.Cells["A22"].Value = "Sample Size:";
                    ws.Cells["C22"].Value = questionnaire.SampleSizeRound1;
                    ws.Cells["E22"].Value = questionnaire.SampleSizeRound2;
                    ws.Cells["G22"].Value = questionnaire.SampleSizeRound3;
                    ws.Cells["A22:B22"].Merge = true;
                    ws.Cells["C22:D22"].Merge = true;
                    ws.Cells["E22:F22"].Merge = true;
                    ws.Cells["G22:H22"].Merge = true;
                    xlsService.ExcelWrapText(ws, "A22:H22");
                    xlsService.ExcelSetBorder(ws, "A22:H22");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A22:B22");
                    xlsService.ExcelSetArialSize10(ws, "A22:H22");
                    xlsService.ExcelSetFontBold(ws, "A22");
                    xlsService.ExcelSetVerticalAlignCenter(ws, "C22:H22");
                    xlsService.ExcelSetHorizontalAlignCenter(ws, "C22:H22");


                    ws.Cells["A23"].Value = "Source File (document name, hardcopy/softcopy & provided by):";
                    ws.Cells["C23"].Value = questionnaire.SourceFileType;
                    ws.Cells["E23"].Value = string.Empty;
                    ws.Cells["G23"].Value = string.Empty;
                    ws.Cells["A23:B23"].Merge = true;
                    ws.Cells["C23:D23"].Merge = true;
                    ws.Cells["E23:F23"].Merge = true;
                    ws.Cells["G23:H23"].Merge = true;
                    xlsService.ExcelWrapText(ws, "A23:H23");
                    xlsService.ExcelSetBorder(ws, "A23:H23");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A23:B23");
                    xlsService.ExcelSetArialSize10(ws, "A23:H23");
                    xlsService.ExcelSetFontBold(ws, "A23");
                    xlsService.ExcelSetVerticalAlignCenter(ws, "A23:H23");
                    ws.Row(23).Height = 320;


                    ws.Cells["A24"].Value = "Test Performed By:";
                    ws.Cells["C24"].Value = questionnaire.TestPerfomedBy;
                    ws.Cells["E24"].Value = string.Empty;
                    ws.Cells["G24"].Value = string.Empty;
                    ws.Cells["A24:B24"].Merge = true;
                    ws.Cells["C24:D24"].Merge = true;
                    ws.Cells["E24:F24"].Merge = true;
                    ws.Cells["G24:H24"].Merge = true;
                    xlsService.ExcelWrapText(ws, "A24:H24");
                    xlsService.ExcelSetBorder(ws, "A24:H24");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A24:B24");
                    xlsService.ExcelSetArialSize10(ws, "A24:H24");
                    xlsService.ExcelSetFontBold(ws, "A24");
                    xlsService.ExcelSetVerticalAlignCenter(ws, "A24:H24");
                    xlsService.ExcelSetHorizontalAlignCenter(ws, "C24:H24");


                    ws.Cells["A25"].Value = "Date Testing Performed:";
                    ws.Cells["C25"].Value = questionnaire.DateOfTesting != null ? questionnaire.DateOfTesting.Value.ToString("MM/dd/yyyy") : string.Empty;
                    ws.Cells["E25"].Value = string.Empty;
                    ws.Cells["G25"].Value = string.Empty;
                    ws.Cells["A25:B25"].Merge = true;
                    ws.Cells["C25:D25"].Merge = true;
                    ws.Cells["E25:F25"].Merge = true;
                    ws.Cells["G25:H25"].Merge = true;
                    xlsService.ExcelWrapText(ws, "A25:H25");
                    xlsService.ExcelSetBorder(ws, "A25:H25");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A25:B25");
                    xlsService.ExcelSetArialSize10(ws, "A25:H25");
                    xlsService.ExcelSetFontBold(ws, "A25");
                    xlsService.ExcelSetVerticalAlignCenter(ws, "A25:H25");
                    xlsService.ExcelSetHorizontalAlignCenter(ws, "C25:H25");


                    ws.Cells["A26"].Value = "Test of Design (TOD):";
                    ws.Cells["C26"].Value = questionnaire.TestOfDesign != null ? questionnaire.TestOfDesign : string.Empty;
                    ws.Cells["E26"].Value = string.Empty;
                    ws.Cells["G26"].Value = string.Empty;
                    ws.Cells["A26:B26"].Merge = true;
                    ws.Cells["C26:D26"].Merge = true;
                    ws.Cells["E26:F26"].Merge = true;
                    ws.Cells["G26:H26"].Merge = true;
                    xlsService.ExcelWrapText(ws, "A26:H26");
                    xlsService.ExcelSetBorder(ws, "A26:H26");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A26:B26");
                    xlsService.ExcelSetArialSize10(ws, "A26:H26");
                    xlsService.ExcelSetFontBold(ws, "A26");
                    xlsService.ExcelSetVerticalAlignCenter(ws, "A26:H26");
                    xlsService.ExcelSetHorizontalAlignCenter(ws, "C26:H26");


                    ws.Cells["A27"].Value = "Test of Operating Effectiveness (TOE):";
                    ws.Cells["C27"].Value = questionnaire.TestOperatingEffectiveness != null ? questionnaire.TestOperatingEffectiveness : string.Empty;
                    ws.Cells["E27"].Value = string.Empty;
                    ws.Cells["G27"].Value = string.Empty;
                    ws.Cells["A27:B27"].Merge = true;
                    ws.Cells["C27:D27"].Merge = true;
                    ws.Cells["E27:F27"].Merge = true;
                    ws.Cells["G27:H27"].Merge = true;
                    xlsService.ExcelWrapText(ws, "A27:H27");
                    xlsService.ExcelSetBorder(ws, "A27:H27");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A27:B27");
                    xlsService.ExcelSetArialSize10(ws, "A27:H27");
                    xlsService.ExcelSetFontBold(ws, "A27");
                    xlsService.ExcelSetVerticalAlignCenter(ws, "A27:H27");
                    xlsService.ExcelSetHorizontalAlignCenter(ws, "C27:H27");


                    #region Round Sample

                    ws.Cells["A29"].Value = "Sample #";
                    ws.Cells["B29"].Value = "Period";
                    ws.Cells["C29"].Value = "Any investigation?";
                    ws.Cells["D29"].Value = "Medrio costs agree to the Clinical trial agreements [A]";
                    ws.Cells["E29"].Value = "Evidence of review";
                    ws.Cells["F29"].Value = "Reviewer";
                    ws.Cells["G29"].Value = "Review date";
                    ws.Cells["H29"].Value = "Notes";
                    xlsService.ExcelWrapText(ws, "A29:H29");
                    xlsService.ExcelSetBorder(ws, "A29:H29");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A29:H29");
                    xlsService.ExcelSetArialSize10(ws, "A29:H29");
                    xlsService.ExcelSetFontBold(ws, "A29:H29");
                    xlsService.ExcelSetVerticalAlignCenter(ws, "A29:H29");
                    xlsService.ExcelSetHorizontalAlignCenter(ws, "A29:H29");

                    int row = 30;

                    //Round 1
                    ws.Cells["A" + row].Value = "Round 1";
                    ws.Cells["A" + row + ":H" + row].Merge = true;
                    xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetBackgroundColorLightGray(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetFontBold(ws, "A" + row + ":H" + row);
                    row++;

                    var listRound1 = questionnaire.ListSampleRound.Where(x => x.RoundName == "Round 1");
                    if (listRound1 != null && listRound1.Count() > 0)
                    {
                        foreach (var item in listRound1)
                        {
                            ws.Cells["A" + row].Value = $"{(item.SampleNum != null ? item.SampleNum.ToString() : string.Empty)}";
                            ws.Cells["B" + row].Value = $"{(item.MedrioActivity16A != null ? item.MedrioActivity16A.ToString() : string.Empty)} [{item.TickBox1Value}]";
                            ws.Cells["C" + row].Value = $"{(item.InvestigationPerformed16B != null ? item.InvestigationPerformed16B.ToString() : string.Empty)} [{item.TickBox2Value}]";
                            ws.Cells["D" + row].Value = $"{(item.MedrioCostAgreements16C != null ? item.MedrioCostAgreements16C.ToString() : string.Empty)} [{item.TickBox3Value}]";
                            ws.Cells["E" + row].Value = $"{(item.MedrioReportReview16D != null ? item.MedrioReportReview16D.ToString() : string.Empty)} [{item.TickBox4Value}]";
                            ws.Cells["F" + row].Value = $"{(item.MedrioReviewers16E != null ? item.MedrioReviewers16E.ToString() : string.Empty)} [{item.TickBox5Value}]";
                            ws.Cells["G" + row].Value = $"{(item.DateMedrioReviewed16F != null ? item.DateMedrioReviewed16F.Value.ToString("MM/dd/yyyy") : string.Empty)} [{item.TickBox6Value}]";
                            ws.Cells["H" + row].Value = item.Notes;
                            xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetHorizontalAlignCenter(ws, "A" + row + ":H" + row);
                            ws.Row(row).Height = 30;
                            row++;
                        }
                    }

                    //Round 2
                    ws.Cells["A" + row].Value = "Round 2";
                    ws.Cells["A" + row + ":H" + row].Merge = true;
                    xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetBackgroundColorLightGray(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetFontBold(ws, "A" + row + ":H" + row);
                    row++;

                    var listRound2 = questionnaire.ListSampleRound.Where(x => x.RoundName == "Round 2");
                    if (listRound2 != null && listRound2.Count() > 0)
                    {
                        foreach (var item in listRound2)
                        {
                            ws.Cells["A" + row].Value = $"{(item.SampleNum != null ? item.SampleNum.ToString() : string.Empty)}";
                            ws.Cells["B" + row].Value = $"{(item.MedrioActivity16A != null ? item.MedrioActivity16A.ToString() : string.Empty)} [{item.TickBox1Value}]";
                            ws.Cells["C" + row].Value = $"{(item.InvestigationPerformed16B != null ? item.InvestigationPerformed16B.ToString() : string.Empty)} [{item.TickBox2Value}]";
                            ws.Cells["D" + row].Value = $"{(item.MedrioCostAgreements16C != null ? item.MedrioCostAgreements16C.ToString() : string.Empty)} [{item.TickBox3Value}]";
                            ws.Cells["E" + row].Value = $"{(item.MedrioReportReview16D != null ? item.MedrioReportReview16D.ToString() : string.Empty)} [{item.TickBox4Value}]";
                            ws.Cells["F" + row].Value = $"{(item.MedrioReviewers16E != null ? item.MedrioReviewers16E.ToString() : string.Empty)} [{item.TickBox5Value}]";
                            ws.Cells["G" + row].Value = $"{(item.DateMedrioReviewed16F != null ? item.DateMedrioReviewed16F.Value.ToString("MM/dd/yyyy") : string.Empty)} [{item.TickBox6Value}]";
                            ws.Cells["H" + row].Value = item.Notes;
                            xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetHorizontalAlignCenter(ws, "A" + row + ":H" + row);
                            ws.Row(row).Height = 30;
                            row++;
                        }
                    }

                    //Round 3
                    ws.Cells["A" + row].Value = "Round 3";
                    ws.Cells["A" + row + ":H" + row].Merge = true;
                    xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetBackgroundColorLightGray(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetFontBold(ws, "A" + row + ":H" + row);
                    row++;

                    var listRound3 = questionnaire.ListSampleRound.Where(x => x.RoundName == "Round 3");
                    if (listRound3 != null && listRound3.Count() > 0)
                    {
                        foreach (var item in listRound3)
                        {
                            ws.Cells["A" + row].Value = $"{(item.SampleNum != null ? item.SampleNum.ToString() : string.Empty)}";
                            ws.Cells["B" + row].Value = $"{(item.MedrioActivity16A != null ? item.MedrioActivity16A.ToString() : string.Empty)} [{item.TickBox1Value}]";
                            ws.Cells["C" + row].Value = $"{(item.InvestigationPerformed16B != null ? item.InvestigationPerformed16B.ToString() : string.Empty)} [{item.TickBox2Value}]";
                            ws.Cells["D" + row].Value = $"{(item.MedrioCostAgreements16C != null ? item.MedrioCostAgreements16C.ToString() : string.Empty)} [{item.TickBox3Value}]";
                            ws.Cells["E" + row].Value = $"{(item.MedrioReportReview16D != null ? item.MedrioReportReview16D.ToString() : string.Empty)} [{item.TickBox4Value}]";
                            ws.Cells["F" + row].Value = $"{(item.MedrioReviewers16E != null ? item.MedrioReviewers16E.ToString() : string.Empty)} [{item.TickBox5Value}]";
                            ws.Cells["G" + row].Value = $"{(item.DateMedrioReviewed16F != null ? item.DateMedrioReviewed16F.Value.ToString("MM/dd/yyyy") : string.Empty)} [{item.TickBox6Value}]";
                            ws.Cells["H" + row].Value = item.Notes;
                            xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetHorizontalAlignCenter(ws, "A" + row + ":H" + row);
                            ws.Row(row).Height = 30;
                            row++;
                        }
                    }

                    #endregion 


                    #region Round Header
                    row++;
                    ws.Cells["C" + row].Value = "Round 1";
                    ws.Cells["C" + row + ":D" + row].Merge = true;
                    xlsService.ExcelWrapText(ws, "C" + row + ":D" + row);
                    xlsService.ExcelSetBorder(ws, "C" + row + ":D" + row);

                    ws.Cells["E" + row].Value = "Round 2";
                    ws.Cells["E" + row + ":F" + row].Merge = true;
                    xlsService.ExcelWrapText(ws, "E" + row + ":F" + row);
                    xlsService.ExcelSetBorder(ws, "E" + row + ":F" + row);

                    ws.Cells["G" + row].Value = "Round 3";
                    ws.Cells["G" + row + ":H" + row].Merge = true;
                    xlsService.ExcelWrapText(ws, "G" + row + ":H" + row);
                    xlsService.ExcelSetBorder(ws, "G" + row + ":H" + row);

                    xlsService.ExcelSetBackgroundColorGray(ws, "C" + row + ":H" + row);
                    xlsService.ExcelSetArialSize10(ws, "C" + row + ":H" + row);
                    xlsService.ExcelSetFontBold(ws, "C" + row + ":H" + row);
                    xlsService.ExcelSetVerticalAlignCenter(ws, "C" + row + ":H" + row);
                    xlsService.ExcelSetHorizontalAlignCenter(ws, "C" + row + ":H" + row);
                    row++;
                    #endregion


                    ws.Cells["A" + row].Value = "Reviewer:";
                    ws.Cells["C" + row].Value = questionnaire.ReviewedBy != null ? questionnaire.ReviewedBy : string.Empty;
                    ws.Cells["E" + row].Value = string.Empty;
                    ws.Cells["G" + row].Value = string.Empty;
                    ws.Cells["A" + row + ":B" + row].Merge = true;
                    ws.Cells["C" + row + ":D" + row].Merge = true;
                    ws.Cells["E" + row + ":F" + row].Merge = true;
                    ws.Cells["G" + row + ":H" + row].Merge = true;
                    xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                    xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetFontBold(ws, "A" + row + ":B" + row);
                    xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetHorizontalAlignCenter(ws, "A" + row + ":H" + row);
                    row++;


                    ws.Cells["A" + row].Value = "Review Date:";
                    ws.Cells["C" + row].Value = questionnaire.DateOfReviewed != null ? questionnaire.DateOfReviewed.Value.ToString("MM/dd/yyyy") : string.Empty;
                    ws.Cells["E" + row].Value = string.Empty;
                    ws.Cells["G" + row].Value = string.Empty;
                    ws.Cells["A" + row + ":B" + row].Merge = true;
                    ws.Cells["C" + row + ":D" + row].Merge = true;
                    ws.Cells["E" + row + ":F" + row].Merge = true;
                    ws.Cells["G" + row + ":H" + row].Merge = true;
                    xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                    xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetFontBold(ws, "A" + row + ":B" + row);
                    xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetHorizontalAlignCenter(ws, "A" + row + ":H" + row);
                    row++;


                    ws.Cells["A" + row].Value = "Test Findings Assessment:";
                    ws.Cells["C" + row].Value = questionnaire.ReviewersNote != null ? questionnaire.ReviewersNote : string.Empty;
                    ws.Cells["E" + row].Value = string.Empty;
                    ws.Cells["G" + row].Value = string.Empty;
                    ws.Cells["A" + row + ":B" + row].Merge = true;
                    ws.Cells["C" + row + ":D" + row].Merge = true;
                    ws.Cells["E" + row + ":F" + row].Merge = true;
                    ws.Cells["G" + row + ":H" + row].Merge = true;
                    xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                    xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetFontBold(ws, "A" + row + ":B" + row);
                    xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetHorizontalAlignCenter(ws, "A" + row + ":H" + row);
                    row++;

                    #region Legends

                    row++;
                    ws.Cells["A" + row].Value = "Legend";
                    ws.Cells["A" + row + ":B" + row].Merge = true;
                    xlsService.ExcelWrapText(ws, "A" + row + ":B" + row);
                    xlsService.ExcelSetBorder(ws, "A" + row + ":B" + row);
                    xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                    xlsService.ExcelSetArialSize10(ws, "A" + row + ":B" + row);
                    xlsService.ExcelSetFontBold(ws, "A" + row + ":B" + row);
                    row++;
                    if (questionnaire.ListNotes != null && questionnaire.ListNotes.Count > 0)
                    {
                        foreach (var item in questionnaire.ListNotes)
                        {
                            ws.Cells["A" + row].Value = item.Key != null ? item.Key : string.Empty;
                            ws.Cells["C" + row].Value = item.Description != null ? item.Description : string.Empty;
                            ws.Cells["A" + row + ":B" + row].Merge = true;
                            ws.Cells["C" + row + ":H" + row].Merge = true;
                            xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetFontBold(ws, "A" + row + ":B" + row);
                            xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetHorizontalAlignCenter(ws, "A" + row + ":B" + row);
                            ws.Row(row).Height = 30;
                            row++;
                        }

                    }

                    #endregion


                    string startupPath = Directory.GetCurrentDirectory();
                    string strSourceDownload = Path.Combine(startupPath, "include", "upload", "soxquestionnaire");
                    //string strSourceDownload = startupPath + "\\include\\questionnaire\\download\\"; 0323

                    if (!Directory.Exists(strSourceDownload))
                    {
                        Directory.CreateDirectory(strSourceDownload);
                    }
                    var ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    string filename = $"{Guid.NewGuid()}.xlsx";
                    string strOutput = Path.Combine(strSourceDownload, filename);

                    //Check if file not exists
                    if (System.IO.File.Exists(strOutput))
                    {
                        System.IO.File.Delete(strOutput);
                    }

                    xls.SaveAs(new FileInfo(strOutput));
                    excelFilename = filename;
                }

            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error CreateExcel {ex}", "ErrorCreateExcel");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "CreateExcel");
            }

            return excelFilename;
        }

        [HttpPost("excel/cta/create")]
        public string CreateExcelXlsm([FromBody] Questionnaire questionnaire)
        {
            //List<string> excelFilename = new List<string>();
            string excelFilename = string.Empty;

            try
            {

                string startupPath = Directory.GetCurrentDirectory();
                string path = Path.Combine(startupPath, "include", "Questionnaire", "SOX Package CTA.xlsm");

                var fi = new FileInfo(path);

                //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage xls = new ExcelPackage(fi))
                {

                    //if(xls.Workbook.VbaProject == null)
                    //{
                    //    xls.Workbook.CreateVBAProject();
                    //}
                    //var ws = xls.Workbook.Worksheets.Add(questionnaire.ControlId);
                    //var cleanTabName = questionnaire.ControlId.Replace(" ", "_").Replace(".","_");
                    var ws = xls.Workbook.Worksheets.Add(questionnaire.ControlId);

                    ws.View.ShowGridLines = false;

                    ExcelService xlsService = new ExcelService();
                    ws.Column(1).Width = 30;
                    ws.Column(2).Width = 30;
                    ws.Column(3).Width = 30;
                    ws.Column(4).Width = 30;
                    ws.Column(5).Width = 30;
                    ws.Column(6).Width = 30;
                    ws.Column(7).Width = 30;
                    ws.Column(8).Width = 30;

                    ws.Cells["A6"].Value = "Process:";
                    ws.Cells["C6"].Value = questionnaire.Process != null ? questionnaire.Process : string.Empty;
                    ws.Cells["A6:B6"].Merge = true;
                    ws.Cells["C6:H6"].Merge = true;
                    xlsService.ExcelWrapText(ws, "A6:H6");
                    xlsService.ExcelSetBorder(ws, "A6:H6");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A6:B6");
                    xlsService.ExcelSetArialSize10(ws, "A6:H6");
                    xlsService.ExcelSetFontBold(ws, "A6");
                  //  xlsService.ExcelSetBackgroundColorGray(ws, "6:H6");

                    ws.Cells["A7"].Value = "Control Owner:";
                    ws.Cells["C7"].Value = questionnaire.ControlOwner != null ? questionnaire.ControlOwner : string.Empty;
                    ws.Cells["A7:B7"].Merge = true;
                    ws.Cells["C7:H7"].Merge = true;
                    xlsService.ExcelWrapText(ws, "A7:H7");
                    xlsService.ExcelSetBorder(ws, "A7:H7");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A7:B7");
                    xlsService.ExcelSetArialSize10(ws, "A7:H7");
                    xlsService.ExcelSetFontBold(ws, "A7");

                    ws.Cells["A8"].Value = "Control Id:";
                    ws.Cells["C8"].Value = questionnaire.ControlId != null ? questionnaire.ControlId : string.Empty;
                    ws.Cells["A8:B8"].Merge = true;
                    ws.Cells["C8:H8"].Merge = true;
                    xlsService.ExcelWrapText(ws, "A8:H8");
                    xlsService.ExcelSetBorder(ws, "A8:H8");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A8:B8");
                    xlsService.ExcelSetArialSize10(ws, "A8:H8");
                    xlsService.ExcelSetFontBold(ws, "A8");

                    ws.Cells["A9"].Value = "Control Description:";
                    ws.Cells["C9"].Value = questionnaire.ControlActivity != null ? questionnaire.ControlActivity : string.Empty;
                    ws.Cells["A9:B9"].Merge = true;
                    ws.Cells["C9:H9"].Merge = true;
                    //trim
                  
                    xlsService.ExcelWrapText(ws, "A9:H9");
                    xlsService.ExcelSetBorder(ws, "A9:H9");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A9:B9");
                    xlsService.ExcelSetArialSize10(ws, "A9:H9");
                    xlsService.ExcelSetFontBold(ws, "A9");
                    xlsService.ExcelSetVerticalAlignCenter(ws, "A9:H9");
                    ws.Row(9).Height = 60;

                    ws.Cells["A11"].Value = "Test Validation Approach:";
                    ws.Cells["C11"].Value = questionnaire.TestValidation != null ? questionnaire.TestValidation : string.Empty;
                    ws.Cells["A11:B11"].Merge = true;
                    ws.Cells["C11:H11"].Merge = true;
                    xlsService.ExcelWrapText(ws, "A10:H10");
                    xlsService.ExcelWrapText(ws, "A11:H11");
                    xlsService.ExcelSetBorder(ws, "A11:H11");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A11:B11");
                    xlsService.ExcelSetArialSize10(ws, "A11:H11");
                    xlsService.ExcelSetFontBold(ws, "A11");
                    xlsService.ExcelSetVerticalAlignCenter(ws, "A11:H11");
                    ws.Row(10).Height = 60;
                    ws.Row(11).Height = 60;

                    ws.Cells["A12"].Value = "Test Method Used:";
                    ws.Cells["C12"].Value = questionnaire.MethodUsed != null ? questionnaire.MethodUsed : string.Empty;
                    ws.Cells["A12:B12"].Merge = true;
                    ws.Cells["C12:H12"].Merge = true;
                    xlsService.ExcelWrapText(ws, "A12:H12");
                    xlsService.ExcelSetBorder(ws, "A12:H12");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A12:B12");
                    xlsService.ExcelSetArialSize10(ws, "A12:H12");
                    xlsService.ExcelSetFontBold(ws, "A12");


                    ws.Cells["A13"].Value = "Control Frequency:";
                    ws.Cells["C13"].Value = questionnaire.ControlFrequency != null ? questionnaire.ControlFrequency : string.Empty;
                    ws.Cells["A13:B13"].Merge = true;
                    ws.Cells["C13:H13"].Merge = true;
                    xlsService.ExcelWrapText(ws, "A13:H13");
                    xlsService.ExcelSetBorder(ws, "A13:H13");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A13:B13");
                    xlsService.ExcelSetArialSize10(ws, "A13:H13");
                    xlsService.ExcelSetFontBold(ws, "A13");


                    ws.Cells["A14"].Value = "Control In Place Date:";
                    ws.Cells["C14"].Value = questionnaire.ControlPlaceDate != null ? questionnaire.ControlPlaceDate : string.Empty;
                    ws.Cells["A14:B14"].Merge = true;
                    ws.Cells["C14:H14"].Merge = true;
                    xlsService.ExcelWrapText(ws, "A14:H14");
                    xlsService.ExcelSetBorder(ws, "A14:H14");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A14:B14");
                    xlsService.ExcelSetArialSize10(ws, "A14:H14");
                    xlsService.ExcelSetFontBold(ws, "A14");


                    ws.Cells["A15"].Value = "Risk Assessment:";
                    ws.Cells["C15"].Value = questionnaire.RiskAssessment != null ? questionnaire.RiskAssessment : string.Empty;
                    ws.Cells["A15:B15"].Merge = true;
                    ws.Cells["C15:H15"].Merge = true;
                    xlsService.ExcelWrapText(ws, "A15:H15");
                    xlsService.ExcelSetBorder(ws, "A15:H15");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A15:B15");
                    xlsService.ExcelSetArialSize10(ws, "A15:H15");
                    xlsService.ExcelSetFontBold(ws, "A15");


                    ws.Cells["A16"].Value = "Sample Period:";
                    ws.Cells["C16"].Value = questionnaire.SamplePeriod != null ? questionnaire.SamplePeriod : string.Empty;
                    ws.Cells["A16:B16"].Merge = true;
                    ws.Cells["C16:H16"].Merge = true;
                    xlsService.ExcelWrapText(ws, "A16:H16");
                    xlsService.ExcelSetBorder(ws, "A16:H16");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A16:B16");
                    xlsService.ExcelSetArialSize10(ws, "A16:H16");
                    xlsService.ExcelSetFontBold(ws, "A16");


                    ws.Cells["A17"].Value = "Population Size:";
                    ws.Cells["C17"].Value = questionnaire.PopulationSize != null ? questionnaire.PopulationSize : string.Empty;
                    ws.Cells["A17:B17"].Merge = true;
                    ws.Cells["C17:H17"].Merge = true;
                    xlsService.ExcelWrapText(ws, "A17:H17");
                    xlsService.ExcelSetBorder(ws, "A17:H17");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A17:B17");
                    xlsService.ExcelSetArialSize10(ws, "A17:H17");
                    xlsService.ExcelSetFontBold(ws, "A17");


                    ws.Cells["A18"].Value = "Sample Derivation:";
                    ws.Cells["C18"].Value = questionnaire.SampleSizeDerivation != null ? questionnaire.SampleSizeDerivation : string.Empty;
                    ws.Cells["A18:B18"].Merge = true;
                    ws.Cells["C18:H18"].Merge = true;
                    xlsService.ExcelWrapText(ws, "A18:H18");
                    xlsService.ExcelSetBorder(ws, "A18:H18");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A18:B18");
                    xlsService.ExcelSetArialSize10(ws, "A18:H18");
                    xlsService.ExcelSetFontBold(ws, "A18");


                    ws.Cells["A19"].Value = "Electronic Audit Evidence:";
                    ws.Cells["C19"].Value = questionnaire.IpeInformation != null ? questionnaire.IpeInformation : string.Empty;
                    ws.Cells["A19:B19"].Merge = true;
                    ws.Cells["C19:H19"].Merge = true;
                    xlsService.ExcelWrapText(ws, "A19:H19");
                    xlsService.ExcelSetBorder(ws, "A19:H19");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A19:B19");
                    xlsService.ExcelSetArialSize10(ws, "A19:H19");
                    xlsService.ExcelSetFontBold(ws, "A19");


                    #region Round Header
                    ws.Cells["C21"].Value = "Round 1";
                    ws.Cells["C21:D21"].Merge = true;
                    xlsService.ExcelWrapText(ws, "C21:D21");
                    xlsService.ExcelSetBorder(ws, "C21:D21");

                    ws.Cells["E21"].Value = "Round 2";
                    ws.Cells["E21:F21"].Merge = true;
                    xlsService.ExcelWrapText(ws, "E21:F21");
                    xlsService.ExcelSetBorder(ws, "E21:F21");

                    ws.Cells["G21"].Value = "Round 3";
                    ws.Cells["G21:H21"].Merge = true;
                    xlsService.ExcelWrapText(ws, "G21:H21");
                    xlsService.ExcelSetBorder(ws, "G21:H21");

                    xlsService.ExcelSetBackgroundColorGray(ws, "C21:H21");
                    xlsService.ExcelSetArialSize10(ws, "C21:H21");
                    xlsService.ExcelSetFontBold(ws, "C21:H21");
                    xlsService.ExcelSetVerticalAlignCenter(ws, "C21:H21");
                    xlsService.ExcelSetHorizontalAlignCenter(ws, "C21:H21");
                    #endregion

                    ws.Cells["A22"].Value = "Sample Size:";
                    ws.Cells["C22"].Value = questionnaire.SampleSizeRound1;
                    ws.Cells["E22"].Value = questionnaire.SampleSizeRound2;
                    ws.Cells["G22"].Value = questionnaire.SampleSizeRound3;
                    ws.Cells["A22:B22"].Merge = true;
                    ws.Cells["C22:D22"].Merge = true;
                    ws.Cells["E22:F22"].Merge = true;
                    ws.Cells["G22:H22"].Merge = true;
                    xlsService.ExcelWrapText(ws, "A22:H22");
                    xlsService.ExcelSetBorder(ws, "A22:H22");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A22:B22");
                    xlsService.ExcelSetArialSize10(ws, "A22:H22");
                    xlsService.ExcelSetFontBold(ws, "A22");
                    xlsService.ExcelSetVerticalAlignCenter(ws, "C22:H22");
                    xlsService.ExcelSetHorizontalAlignCenter(ws, "C22:H22");


                    ws.Cells["A23"].Value = "Source File (document name, hardcopy/softcopy & provided by):";
                    ws.Cells["C23"].Value = questionnaire.SourceFileType;
                    ws.Cells["E23"].Value = string.Empty;
                    ws.Cells["G23"].Value = string.Empty;
                    ws.Cells["A23:B23"].Merge = true;
                    ws.Cells["C23:D23"].Merge = true;
                    ws.Cells["E23:F23"].Merge = true;
                    ws.Cells["G23:H23"].Merge = true;
                    xlsService.ExcelWrapText(ws, "A23:H23");
                    xlsService.ExcelSetBorder(ws, "A23:H23");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A23:B23");
                    xlsService.ExcelSetArialSize10(ws, "A23:H23");
                    xlsService.ExcelSetFontBold(ws, "A23");
                    xlsService.ExcelSetVerticalAlignCenter(ws, "A23:H23");
                    ws.Row(23).Height = 320;


                    ws.Cells["A24"].Value = "Test Performed By:";
                    ws.Cells["C24"].Value = questionnaire.TestPerfomedBy;
                    ws.Cells["E24"].Value = string.Empty;
                    ws.Cells["G24"].Value = string.Empty;
                    ws.Cells["A24:B24"].Merge = true;
                    ws.Cells["C24:D24"].Merge = true;
                    ws.Cells["E24:F24"].Merge = true;
                    ws.Cells["G24:H24"].Merge = true;
                    xlsService.ExcelWrapText(ws, "A24:H24");
                    xlsService.ExcelSetBorder(ws, "A24:H24");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A24:B24");
                    xlsService.ExcelSetArialSize10(ws, "A24:H24");
                    xlsService.ExcelSetFontBold(ws, "A24");
                    xlsService.ExcelSetVerticalAlignCenter(ws, "A24:H24");
                    xlsService.ExcelSetHorizontalAlignCenter(ws, "C24:H24");


                    ws.Cells["A25"].Value = "Date Testing Performed:";
                    ws.Cells["C25"].Value = questionnaire.DateOfTesting != null ? questionnaire.DateOfTesting.Value.ToString("MM/dd/yyyy") : string.Empty;
                    ws.Cells["E25"].Value = string.Empty;
                    ws.Cells["G25"].Value = string.Empty;
                    ws.Cells["A25:B25"].Merge = true;
                    ws.Cells["C25:D25"].Merge = true;
                    ws.Cells["E25:F25"].Merge = true;
                    ws.Cells["G25:H25"].Merge = true;
                    xlsService.ExcelWrapText(ws, "A25:H25");
                    xlsService.ExcelSetBorder(ws, "A25:H25");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A25:B25");
                    xlsService.ExcelSetArialSize10(ws, "A25:H25");
                    xlsService.ExcelSetFontBold(ws, "A25");
                    xlsService.ExcelSetVerticalAlignCenter(ws, "A25:H25");
                    xlsService.ExcelSetHorizontalAlignCenter(ws, "C25:H25");


                    ws.Cells["A26"].Value = "Test of Design (TOD):";
                    ws.Cells["C26"].Value = questionnaire.TestOfDesign != null ? questionnaire.TestOfDesign : string.Empty;
                    ws.Cells["E26"].Value = string.Empty;
                    ws.Cells["G26"].Value = string.Empty;
                    ws.Cells["A26:B26"].Merge = true;
                    ws.Cells["C26:D26"].Merge = true;
                    ws.Cells["E26:F26"].Merge = true;
                    ws.Cells["G26:H26"].Merge = true;
                    xlsService.ExcelWrapText(ws, "A26:H26");
                    xlsService.ExcelSetBorder(ws, "A26:H26");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A26:B26");
                    xlsService.ExcelSetArialSize10(ws, "A26:H26");
                    xlsService.ExcelSetFontBold(ws, "A26");
                    xlsService.ExcelSetVerticalAlignCenter(ws, "A26:H26");
                    xlsService.ExcelSetHorizontalAlignCenter(ws, "C26:H26");


                    ws.Cells["A27"].Value = "Test of Operating Effectiveness (TOE):";
                    ws.Cells["C27"].Value = questionnaire.TestOperatingEffectiveness != null ? questionnaire.TestOperatingEffectiveness : string.Empty;
                    ws.Cells["E27"].Value = string.Empty;
                    ws.Cells["G27"].Value = string.Empty;
                    ws.Cells["A27:B27"].Merge = true;
                    ws.Cells["C27:D27"].Merge = true;
                    ws.Cells["E27:F27"].Merge = true;
                    ws.Cells["G27:H27"].Merge = true;
                    xlsService.ExcelWrapText(ws, "A27:H27");
                    xlsService.ExcelSetBorder(ws, "A27:H27");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A27:B27");
                    xlsService.ExcelSetArialSize10(ws, "A27:H27");
                    xlsService.ExcelSetFontBold(ws, "A27");
                    xlsService.ExcelSetVerticalAlignCenter(ws, "A27:H27");
                    xlsService.ExcelSetHorizontalAlignCenter(ws, "C27:H27");


                    #region Round Sample

                    ws.Cells["A29"].Value = "Sample #";
                    ws.Cells["B29"].Value = "Period";
                    ws.Cells["C29"].Value = "Any investigation?";
                    ws.Cells["D29"].Value = "Medrio costs agree to the Clinical trial agreements [A]";
                    ws.Cells["E29"].Value = "Evidence of review";
                    ws.Cells["F29"].Value = "Reviewer";
                    ws.Cells["G29"].Value = "Review date";
                    ws.Cells["H29"].Value = "Notes";
                    xlsService.ExcelWrapText(ws, "A29:H29");
                    xlsService.ExcelSetBorder(ws, "A29:H29");
                    xlsService.ExcelSetBackgroundColorGray(ws, "A29:H29");
                    xlsService.ExcelSetArialSize10(ws, "A29:H29");
                    xlsService.ExcelSetFontBold(ws, "A29:H29");
                    xlsService.ExcelSetVerticalAlignCenter(ws, "A29:H29");
                    xlsService.ExcelSetHorizontalAlignCenter(ws, "A29:H29");

                    int row = 30;

                    //Round 1
                    ws.Cells["A" + row].Value = "Round 1";
                    ws.Cells["A" + row + ":H" + row].Merge = true;
                    xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetBackgroundColorLightGray(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetFontBold(ws, "A" + row + ":H" + row);
                    row++;

                    var listRound1 = questionnaire.ListSampleRound.Where(x => x.RoundName == "Round 1");
                    if (listRound1 != null && listRound1.Count() > 0)
                    {
                        foreach (var item in listRound1)
                        {
                            ws.Cells["A" + row].Value = $"{(item.SampleNum != null ? item.SampleNum.ToString() : string.Empty)}";
                            ws.Cells["B" + row].Value = $"{(item.MedrioActivity16A != null ? item.MedrioActivity16A.ToString() : string.Empty)} [{item.TickBox1Value}]";
                            ws.Cells["C" + row].Value = $"{(item.InvestigationPerformed16B != null ? item.InvestigationPerformed16B.ToString() : string.Empty)} [{item.TickBox2Value}]";
                            ws.Cells["D" + row].Value = $"{(item.MedrioCostAgreements16C != null ? item.MedrioCostAgreements16C.ToString() : string.Empty)} [{item.TickBox3Value}]";
                            ws.Cells["E" + row].Value = $"{(item.MedrioReportReview16D != null ? item.MedrioReportReview16D.ToString() : string.Empty)} [{item.TickBox4Value}]";
                            ws.Cells["F" + row].Value = $"{(item.MedrioReviewers16E != null ? item.MedrioReviewers16E.ToString() : string.Empty)} [{item.TickBox5Value}]";
                            ws.Cells["G" + row].Value = $"{(item.DateMedrioReviewed16F != null ? item.DateMedrioReviewed16F.Value.ToString("MM/dd/yyyy") : string.Empty)} [{item.TickBox6Value}]";
                            ws.Cells["H" + row].Value = item.Notes;
                            xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetHorizontalAlignCenter(ws, "A" + row + ":H" + row);
                            ws.Row(row).Height = 30;
                            row++;
                        }
                    }

                    //Round 2
                    ws.Cells["A" + row].Value = "Round 2";
                    ws.Cells["A" + row + ":H" + row].Merge = true;
                    xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetBackgroundColorLightGray(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetFontBold(ws, "A" + row + ":H" + row);
                    row++;

                    var listRound2 = questionnaire.ListSampleRound.Where(x => x.RoundName == "Round 2");
                    if (listRound2 != null && listRound2.Count() > 0)
                    {
                        foreach (var item in listRound2)
                        {
                            ws.Cells["A" + row].Value = $"{(item.SampleNum != null ? item.SampleNum.ToString() : string.Empty)}";
                            ws.Cells["B" + row].Value = $"{(item.MedrioActivity16A != null ? item.MedrioActivity16A.ToString() : string.Empty)} [{item.TickBox1Value}]";
                            ws.Cells["C" + row].Value = $"{(item.InvestigationPerformed16B != null ? item.InvestigationPerformed16B.ToString() : string.Empty)} [{item.TickBox2Value}]";
                            ws.Cells["D" + row].Value = $"{(item.MedrioCostAgreements16C != null ? item.MedrioCostAgreements16C.ToString() : string.Empty)} [{item.TickBox3Value}]";
                            ws.Cells["E" + row].Value = $"{(item.MedrioReportReview16D != null ? item.MedrioReportReview16D.ToString() : string.Empty)} [{item.TickBox4Value}]";
                            ws.Cells["F" + row].Value = $"{(item.MedrioReviewers16E != null ? item.MedrioReviewers16E.ToString() : string.Empty)} [{item.TickBox5Value}]";
                            ws.Cells["G" + row].Value = $"{(item.DateMedrioReviewed16F != null ? item.DateMedrioReviewed16F.Value.ToString("MM/dd/yyyy") : string.Empty)} [{item.TickBox6Value}]";
                            ws.Cells["H" + row].Value = item.Notes;
                            xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetHorizontalAlignCenter(ws, "A" + row + ":H" + row);
                            ws.Row(row).Height = 30;
                            row++;
                        }
                    }

                    //Round 3
                    ws.Cells["A" + row].Value = "Round 3";
                    ws.Cells["A" + row + ":H" + row].Merge = true;
                    xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetBackgroundColorLightGray(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetFontBold(ws, "A" + row + ":H" + row);
                    row++;

                    var listRound3 = questionnaire.ListSampleRound.Where(x => x.RoundName == "Round 3");
                    if (listRound3 != null && listRound3.Count() > 0)
                    {
                        foreach (var item in listRound3)
                        {
                            ws.Cells["A" + row].Value = $"{(item.SampleNum != null ? item.SampleNum.ToString() : string.Empty)}";
                            ws.Cells["B" + row].Value = $"{(item.MedrioActivity16A != null ? item.MedrioActivity16A.ToString() : string.Empty)} [{item.TickBox1Value}]";
                            ws.Cells["C" + row].Value = $"{(item.InvestigationPerformed16B != null ? item.InvestigationPerformed16B.ToString() : string.Empty)} [{item.TickBox2Value}]";
                            ws.Cells["D" + row].Value = $"{(item.MedrioCostAgreements16C != null ? item.MedrioCostAgreements16C.ToString() : string.Empty)} [{item.TickBox3Value}]";
                            ws.Cells["E" + row].Value = $"{(item.MedrioReportReview16D != null ? item.MedrioReportReview16D.ToString() : string.Empty)} [{item.TickBox4Value}]";
                            ws.Cells["F" + row].Value = $"{(item.MedrioReviewers16E != null ? item.MedrioReviewers16E.ToString() : string.Empty)} [{item.TickBox5Value}]";
                            ws.Cells["G" + row].Value = $"{(item.DateMedrioReviewed16F != null ? item.DateMedrioReviewed16F.Value.ToString("MM/dd/yyyy") : string.Empty)} [{item.TickBox6Value}]";
                            ws.Cells["H" + row].Value = item.Notes;
                            xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetHorizontalAlignCenter(ws, "A" + row + ":H" + row);
                            ws.Row(row).Height = 30;
                            row++;
                        }
                    }

                    #endregion 


                    #region Round Header
                    row++;
                    ws.Cells["C" + row].Value = "Round 1";
                    ws.Cells["C" + row + ":D" + row].Merge = true;
                    xlsService.ExcelWrapText(ws, "C" + row + ":D" + row);
                    xlsService.ExcelSetBorder(ws, "C" + row + ":D" + row);

                    ws.Cells["E" + row].Value = "Round 2";
                    ws.Cells["E" + row + ":F" + row].Merge = true;
                    xlsService.ExcelWrapText(ws, "E" + row + ":F" + row);
                    xlsService.ExcelSetBorder(ws, "E" + row + ":F" + row);

                    ws.Cells["G" + row].Value = "Round 3";
                    ws.Cells["G" + row + ":H" + row].Merge = true;
                    xlsService.ExcelWrapText(ws, "G" + row + ":H" + row);
                    xlsService.ExcelSetBorder(ws, "G" + row + ":H" + row);

                    xlsService.ExcelSetBackgroundColorGray(ws, "C" + row + ":H" + row);
                    xlsService.ExcelSetArialSize10(ws, "C" + row + ":H" + row);
                    xlsService.ExcelSetFontBold(ws, "C" + row + ":H" + row);
                    xlsService.ExcelSetVerticalAlignCenter(ws, "C" + row + ":H" + row);
                    xlsService.ExcelSetHorizontalAlignCenter(ws, "C" + row + ":H" + row);
                    row++;
                    #endregion


                    ws.Cells["A" + row].Value = "Reviewer:";
                    ws.Cells["C" + row].Value = questionnaire.ReviewedBy != null ? questionnaire.ReviewedBy : string.Empty;
                    ws.Cells["E" + row].Value = string.Empty;
                    ws.Cells["G" + row].Value = string.Empty;
                    ws.Cells["A" + row + ":B" + row].Merge = true;
                    ws.Cells["C" + row + ":D" + row].Merge = true;
                    ws.Cells["E" + row + ":F" + row].Merge = true;
                    ws.Cells["G" + row + ":H" + row].Merge = true;
                    xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                    xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetFontBold(ws, "A" + row + ":B" + row);
                    xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetHorizontalAlignCenter(ws, "A" + row + ":H" + row);
                    row++;


                    ws.Cells["A" + row].Value = "Review Date:";
                    ws.Cells["C" + row].Value = questionnaire.DateOfReviewed != null ? questionnaire.DateOfReviewed.Value.ToString("MM/dd/yyyy") : string.Empty;
                    ws.Cells["E" + row].Value = string.Empty;
                    ws.Cells["G" + row].Value = string.Empty;
                    ws.Cells["A" + row + ":B" + row].Merge = true;
                    ws.Cells["C" + row + ":D" + row].Merge = true;
                    ws.Cells["E" + row + ":F" + row].Merge = true;
                    ws.Cells["G" + row + ":H" + row].Merge = true;
                    xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                    xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetFontBold(ws, "A" + row + ":B" + row);
                    xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetHorizontalAlignCenter(ws, "A" + row + ":H" + row);
                    row++;


                    ws.Cells["A" + row].Value = "Test Findings Assessment:";
                    ws.Cells["C" + row].Value = questionnaire.ReviewersNote != null ? questionnaire.ReviewersNote : string.Empty;
                    ws.Cells["E" + row].Value = string.Empty;
                    ws.Cells["G" + row].Value = string.Empty;
                    ws.Cells["A" + row + ":B" + row].Merge = true;
                    ws.Cells["C" + row + ":D" + row].Merge = true;
                    ws.Cells["E" + row + ":F" + row].Merge = true;
                    ws.Cells["G" + row + ":H" + row].Merge = true;
                    xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                    xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetFontBold(ws, "A" + row + ":B" + row);
                    xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":H" + row);
                    xlsService.ExcelSetHorizontalAlignCenter(ws, "A" + row + ":H" + row);
                    row++;

                    #region Legends

                    row++;
                    ws.Cells["A" + row].Value = "Legend";
                    ws.Cells["A" + row + ":B" + row].Merge = true;
                    xlsService.ExcelWrapText(ws, "A" + row + ":B" + row);
                    xlsService.ExcelSetBorder(ws, "A" + row + ":B" + row);
                    xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                    xlsService.ExcelSetArialSize10(ws, "A" + row + ":B" + row);
                    xlsService.ExcelSetFontBold(ws, "A" + row + ":B" + row);
                    row++;
                    if (questionnaire.ListNotes != null && questionnaire.ListNotes.Count > 0)
                    {
                        foreach (var item in questionnaire.ListNotes)
                        {
                            ws.Cells["A" + row].Value = item.Key != null ? item.Key : string.Empty;
                            ws.Cells["C" + row].Value = item.Description != null ? item.Description : string.Empty;
                            ws.Cells["A" + row + ":B" + row].Merge = true;
                            ws.Cells["C" + row + ":H" + row].Merge = true;
                            xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetFontBold(ws, "A" + row + ":B" + row);
                            xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetHorizontalAlignCenter(ws, "A" + row + ":B" + row);
                            ws.Row(row).Height = 30;
                            row++;
                        }

                    }

                    #endregion

                    ws.Protection.IsProtected = true;
                    ws.Protection.AllowAutoFilter = false;
                    ws.Protection.AllowDeleteColumns = false;
                    ws.Protection.AllowDeleteRows = false;
                    ws.Protection.AllowEditObject = true;
                    ws.Protection.AllowEditScenarios = true;
                    ws.Protection.AllowFormatCells = false;
                    ws.Protection.AllowFormatColumns = false;
                    ws.Protection.AllowFormatRows = false;
                    ws.Protection.AllowInsertColumns = false;
                    ws.Protection.AllowInsertHyperlinks = false;
                    ws.Protection.AllowInsertRows = false;
                    ws.Protection.AllowPivotTables = false;
                    ws.Protection.AllowSelectLockedCells = true;
                    ws.Protection.AllowSelectUnlockedCells = true;
                    ws.Protection.AllowSort = false;
                    ws.Protection.SetPassword("123qwe!!");

                    //string startupPath = Directory.GetCurrentDirectory();
                    string strSourceDownload = Path.Combine(startupPath, "include", "upload", "soxquestionnaire");
                    //string strSourceDownload = startupPath + "\\include\\questionnaire\\download\\";032411

                    if (!Directory.Exists(strSourceDownload))
                    {
                        Directory.CreateDirectory(strSourceDownload);
                    }
                    var ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    string filename = $"{Guid.NewGuid()}.xlsm";
                    string strOutput = Path.Combine(strSourceDownload, filename);

                    //Check if file not exists
                    if (System.IO.File.Exists(strOutput))
                    {
                        System.IO.File.Delete(strOutput);
                    }

                    xls.SaveAs(new FileInfo(strOutput));
                    excelFilename = filename;

                }

            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error CreateExcelXlsm {ex}", "ErrorCreateExcelXlsm");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "CreateExcelXlsm");
            }

            return excelFilename;
        }

        //Old questionnaire excel generation
        [HttpPost("excel/elc/create")]
        public string CreateExcelElcXlsm([FromBody] QuestionnaireExcelData questionnaireInput)
        {
            //List<string> excelFilename = new List<string>();
            string excelFilename = string.Empty;

            try
            {

                string startupPath = Directory.GetCurrentDirectory();
                //string path = Path.Combine(startupPath, "include", "questionnaire", "SOX Package CTA.xlsm");

                //var fi = new FileInfo(path);

                //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage xls = new ExcelPackage())
                {

                    //if(xls.Workbook.VbaProject == null)
                    //{
                    //    xls.Workbook.CreateVBAProject();
                    //}
                    //var ws = xls.Workbook.Worksheets.Add(questionnaire.ControlId);
                    //var cleanTabName = questionnaire.ControlId.Replace(" ", "_").Replace(".","_");

                    //var controlName = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("control id")).Select(x => x.StrAnswer).FirstOrDefault();
                    var controlName = questionnaireInput.Rcm.ControlId;
                    if (controlName != null)
                    {
                        var ws = xls.Workbook.Worksheets.Add(controlName);

                        ws.View.ShowGridLines = false;

                        ExcelService xlsService = new ExcelService();
                        ws.Column(1).Width = 30;
                        ws.Column(2).Width = 30;
                        ws.Column(3).Width = 30;
                        ws.Column(4).Width = 30;
                        ws.Column(5).Width = 30;
                        ws.Column(6).Width = 30;
                        ws.Column(7).Width = 30;
                        ws.Column(8).Width = 30;
                        ws.Column(9).Width = 30;
                        ws.Column(10).Width = 30;
                        ws.Column(11).Width = 30;
                        ws.Column(12).Width = 30;
                        ws.Column(13).Width = 30;
                        ws.Column(14).Width = 30;
                        ws.Column(15).Width = 30;


                        int row = 1; //Row 1
                        string clientName = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("client name?")).Select(x => x.StrAnswer).FirstOrDefault();
                        ws.Cells["A" + row].Value = clientName != null ? clientName : string.Empty;
                        xlsService.ExcelSetArialSize10(ws, "A" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);

                        row++; //Row 2
                        string controlId = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("control id?")).Select(x => x.StrAnswer).FirstOrDefault();
                        ws.Cells["A" + row].Value = controlId != null ? controlId : string.Empty;
                        xlsService.ExcelSetArialSize10(ws, "A" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);

                        //Row 6
                        row += 4;
                        ws.Cells["A" + row].Value = "Process:";
                        ws.Cells["C" + row].Value = questionnaireInput.Rcm.Process != null ? questionnaireInput.Rcm.Process : string.Empty;
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":H" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);
                        //set excel row height base on lenght on text
                        xlsService.XlsSetRow(ws, row, questionnaireInput.Rcm.Process, 380, 15);

                        //Row 7
                        row++;
                        string controlOwner = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("control owner?")).Select(x => x.StrAnswer).FirstOrDefault();
                        ws.Cells["A" + row].Value = "Control Owner:";
                        ws.Cells["C" + row].Value = controlOwner != null ? controlOwner : string.Empty;
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":H" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);
                        //set excel row height base on lenght on text
                        xlsService.XlsSetRow(ws, row, controlOwner, 380, 15);

                        //Row 8
                        row++;
                        ws.Cells["A" + row].Value = "Control Id:";
                        ws.Cells["C" + row].Value = controlId != null ? controlId : string.Empty;
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":H" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);
                        //set excel row height base on lenght on text
                        xlsService.XlsSetRow(ws, row, controlId, 380, 15);


                        //Row 9
                        row++;
                        string controlDesc = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("what is the control activity?")).Select(x => x.StrAnswer).FirstOrDefault();
                        ws.Cells["A" + row].Value = "Control Description:";
                        ws.Cells["C" + row].Value = controlDesc != null ? controlDesc : string.Empty;
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":H" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);
                        xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":H" + row);
                        //ws.Row(row).Height = 60;

                        //set excel row height base on lenght on text
                        xlsService.XlsSetRow(ws, row, controlDesc, 300, 30);





                        //Row 11
                        row++;
                        string testValidation = questionnaireInput.ListUserInputItem.Where(x =>
                            x.StrQuestion.ToLower().Contains("what are the procedures to test control?") ||
                            x.StrQuestion.ToLower().Contains("what are the procedures to test this control")
                            ).Select(x => x.StrAnswer).FirstOrDefault();
                        ws.Cells["A" + row].Value = "Test Validation Approach:";
                        ws.Cells["C" + row].Value = testValidation != null ? testValidation : string.Empty;
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":H" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);
                        xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":H" + row);
                        //ws.Row(row).Height = 60;
                        //set excel row height base on lenght on text
                        xlsService.XlsSetRow(ws, row, testValidation, 300, 15);

                        //Row 12
                        //15. What procedures did you use to select the samples?
                        if (questionnaireInput.Rcm.TestProc != null && questionnaireInput.Rcm.TestProc != string.Empty)
                        {

                            //string methodUsedA = questionnaireInput.ListUserInputItem.Where(x =>
                            //    x.StrQuestion.ToLower().Contains("procedures did you use to select the samples?") ||
                            //    x.StrQuestion.ToLower().Contains("what is the test method used?")
                            //).Select(x => x.StrAnswer).FirstOrDefault();

                            row++;
                            ws.Cells["A" + row].Value = "Test Method Used:";
                            ws.Cells["C" + row].Value = questionnaireInput.Rcm.TestProc != null ? questionnaireInput.Rcm.TestProc : string.Empty;
                            ws.Cells["A" + row + ":B" + row].Merge = true;
                            ws.Cells["C" + row + ":H" + row].Merge = true;
                            xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                            xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetFontBold(ws, "A" + row);
                            //set excel row height base on lenght on text
                            xlsService.XlsSetRow(ws, row, questionnaireInput.Rcm.TestProc, 300, 30);
                        }


                        //Row 13
                        row++;
                        string controlFreq = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("how often does this control happen")).Select(x => x.StrAnswer).FirstOrDefault();
                        ws.Cells["A" + row].Value = "Control Frequency:";
                        ws.Cells["C" + row].Value = controlFreq != null ? controlFreq : string.Empty;
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":H" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);

                        //Row 14
                        //6. When was the control first put in place? Enter date like 1/1/xx.
                        row++;
                        string controlPlaceDate = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("control first put in place?")).Select(x => x.StrAnswer).FirstOrDefault();
                        string parseControlPlaceDate = string.Empty;
                        if (controlPlaceDate != null && controlPlaceDate != string.Empty)
                        {
                            DateTime tempcontrolPlaceDate;
                            //try to parse string to date time
                            if (DateTime.TryParse(controlPlaceDate, out tempcontrolPlaceDate))
                            {
                                parseControlPlaceDate = tempcontrolPlaceDate.ToString("MM/dd/yyyy");
                            }
                            else
                            {
                                //if it failed to parse date time, then we will just remove time string
                                int index = controlPlaceDate.IndexOf(" ");
                                if (index > 0)
                                {
                                    parseControlPlaceDate = controlPlaceDate.Substring(0, index);
                                }
                            }
                        }

                        ws.Cells["A" + row].Value = "Control In Place Date:";
                        ws.Cells["C" + row].Style.Numberformat.Format = "@";
                        ws.Cells["C" + row].Value = parseControlPlaceDate != null ? parseControlPlaceDate : string.Empty;
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":H" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);


                        //Row 15
                        row++;
                        string levelRisk = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("what is the level of risk for the control")).Select(x => x.StrAnswer).FirstOrDefault();
                        ws.Cells["A" + row].Value = "Risk Assessment:";
                        ws.Cells["C" + row].Value = levelRisk != null ? levelRisk : string.Empty;
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":H" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);

                        row++;
                        //Row 16
                        //12.What date ranges are you selecting samples?
                        string samplePeriod = questionnaireInput.ListUserInputItem.Where(x =>
                        //x.StrQuestion.ToLower().Contains("date range you are testing") || 
                        x.StrQuestion.ToLower().Contains("what date ranges are you selecting samples"))
                            .Select(x => x.StrAnswer).FirstOrDefault();
                        string samplePeriod2 = questionnaireInput.ListUserInputItem.Where(x =>
                        //x.StrQuestion.ToLower().Contains("date range you are testing") || 
                        x.StrQuestion.ToLower().Contains("what date ranges are you selecting samples"))
                            .Select(x => x.StrAnswer2).FirstOrDefault();
                        string parsesamplePeriod = string.Empty;
                        string parsesamplePeriod2 = string.Empty;
                        if (samplePeriod != null && samplePeriod != string.Empty)
                        {
                            //parsesamplePeriod = DateTime.Parse(samplePeriod).ToString("MM/dd/yyyy");
                            DateTime tempsamplePeriod;
                            //try to parse string to date time
                            if (DateTime.TryParse(samplePeriod, out tempsamplePeriod))
                            {
                                parsesamplePeriod = tempsamplePeriod.ToString("MM/dd/yyyy");
                            }
                            else
                            {
                                //if it failed to parse date time, then we will just remove time string
                                int index = samplePeriod.IndexOf(" ");
                                if (index > 0)
                                {
                                    parsesamplePeriod = samplePeriod.Substring(0, index);
                                }
                            }
                        }
                        if (samplePeriod2 != null && samplePeriod2 != string.Empty)
                        {
                            //parsesamplePeriod2 = DateTime.Parse(samplePeriod2).ToString("MM/dd/yyyy");
                            DateTime tempsamplePeriod2;
                            //try to parse string to date time
                            if (DateTime.TryParse(samplePeriod2, out tempsamplePeriod2))
                            {
                                parsesamplePeriod2 = tempsamplePeriod2.ToString("MM/dd/yyyy");
                            }
                            else
                            {
                                //if it failed to parse date time, then we will just remove time string
                                int index = samplePeriod2.IndexOf(" ");
                                if (index > 0)
                                {
                                    parsesamplePeriod2 = samplePeriod2.Substring(0, index);
                                }
                            }
                        }
                        string validateSamplePeriod = string.Empty;
                        if (parsesamplePeriod != null && parsesamplePeriod != string.Empty)
                        {
                            validateSamplePeriod = (parsesamplePeriod != null ? parsesamplePeriod : string.Empty) + " to " + (parsesamplePeriod2 != null ? parsesamplePeriod2 : string.Empty);
                        }

                        ws.Cells["A" + row].Value = "Sample Period:";
                        ws.Cells["C" + row].Style.Numberformat.Format = "@";
                        ws.Cells["C" + row].Value = validateSamplePeriod;
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":H" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);

                        row++;
                        //Row 17
                        string populationSize = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("what is the population size?")).Select(x => x.StrAnswer).FirstOrDefault();
                        ws.Cells["A" + row].Value = "Population Size:";
                        ws.Cells["C" + row].Value = populationSize != null ? populationSize : string.Empty;
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":H" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);

                        row++;
                        //Row 18
                        string sampleDerivation = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("what is the sub-sample size")).Select(x => x.StrAnswer).FirstOrDefault();
                        ws.Cells["A" + row].Value = "Sample Derivation:";
                        ws.Cells["C" + row].Value = sampleDerivation != null ? sampleDerivation : string.Empty;
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":H" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);

                        //dont display "Electronic Audit Evidence" if answer is blank
                        string auditEvidenceQ = questionnaireInput.ListUserInputItem.Where(x =>
                                x.StrQuestion.ToLower().Contains("(ipe)")
                            ).Select(x => x.StrQuestion).FirstOrDefault();
                        if (auditEvidenceQ != null && auditEvidenceQ != string.Empty)
                        {
                            row++;

                            string auditEvidenceA = questionnaireInput.ListUserInputItem.Where(x =>
                                x.StrQuestion.ToLower().Contains("(ipe)")
                            ).Select(x => x.StrAnswer).FirstOrDefault();

                            //Row 19
                            ws.Cells["A" + row].Value = "Electronic Audit Evidence:";
                            ws.Cells["C" + row].Value = auditEvidenceA != null ? auditEvidenceA : string.Empty;
                            ws.Cells["A" + row + ":B" + row].Merge = true;
                            ws.Cells["C" + row + ":H" + row].Merge = true;
                            xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                            xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetFontBold(ws, "A" + row);
                        }


                        row += 2;
                        //Row 21
                        #region Round Header
                        ws.Cells["C" + row].Value = "Round 1";
                        ws.Cells["C" + row + ":D" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "C" + row + ":D" + row);
                        xlsService.ExcelSetBorder(ws, "C" + row + ":D" + row);

                        ws.Cells["E" + row].Value = "Round 2";
                        ws.Cells["E" + row + ":F" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "E" + row + ":F" + row);
                        xlsService.ExcelSetBorder(ws, "E" + row + ":F" + row);

                        ws.Cells["G" + row].Value = "Round 3";
                        ws.Cells["G" + row + ":H" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "G" + row + ":H" + row);
                        xlsService.ExcelSetBorder(ws, "G" + row + ":H" + row);

                        xlsService.ExcelSetBackgroundColorGray(ws, "C" + row + ":H" + row);
                        xlsService.ExcelSetArialSize10(ws, "C" + row + ":H" + row);
                        xlsService.ExcelSetFontBold(ws, "C" + row + ":H" + row);
                        xlsService.ExcelSetVerticalAlignCenter(ws, "C" + row + ":H" + row);
                        xlsService.ExcelSetHorizontalAlignCenter(ws, "C" + row + ":H" + row);
                        #endregion


                        row++;
                        //Row 22
                        string roundName = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("select the testing phase")).Select(x => x.StrAnswer).FirstOrDefault();
                        string sampleSize = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("what is the sub-sample size") || x.StrQuestion.ToLower().Contains("what is the sub-sample size?")).Select(x => x.StrAnswer).FirstOrDefault();

                        ws.Cells["A" + row].Value = "Sample Size:";
                        ws.Cells["C" + row].Value = roundName == "Round 1" ? (sampleSize != null ? sampleSize : string.Empty) : string.Empty;
                        ws.Cells["E" + row].Value = roundName == "Round 2" ? (sampleSize != null ? sampleSize : string.Empty) : string.Empty;
                        ws.Cells["G" + row].Value = roundName == "Round 3" ? (sampleSize != null ? sampleSize : string.Empty) : string.Empty;
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":D" + row].Merge = true;
                        ws.Cells["E" + row + ":F" + row].Merge = true;
                        ws.Cells["G" + row + ":H" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);
                        xlsService.ExcelSetVerticalAlignCenter(ws, "C" + row + ":H" + row);
                        xlsService.ExcelSetHorizontalAlignCenter(ws, "C" + row + ":H" + row);

                        row++;
                        //Row 23
                        string str16A = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("source file")).Select(x => x.StrAnswer).FirstOrDefault();
                        //string str16A = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("16.a") || x.StrQuestion.ToLower().Contains("16. a")).Select(x => x.StrAnswer).FirstOrDefault();
                        string str16B = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("Who provided the document")).Select(x => x.StrAnswer).FirstOrDefault();
                        //string str16B = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("16.b") || x.StrQuestion.ToLower().Contains("16. b")).Select(x => x.StrAnswer).FirstOrDefault();
                        string str16C = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("List down the document")).Select(x => x.StrAnswer).FirstOrDefault();
                        //string str16C = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("16.c") || x.StrQuestion.ToLower().Contains("16. c")).Select(x => x.StrAnswer).FirstOrDefault();

                        string sourceFile = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("the source file")).Select(x => x.StrAnswer).FirstOrDefault();

                        sourceFile = (str16A != null && str16A != string.Empty ? str16A : string.Empty) + Environment.NewLine +
                            (str16B != null && str16A != string.Empty ? $" provided by {str16B}" : string.Empty) + Environment.NewLine +
                            (str16C != null ? str16C : string.Empty);
                        ws.Cells["A" + row].Value = "Source File (document name, hardcopy/softcopy & provided by):";
                        ws.Cells["C" + row].Value = roundName == "Round 1" ? (sourceFile != null ? sourceFile : string.Empty) : string.Empty;
                        ws.Cells["E" + row].Value = roundName == "Round 2" ? (sourceFile != null ? sourceFile : string.Empty) : string.Empty;
                        ws.Cells["G" + row].Value = roundName == "Round 3" ? (sourceFile != null ? sourceFile : string.Empty) : string.Empty;
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":D" + row].Merge = true;
                        ws.Cells["E" + row + ":F" + row].Merge = true;
                        ws.Cells["G" + row + ":H" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);
                        xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":H" + row);
                        //set excel row height base on lenght on text
                        xlsService.XlsSetRow(ws, row, sourceFile, 130, 150);
                        //ws.Row(row).Height = 320;

                        row++;
                        //Row 24
                        string testPerformed = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("performed the testing")).Select(x => x.StrAnswer).FirstOrDefault();
                        ws.Cells["A" + row].Value = "Test Performed By:";
                        ws.Cells["C" + row].Value = roundName == "Round 1" ? (testPerformed != null ? testPerformed : string.Empty) : string.Empty;
                        ws.Cells["E" + row].Value = roundName == "Round 2" ? (testPerformed != null ? testPerformed : string.Empty) : string.Empty;
                        ws.Cells["G" + row].Value = roundName == "Round 3" ? (testPerformed != null ? testPerformed : string.Empty) : string.Empty;
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":D" + row].Merge = true;
                        ws.Cells["E" + row + ":F" + row].Merge = true;
                        ws.Cells["G" + row + ":H" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);
                        xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetHorizontalAlignCenter(ws, "C" + row + ":H" + row);

                        row++;
                        //Row 25
                        string dtTestingPerformed = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("when was the testing completed")).Select(x => x.StrAnswer).FirstOrDefault();
                        string parsedtTestingPerformed = string.Empty;
                        if (dtTestingPerformed != null && dtTestingPerformed != string.Empty)
                        {
                            //parsedtTestingPerformed = DateTime.Parse(dtTestingPerformed).ToString("MM/dd/yyyy");
                            DateTime tempdtTestingPerformed;
                            //try to parse string to date time
                            if (DateTime.TryParse(dtTestingPerformed, out tempdtTestingPerformed))
                            {
                                parsedtTestingPerformed = tempdtTestingPerformed.ToString("MM/dd/yyyy");
                            }
                            else
                            {
                                //if it failed to parse date time, then we will just remove time string
                                int index = dtTestingPerformed.IndexOf(" ");
                                if (index > 0)
                                {
                                    parsedtTestingPerformed = dtTestingPerformed.Substring(0, index);
                                }
                            }
                        }

                        ws.Cells["A" + row].Value = "Date Testing Performed:";
                        ws.Cells["C" + row + ":G" + row].Style.Numberformat.Format = "@";
                        ws.Cells["C" + row].Value = roundName == "Round 1" ? (parsedtTestingPerformed != null ? parsedtTestingPerformed : string.Empty) : string.Empty;
                        ws.Cells["E" + row].Value = roundName == "Round 2" ? (parsedtTestingPerformed != null ? parsedtTestingPerformed : string.Empty) : string.Empty;
                        ws.Cells["G" + row].Value = roundName == "Round 3" ? (parsedtTestingPerformed != null ? parsedtTestingPerformed : string.Empty) : string.Empty;
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":D" + row].Merge = true;
                        ws.Cells["E" + row + ":F" + row].Merge = true;
                        ws.Cells["G" + row + ":H" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);
                        xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetHorizontalAlignCenter(ws, "C" + row + ":H" + row);

                        row++;
                        //Row 26
                        string testOfDesign = string.Empty;
                        //#18 in questionnaire
                        //if (clientName == "ERI" || clientName == "McGrath")
                        //    testOfDesign = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("what is the testing result")).Select(x => x.StrAnswer).FirstOrDefault();
                        //else
                        testOfDesign = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("(tod)")).Select(x => x.StrAnswer).FirstOrDefault();
                        ws.Cells["A" + row].Value = "Test of Design (TOD):";
                        ws.Cells["C" + row].Value = roundName == "Round 1" ? (testOfDesign != null ? testOfDesign : string.Empty) : string.Empty;
                        ws.Cells["E" + row].Value = roundName == "Round 2" ? (testOfDesign != null ? testOfDesign : string.Empty) : string.Empty;
                        ws.Cells["G" + row].Value = roundName == "Round 3" ? (testOfDesign != null ? testOfDesign : string.Empty) : string.Empty;
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":D" + row].Merge = true;
                        ws.Cells["E" + row + ":F" + row].Merge = true;
                        ws.Cells["G" + row + ":H" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);
                        xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetHorizontalAlignCenter(ws, "C" + row + ":H" + row);

                        row++;
                        //Row 27
                        //#19 in questionnaire
                        string testOperatingEffect = string.Empty;
                        //if (clientName == "ERI" || clientName == "McGrath")
                        //    testOperatingEffect = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("what is the testing result")).Select(x => x.StrAnswer).FirstOrDefault();
                        //else

                        //
                        
                            testOperatingEffect = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("(toe)")).Select(x => x.StrAnswer).FirstOrDefault();
                            
                       
                        
                        ws.Cells["C" + row].Value = roundName == "Round 1" ? (testOperatingEffect != null ? testOperatingEffect : string.Empty) : string.Empty;
                        ws.Cells["E" + row].Value = roundName == "Round 2" ? (testOperatingEffect != null ? testOperatingEffect : string.Empty) : string.Empty;
                        ws.Cells["G" + row].Value = roundName == "Round 3" ? (testOperatingEffect != null ? testOperatingEffect : string.Empty) : string.Empty;
                        ws.Cells["A" + row].Value = "Test of Operating Effectiveness (TOE):";
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":D" + row].Merge = true;
                        ws.Cells["E" + row + ":F" + row].Merge = true;
                        ws.Cells["G" + row + ":H" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);
                        xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetHorizontalAlignCenter(ws, "C" + row + ":H" + row);

                        //int row = 30;
                        //Row 30;
                        row += 3;
                        var listRound = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.Equals("Round Reference")).FirstOrDefault();


                        #region Table 1

                        if (listRound != null && listRound.ListRoundItem != null)
                        {

                            if (clientName == "ERI" && controlId == "HRP 2.1")
                            {
                                string title = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("(rt title)")).Select(x => x.StrQuestion.Replace("(RT Title)", "")).FirstOrDefault();
                                ws.Cells["A" + row].Value = $"{title}";
                                ws.Cells["A" + row + ":C" + row].Merge = true;
                                xlsService.ExcelSetBorder(ws, row, 1, row, 3);
                                xlsService.ExcelSetBackgroundColorGray(ws, row, 1, row, 3);
                                xlsService.ExcelSetArialSize12(ws, row, 1, row, 3);
                                xlsService.ExcelSetFontBold(ws, row, 1, row, 3);

                                string answer = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("(rt title)")).Select(x => x.StrAnswer).FirstOrDefault();
                                ws.Cells["D" + row].Value = $"{answer}";
                                ws.Cells["D" + row + ":F" + row].Merge = true;
                                xlsService.ExcelWrapText(ws, row, 4, row, 6);
                                xlsService.ExcelSetBorder(ws, row, 4, row, 6);
                                xlsService.ExcelSetBackgroundColorLightGray(ws, row, 4, row, 6);
                                xlsService.ExcelSetArialSize10(ws, row, 4, row, 6);
                                xlsService.ExcelSetFontBold(ws, row, 4, row, 6);
                                xlsService.ExcelSetVerticalAlignCenter(ws, row, 4, row, 6);
                                row++;
                            }

                            #region Round Sample

                            int columnRoundHeader = 1;

                            #region Round Header
                            var roundHeader = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.Contains("(RT)"));
                            if (roundHeader != null)
                            {
                                ws.Cells[row, columnRoundHeader].Value = "Sample #";
                                columnRoundHeader++;
                                int countHeaderNotes = 0;
                                foreach (var item in roundHeader)
                                {
                                    //filter policy changes for ELC 1.1
                                    if (!item.StrQuestion.Contains("Policy Changes? (R1)") &&
                                        !item.StrQuestion.Contains("Policy Changes? (R2)") &&
                                        !item.StrQuestion.Contains("Policy Changes? (R3)"))
                                    {
                                        string strHeader = item.StrQuestion.Replace("(RT)", "");
                                        string tempNote = string.Empty;

                                        if (questionnaireInput.ListHeaderNote != null && questionnaireInput.ListHeaderNote.Count > 0)
                                        {
                                            var checkHeaderNote = questionnaireInput.ListHeaderNote.Where(x => x.Position.Equals(countHeaderNotes)).FirstOrDefault();
                                            if (checkHeaderNote != null && checkHeaderNote.HeaderNoteText != string.Empty)
                                            {
                                                //strHeader = strHeader + (checkHeaderNote.HeaderNoteText != string.Empty ? $" [{checkHeaderNote.HeaderNoteText}] " : string.Empty);
                                                tempNote = (checkHeaderNote.HeaderNoteText != string.Empty ? checkHeaderNote.HeaderNoteText : string.Empty);
                                            }
                                        }

                                        //ws.Cells[row, columnRoundHeader].Value = strHeader;

                                        xlsService.TestingAttributeFormat(ws, row, columnRoundHeader, strHeader, tempNote);
                                        columnRoundHeader++;
                                        countHeaderNotes++;
                                    }
                                }

                                xlsService.ExcelWrapText(ws, row, 1, row, columnRoundHeader - 1);
                                xlsService.ExcelSetBorder(ws, row, 1, row, columnRoundHeader - 1);
                                xlsService.ExcelSetBackgroundColorGray(ws, row, 1, row, columnRoundHeader - 1);
                                xlsService.ExcelSetArialSize10(ws, row, 1, row, columnRoundHeader - 1);
                                xlsService.ExcelSetFontBold(ws, row, 1, row, columnRoundHeader - 1);
                                xlsService.ExcelSetVerticalAlignCenter(ws, row, 1, row, columnRoundHeader - 1);
                                xlsService.ExcelSetHorizontalAlignCenter(ws, row, 1, row, columnRoundHeader - 1);
                                row++;
                            }

                            #endregion


                            #region Round 1
                            //Round 1
                            ws.Cells[row, 1].Value = "Round 1";
                            ws.Cells[row, 1, row, columnRoundHeader - 1].Merge = true;
                            xlsService.ExcelWrapText(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetBorder(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetBackgroundColorLightGray(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetArialSize10(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetFontBold(ws, row, 1, row, columnRoundHeader - 1);
                            row++;


                            if (listRound != null && listRound.ListRoundItem.Count() > 0)
                            {
                                foreach (var item in listRound.ListRoundItem)
                                {
                                    int countRoundData1 = 1;
                                    if (item.RoundName == "Round 1")
                                    {

                                        if (questionnaireInput.ListPolicyNote != null && questionnaireInput.ListPolicyNote.Count > 0)
                                        {
                                            string policyQuestion1 = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("policy changes? (r1)")).Select(x => x.StrAnswer).FirstOrDefault();
                                            string policyNote1 = questionnaireInput.ListPolicyNote.Where(x => x.Position.Equals(0)).Select(x => x.NoteText).FirstOrDefault();
                                            string policy1 = (policyQuestion1 != null ? policyQuestion1 : string.Empty) + (policyNote1 != null && policyNote1 != string.Empty ? $" [{policyNote1}]" : string.Empty);
                                            ws.Cells[row, 1].Value = "Policy Changes?";
                                            ws.Cells[row, 2].Value = policy1 != null ? policy1 : string.Empty;
                                            ws.Cells[row, 2, row, columnRoundHeader - 1].Merge = true;
                                            xlsService.ExcelWrapText(ws, row, 1, row, columnRoundHeader - 1);
                                            xlsService.ExcelSetBorder(ws, row, 1, row, columnRoundHeader - 1);
                                            xlsService.ExcelSetHorizontalAlignCenter(ws, row, 2, row, columnRoundHeader - 1);
                                            xlsService.ExcelSetVerticalAlignCenter(ws, row, 2, row, columnRoundHeader - 1);
                                            xlsService.ExcelSetArialSize10(ws, row, 1, row, columnRoundHeader - 1);
                                            xlsService.ExcelSetBackgroundColorLightGray(ws, row, 1, row, columnRoundHeader - 1);
                                            row++;
                                        }

                                        #region Populate Data for Round 1
                                        //if (item.A2Q2Samples != null && item.A2Q2Samples != string.Empty)
                                        ws.Cells[row, countRoundData1].Value = $"{(item.A2Q2Samples != null && item.A2Q2Samples != string.Empty ? item.A2Q2Samples.ToString() : string.Empty)}";
                                        countRoundData1++;

                                        //if (item.Answer1 != null && item.Answer1 != string.Empty)
                                        //ws.Cells[row, countRoundData1].Value = $"{(item.Answer1 != null && item.Answer1 != string.Empty ? item.Answer1.ToString() : string.Empty)} " + (item.Note1 != null ? $"[{item.Note1}]" : string.Empty);
                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer1, item.Note1);
                                        countRoundData1++;

                                        //if (item.Answer2 != null && item.Answer2 != string.Empty)
                                        //ws.Cells[row, countRoundData1].Value = $"{(item.Answer2 != null ? item.Answer2.ToString() : string.Empty)} " + (item.Note2 != null ? $"[{item.Note2}]" : string.Empty);
                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer2, item.Note2);
                                        countRoundData1++;

                                        //if (item.Answer3 != null && item.Answer3 != string.Empty)
                                        //ws.Cells[row, countRoundData1].Value = $"{(item.Answer3 != null && item.Answer3 != string.Empty ? item.Answer3.ToString() : string.Empty)} " + (item.Note3 != null ? $"[{item.Note3}]" : string.Empty);
                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer3, item.Note3);
                                        countRoundData1++;

                                        //if (item.Answer4 != null && item.Answer4 != string.Empty)
                                        //ws.Cells[row, countRoundData1].Value = $"{(item.Answer4 != null && item.Answer4 != string.Empty ? item.Answer4.ToString() : string.Empty)} " + (item.Note4 != null ? $"[{item.Note4}]" : string.Empty);
                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer4, item.Note4);
                                        countRoundData1++;

                                        //if (item.Answer5 != null && item.Answer5 != string.Empty)
                                        //ws.Cells[row, countRoundData1].Value = $"{(item.Answer5 != null && item.Answer5 != string.Empty ? item.Answer5.ToString() : string.Empty)} " + (item.Note5 != null ? $"[{item.Note5}]" : string.Empty);
                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer5, item.Note5);
                                        countRoundData1++;

                                        //if (item.Answer6 != null && item.Answer6 != string.Empty)
                                        //ws.Cells[row, countRoundData1].Value = $"{(item.Answer6 != null && item.Answer6 != string.Empty ? item.Answer6.ToString() : string.Empty)} " + (item.Note6 != null ? $"[{item.Note6}]" : string.Empty);
                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer6, item.Note6);
                                        countRoundData1++;

                                        //if (item.Answer7 != null && item.Answer7 != string.Empty)
                                        //ws.Cells[row, countRoundData1].Value = $"{(item.Answer7 != null && item.Answer7 != string.Empty ? item.Answer7.ToString() : string.Empty)} " + (item.Note7 != null ? $"[{item.Note7}]" : string.Empty);
                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer7, item.Note7);
                                        countRoundData1++;

                                        //if (item.Answer8 != null && item.Answer8 != string.Empty)
                                        //ws.Cells[row, countRoundData1].Value = $"{(item.Answer8 != null && item.Answer8 != string.Empty ? item.Answer8.ToString() : string.Empty)} " + (item.Note8 != null ? $"[{item.Note8}]" : string.Empty);
                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer8, item.Note8);
                                        countRoundData1++;

                                        //if (item.Answer9 != null && item.Answer9 != string.Empty)
                                        //ws.Cells[row, countRoundData1].Value = $"{(item.Answer9 != null && item.Answer9 != string.Empty ? item.Answer9.ToString() : string.Empty)} " + (item.Note9 != null ? $"[{item.Note9}]" : string.Empty);
                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer9, item.Note9);
                                        countRoundData1++;

                                        //if (item.Answer10 != null && item.Answer10 != string.Empty)
                                        //ws.Cells[row, countRoundData1].Value = $"{(item.Answer10 != null && item.Answer10 != string.Empty ? item.Answer10.ToString() : string.Empty)} " + (item.Note10 != null ? $"[{item.Note10}]" : string.Empty);
                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer10, item.Note10);
                                        countRoundData1++;

                                        //if (item.Answer11 != null && item.Answer11 != string.Empty)
                                        //ws.Cells[row, countRoundData1].Value = $"{(item.Answer11 != null && item.Answer11 != string.Empty ? item.Answer11.ToString() : string.Empty)} " + (item.Note11 != null ? $"[{item.Note11}]" : string.Empty);
                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer11, item.Note11);
                                        countRoundData1++;

                                        //if (item.Answer12 != null && item.Answer12 != string.Empty)
                                        //ws.Cells[row, countRoundData1].Value = $"{(item.Answer12 != null && item.Answer12 != string.Empty ? item.Answer12.ToString() : string.Empty)} " + (item.Note12 != null ? $"[{item.Note12}]" : string.Empty);
                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer12, item.Note12);
                                        countRoundData1++;

                                        //if (item.Answer13 != null && item.Answer13 != string.Empty)
                                        //ws.Cells[row, countRoundData1].Value = $"{(item.Answer13 != null && item.Answer13 != string.Empty ? item.Answer13.ToString() : string.Empty)} " + (item.Note13 != null ? $"[{item.Note13}]" : string.Empty);
                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer13, item.Note13);
                                        countRoundData1++;

                                        //if (item.Answer14 != null && item.Answer14 != string.Empty)
                                        //ws.Cells[row, countRoundData1].Value = $"{(item.Answer14 != null && item.Answer14 != string.Empty ? item.Answer14.ToString() : string.Empty)} " + (item.Note14 != null ? $"[{item.Note14}]" : string.Empty);
                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer14, item.Note14);
                                        countRoundData1++;

                                        //if (item.Answer15 != null && item.Answer15 != string.Empty)
                                        //ws.Cells[row, countRoundData1].Value = $"{(item.Answer15 != null && item.Answer15 != string.Empty ? item.Answer15.ToString() : string.Empty)} " + (item.Note15 != null ? $"[{item.Note15}]" : string.Empty);
                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer15, item.Note15);
                                        countRoundData1++;
                                        #endregion

                                        xlsService.ExcelWrapText(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetBorder(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetArialSize10(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetVerticalAlignCenter(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetHorizontalAlignCenter(ws, row, 1, row, columnRoundHeader - 1);

                                        ws.Row(row).Height = 30;
                                        row++;
                                    }
                                }
                            }

                            #endregion

                            #region Round 2
                            //Round 2
                            ws.Cells[row, 1].Value = "Round 2";
                            ws.Cells[row, 1, row, columnRoundHeader - 1].Merge = true;
                            xlsService.ExcelWrapText(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetBorder(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetBackgroundColorLightGray(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetArialSize10(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetFontBold(ws, row, 1, row, columnRoundHeader - 1);
                            row++;

                            if (listRound != null && listRound.ListRoundItem.Count() > 0)
                            {
                                foreach (var item in listRound.ListRoundItem)
                                {
                                    int countRoundData1 = 1;
                                    if (item.RoundName == "Round 2")
                                    {

                                        if (questionnaireInput.ListPolicyNote != null && questionnaireInput.ListPolicyNote.Count > 0)
                                        {
                                            string policyQuestion1 = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("policy changes? (r2)")).Select(x => x.StrAnswer).FirstOrDefault();
                                            string policyNote1 = questionnaireInput.ListPolicyNote.Where(x => x.Position.Equals(1)).Select(x => x.NoteText).FirstOrDefault();
                                            string policy1 = (policyQuestion1 != null ? policyQuestion1 : string.Empty) + (policyNote1 != null && policyNote1 != string.Empty ? $" [{policyNote1}]" : string.Empty);
                                            ws.Cells[row, 1].Value = "Policy Changes?";
                                            ws.Cells[row, 2].Value = policy1 != null ? policy1 : string.Empty;
                                            ws.Cells[row, 2, row, columnRoundHeader - 1].Merge = true;
                                            xlsService.ExcelWrapText(ws, row, 1, row, columnRoundHeader - 1);
                                            xlsService.ExcelSetBorder(ws, row, 1, row, columnRoundHeader - 1);
                                            xlsService.ExcelSetHorizontalAlignCenter(ws, row, 2, row, columnRoundHeader - 1);
                                            xlsService.ExcelSetVerticalAlignCenter(ws, row, 2, row, columnRoundHeader - 1);
                                            xlsService.ExcelSetArialSize10(ws, row, 1, row, columnRoundHeader - 1);
                                            xlsService.ExcelSetBackgroundColorLightGray(ws, row, 1, row, columnRoundHeader - 1);
                                            row++;
                                        }


                                        #region Populate Data for Round 2
                                        //if (item.A2Q2Samples != null && item.A2Q2Samples != string.Empty)
                                        ws.Cells[row, countRoundData1].Value = $"{(item.A2Q2Samples != null && item.A2Q2Samples != string.Empty ? item.A2Q2Samples.ToString() : string.Empty)}";
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer1, item.Note1);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer2, item.Note2);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer3, item.Note3);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer4, item.Note4);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer5, item.Note5);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer6, item.Note6);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer7, item.Note7);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer8, item.Note8);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer9, item.Note9);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer10, item.Note10);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer11, item.Note11);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer12, item.Note12);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer13, item.Note13);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer14, item.Note14);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer15, item.Note15);
                                        countRoundData1++;


                                        #endregion

                                        xlsService.ExcelWrapText(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetBorder(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetArialSize10(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetVerticalAlignCenter(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetHorizontalAlignCenter(ws, row, 1, row, columnRoundHeader - 1);

                                        ws.Row(row).Height = 30;
                                        row++;
                                    }
                                }
                            }

                            #endregion

                            #region Round 3
                            //Round 3
                            ws.Cells[row, 1].Value = "Round 3";
                            ws.Cells[row, 1, row, columnRoundHeader - 1].Merge = true;
                            xlsService.ExcelWrapText(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetBorder(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetBackgroundColorLightGray(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetArialSize10(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetFontBold(ws, row, 1, row, columnRoundHeader - 1);
                            row++;

                            if (listRound != null && listRound.ListRoundItem.Count() > 0)
                            {
                                foreach (var item in listRound.ListRoundItem)
                                {
                                    int countRoundData1 = 1;
                                    if (item.RoundName == "Round 3")
                                    {

                                        if (questionnaireInput.ListPolicyNote != null && questionnaireInput.ListPolicyNote.Count > 0)
                                        {
                                            string policyQuestion1 = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("policy changes? (r3)")).Select(x => x.StrAnswer).FirstOrDefault();
                                            string policyNote1 = questionnaireInput.ListPolicyNote.Where(x => x.Position.Equals(2)).Select(x => x.NoteText).FirstOrDefault();
                                            string policy1 = (policyQuestion1 != null ? policyQuestion1 : string.Empty) + (policyNote1 != null && policyNote1 != string.Empty ? $" [{policyNote1}]" : string.Empty);
                                            ws.Cells[row, 1].Value = "Policy Changes?";
                                            ws.Cells[row, 2].Value = policy1 != null ? policy1 : string.Empty;
                                            ws.Cells[row, 2, row, columnRoundHeader - 1].Merge = true;
                                            xlsService.ExcelWrapText(ws, row, 1, row, columnRoundHeader - 1);
                                            xlsService.ExcelSetBorder(ws, row, 1, row, columnRoundHeader - 1);
                                            xlsService.ExcelSetHorizontalAlignCenter(ws, row, 2, row, columnRoundHeader - 1);
                                            xlsService.ExcelSetVerticalAlignCenter(ws, row, 2, row, columnRoundHeader - 1);
                                            xlsService.ExcelSetArialSize10(ws, row, 1, row, columnRoundHeader - 1);
                                            xlsService.ExcelSetBackgroundColorLightGray(ws, row, 1, row, columnRoundHeader - 1);
                                            row++;
                                        }


                                        #region Populate Data for Round 3
                                        //if (item.A2Q2Samples != null && item.A2Q2Samples != string.Empty)
                                        ws.Cells[row, countRoundData1].Value = $"{(item.A2Q2Samples != null && item.A2Q2Samples != string.Empty ? item.A2Q2Samples.ToString() : string.Empty)}";
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer1, item.Note1);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer2, item.Note2);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer3, item.Note3);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer4, item.Note4);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer5, item.Note5);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer6, item.Note6);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer7, item.Note7);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer8, item.Note8);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer9, item.Note9);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer10, item.Note10);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer11, item.Note11);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer12, item.Note12);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer13, item.Note13);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer14, item.Note14);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer15, item.Note15);
                                        countRoundData1++;

                                        #endregion

                                        xlsService.ExcelWrapText(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetBorder(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetArialSize10(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetVerticalAlignCenter(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetHorizontalAlignCenter(ws, row, 1, row, columnRoundHeader - 1);

                                        ws.Row(row).Height = 30;
                                        row++;
                                    }
                                }
                            }

                            #endregion

                            #endregion

                        }

                        #endregion

                        #region Table 2
                        row += 2;
                        if (listRound != null && listRound.ListRoundItem2 != null)
                        {

                            if (clientName == "ERI" && controlId == "HRP 2.1")
                            {
                                string title = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("(rt2 title)")).Select(x => x.StrQuestion.Replace("(RT2 Title)", "")).FirstOrDefault();
                                ws.Cells["A" + row].Value = $"{title}";
                                ws.Cells["A" + row + ":C" + row].Merge = true;
                                xlsService.ExcelSetBorder(ws, row, 1, row, 3);
                                xlsService.ExcelSetBackgroundColorGray(ws, row, 1, row, 3);
                                xlsService.ExcelSetArialSize12(ws, row, 1, row, 3);
                                xlsService.ExcelSetFontBold(ws, row, 1, row, 3);

                                string answer = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("(rt2 title)")).Select(x => x.StrAnswer).FirstOrDefault();
                                ws.Cells["D" + row].Value = $"{answer}";
                                ws.Cells["D" + row + ":F" + row].Merge = true;
                                xlsService.ExcelWrapText(ws, row, 4, row, 6);
                                xlsService.ExcelSetBorder(ws, row, 4, row, 6);
                                xlsService.ExcelSetBackgroundColorLightGray(ws, row, 4, row, 6);
                                xlsService.ExcelSetArialSize10(ws, row, 4, row, 6);
                                xlsService.ExcelSetFontBold(ws, row, 4, row, 6);
                                xlsService.ExcelSetVerticalAlignCenter(ws, row, 4, row, 6);
                                row++;
                            }

                            #region Round Sample

                            int columnRoundHeader = 1;

                            #region Round Header
                            var roundHeader = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.Contains("(RT2)"));
                            if (roundHeader != null)
                            {
                                ws.Cells[row, columnRoundHeader].Value = "Sample #";
                                columnRoundHeader++;
                                int countHeaderNotes = 0;
                                foreach (var item in roundHeader)
                                {

                                    string strHeader = item.StrQuestion.Replace("(RT2)", "");
                                    string tempNote = string.Empty;

                                    if (questionnaireInput.ListHeaderNote2 != null && questionnaireInput.ListHeaderNote2.Count > 0)
                                    {
                                        var checkHeaderNote = questionnaireInput.ListHeaderNote2.Where(x => x.Position.Equals(countHeaderNotes)).FirstOrDefault();
                                        if (checkHeaderNote != null && checkHeaderNote.HeaderNoteText != string.Empty)
                                        {
                                            //strHeader = strHeader + (checkHeaderNote.HeaderNoteText != string.Empty ? $" [{checkHeaderNote.HeaderNoteText}] " : string.Empty);
                                            tempNote = (checkHeaderNote.HeaderNoteText != string.Empty ? checkHeaderNote.HeaderNoteText : string.Empty);
                                        }
                                    }

                                    //ws.Cells[row, columnRoundHeader].Value = strHeader;
                                    xlsService.TestingAttributeFormat(ws, row, columnRoundHeader, strHeader, tempNote);
                                    columnRoundHeader++;
                                    countHeaderNotes++;

                                }

                                xlsService.ExcelWrapText(ws, row, 1, row, columnRoundHeader - 1);
                                xlsService.ExcelSetBorder(ws, row, 1, row, columnRoundHeader - 1);
                                xlsService.ExcelSetBackgroundColorGray(ws, row, 1, row, columnRoundHeader - 1);
                                xlsService.ExcelSetArialSize10(ws, row, 1, row, columnRoundHeader - 1);
                                xlsService.ExcelSetFontBold(ws, row, 1, row, columnRoundHeader - 1);
                                xlsService.ExcelSetVerticalAlignCenter(ws, row, 1, row, columnRoundHeader - 1);
                                xlsService.ExcelSetHorizontalAlignCenter(ws, row, 1, row, columnRoundHeader - 1);
                                row++;
                            }

                            #endregion


                            #region Round 1
                            //Round 1
                            ws.Cells[row, 1].Value = "Round 1";
                            ws.Cells[row, 1, row, columnRoundHeader - 1].Merge = true;
                            xlsService.ExcelWrapText(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetBorder(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetBackgroundColorLightGray(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetArialSize10(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetFontBold(ws, row, 1, row, columnRoundHeader - 1);
                            row++;


                            if (listRound != null && listRound.ListRoundItem2.Count() > 0)
                            {
                                foreach (var item in listRound.ListRoundItem2)
                                {
                                    int countRoundData1 = 1;
                                    if (item.RoundName == "Round 1")
                                    {

                                        #region Populate Data for Round 1

                                        //if (item.A2Q2Samples != null && item.A2Q2Samples != string.Empty)
                                        ws.Cells[row, countRoundData1].Value = $"{(item.A2Q2Samples != null && item.A2Q2Samples != string.Empty ? item.A2Q2Samples.ToString() : string.Empty)}";
                                        countRoundData1++;


                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer1, item.Note1);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer2, item.Note2);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer3, item.Note3);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer4, item.Note4);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer5, item.Note5);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer6, item.Note6);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer7, item.Note7);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer8, item.Note8);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer9, item.Note9);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer10, item.Note10);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer11, item.Note11);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer12, item.Note12);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer13, item.Note13);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer14, item.Note14);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer15, item.Note15);
                                        countRoundData1++;
                                        #endregion

                                        xlsService.ExcelWrapText(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetBorder(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetArialSize10(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetVerticalAlignCenter(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetHorizontalAlignCenter(ws, row, 1, row, columnRoundHeader - 1);

                                        ws.Row(row).Height = 30;
                                        row++;
                                    }
                                }
                            }

                            #endregion

                            #region Round 2
                            //Round 2
                            ws.Cells[row, 1].Value = "Round 2";
                            ws.Cells[row, 1, row, columnRoundHeader - 1].Merge = true;
                            xlsService.ExcelWrapText(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetBorder(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetBackgroundColorLightGray(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetArialSize10(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetFontBold(ws, row, 1, row, columnRoundHeader - 1);
                            row++;

                            if (listRound != null && listRound.ListRoundItem2.Count() > 0)
                            {
                                foreach (var item in listRound.ListRoundItem2)
                                {
                                    int countRoundData1 = 1;
                                    if (item.RoundName == "Round 2")
                                    {


                                        #region Populate Data for Round 2
                                        //if (item.A2Q2Samples != null && item.A2Q2Samples != string.Empty)
                                        ws.Cells[row, countRoundData1].Value = $"{(item.A2Q2Samples != null && item.A2Q2Samples != string.Empty ? item.A2Q2Samples.ToString() : string.Empty)}";
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer1, item.Note1);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer2, item.Note2);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer3, item.Note3);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer4, item.Note4);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer5, item.Note5);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer6, item.Note6);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer7, item.Note7);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer8, item.Note8);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer9, item.Note9);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer10, item.Note10);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer11, item.Note11);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer12, item.Note12);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer13, item.Note13);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer14, item.Note14);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer15, item.Note15);
                                        countRoundData1++;
                                        #endregion

                                        xlsService.ExcelWrapText(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetBorder(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetArialSize10(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetVerticalAlignCenter(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetHorizontalAlignCenter(ws, row, 1, row, columnRoundHeader - 1);

                                        ws.Row(row).Height = 30;
                                        row++;
                                    }
                                }
                            }

                            #endregion

                            #region Round 3
                            //Round 3
                            ws.Cells[row, 1].Value = "Round 3";
                            ws.Cells[row, 1, row, columnRoundHeader - 1].Merge = true;
                            xlsService.ExcelWrapText(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetBorder(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetBackgroundColorLightGray(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetArialSize10(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetFontBold(ws, row, 1, row, columnRoundHeader - 1);
                            row++;

                            if (listRound != null && listRound.ListRoundItem2.Count() > 0)
                            {
                                foreach (var item in listRound.ListRoundItem2)
                                {
                                    int countRoundData1 = 1;
                                    if (item.RoundName == "Round 3")
                                    {

                                        #region Populate Data for Round 3
                                        //if (item.A2Q2Samples != null && item.A2Q2Samples != string.Empty)
                                        ws.Cells[row, countRoundData1].Value = $"{(item.A2Q2Samples != null && item.A2Q2Samples != string.Empty ? item.A2Q2Samples.ToString() : string.Empty)}";
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer1, item.Note1);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer2, item.Note2);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer3, item.Note3);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer4, item.Note4);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer5, item.Note5);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer6, item.Note6);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer7, item.Note7);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer8, item.Note8);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer9, item.Note9);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer10, item.Note10);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer11, item.Note11);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer12, item.Note12);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer13, item.Note13);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer14, item.Note14);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer15, item.Note15);
                                        countRoundData1++;
                                        #endregion

                                        xlsService.ExcelWrapText(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetBorder(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetArialSize10(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetVerticalAlignCenter(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetHorizontalAlignCenter(ws, row, 1, row, columnRoundHeader - 1);

                                        ws.Row(row).Height = 30;
                                        row++;
                                    }
                                }
                            }

                            #endregion

                            #endregion

                        }

                        #endregion

                        #region Table 3
                        row += 2;

                        if (listRound != null && listRound.ListRoundItem3 != null)
                        {

                            if (clientName == "ERI" && controlId == "HRP 2.1")
                            {
                                string title = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("(rt3 title)")).Select(x => x.StrQuestion.Replace("(RT3 Title)", "")).FirstOrDefault();
                                ws.Cells["A" + row].Value = $"{title}";
                                ws.Cells["A" + row + ":C" + row].Merge = true;
                                xlsService.ExcelSetBorder(ws, row, 1, row, 3);
                                xlsService.ExcelSetBackgroundColorGray(ws, row, 1, row, 3);
                                xlsService.ExcelSetArialSize12(ws, row, 1, row, 3);
                                xlsService.ExcelSetFontBold(ws, row, 1, row, 3);

                                string answer = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("(rt3 title)")).Select(x => x.StrAnswer).FirstOrDefault();
                                ws.Cells["D" + row].Value = $"{answer}";
                                ws.Cells["D" + row + ":F" + row].Merge = true;
                                xlsService.ExcelWrapText(ws, row, 4, row, 6);
                                xlsService.ExcelSetBorder(ws, row, 4, row, 6);
                                xlsService.ExcelSetBackgroundColorLightGray(ws, row, 4, row, 6);
                                xlsService.ExcelSetArialSize10(ws, row, 4, row, 6);
                                xlsService.ExcelSetFontBold(ws, row, 4, row, 6);
                                xlsService.ExcelSetVerticalAlignCenter(ws, row, 4, row, 6);
                                row++;
                            }

                            #region Round Sample

                            int columnRoundHeader = 1;

                            #region Round Header
                            var roundHeader = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.Contains("(RT3)"));
                            if (roundHeader != null)
                            {
                                ws.Cells[row, columnRoundHeader].Value = "Sample #";
                                columnRoundHeader++;
                                int countHeaderNotes = 0;
                                foreach (var item in roundHeader)
                                {

                                    string strHeader = item.StrQuestion.Replace("(RT3)", "");
                                    string tempNote = string.Empty;

                                    if (questionnaireInput.ListHeaderNote3 != null && questionnaireInput.ListHeaderNote3.Count > 0)
                                    {
                                        var checkHeaderNote = questionnaireInput.ListHeaderNote3.Where(x => x.Position.Equals(countHeaderNotes)).FirstOrDefault();
                                        if (checkHeaderNote != null && checkHeaderNote.HeaderNoteText != string.Empty)
                                        {
                                            //strHeader = strHeader + (checkHeaderNote.HeaderNoteText != string.Empty ? $" [{checkHeaderNote.HeaderNoteText}] " : string.Empty);
                                            tempNote = (checkHeaderNote.HeaderNoteText != string.Empty ? checkHeaderNote.HeaderNoteText : string.Empty);
                                        }
                                    }

                                    //ws.Cells[row, columnRoundHeader].Value = strHeader;
                                    xlsService.TestingAttributeFormat(ws, row, columnRoundHeader, strHeader, tempNote);
                                    columnRoundHeader++;
                                    countHeaderNotes++;

                                }

                                xlsService.ExcelWrapText(ws, row, 1, row, columnRoundHeader - 1);
                                xlsService.ExcelSetBorder(ws, row, 1, row, columnRoundHeader - 1);
                                xlsService.ExcelSetBackgroundColorGray(ws, row, 1, row, columnRoundHeader - 1);
                                xlsService.ExcelSetArialSize10(ws, row, 1, row, columnRoundHeader - 1);
                                xlsService.ExcelSetFontBold(ws, row, 1, row, columnRoundHeader - 1);
                                xlsService.ExcelSetVerticalAlignCenter(ws, row, 1, row, columnRoundHeader - 1);
                                xlsService.ExcelSetHorizontalAlignCenter(ws, row, 1, row, columnRoundHeader - 1);
                                row++;
                            }

                            #endregion


                            #region Round 1
                            //Round 1
                            ws.Cells[row, 1].Value = "Round 1";
                            ws.Cells[row, 1, row, columnRoundHeader - 1].Merge = true;
                            xlsService.ExcelWrapText(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetBorder(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetBackgroundColorLightGray(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetArialSize10(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetFontBold(ws, row, 1, row, columnRoundHeader - 1);
                            row++;


                            if (listRound != null && listRound.ListRoundItem3.Count() > 0)
                            {
                                foreach (var item in listRound.ListRoundItem3)
                                {
                                    int countRoundData1 = 1;
                                    if (item.RoundName == "Round 1")
                                    {

                                        #region Populate Data for Round 1

                                        //if (item.A2Q2Samples != null && item.A2Q2Samples != string.Empty)
                                        ws.Cells[row, countRoundData1].Value = $"{(item.A2Q2Samples != null && item.A2Q2Samples != string.Empty ? item.A2Q2Samples.ToString() : string.Empty)}";
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer1, item.Note1);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer2, item.Note2);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer3, item.Note3);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer4, item.Note4);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer5, item.Note5);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer6, item.Note6);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer7, item.Note7);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer8, item.Note8);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer9, item.Note9);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer10, item.Note10);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer11, item.Note11);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer12, item.Note12);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer13, item.Note13);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer14, item.Note14);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer15, item.Note15);
                                        countRoundData1++;
                                        #endregion

                                        xlsService.ExcelWrapText(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetBorder(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetArialSize10(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetVerticalAlignCenter(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetHorizontalAlignCenter(ws, row, 1, row, columnRoundHeader - 1);

                                        ws.Row(row).Height = 30;
                                        row++;
                                    }
                                }
                            }

                            #endregion

                            #region Round 2
                            //Round 2
                            ws.Cells[row, 1].Value = "Round 2";
                            ws.Cells[row, 1, row, columnRoundHeader - 1].Merge = true;
                            xlsService.ExcelWrapText(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetBorder(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetBackgroundColorLightGray(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetArialSize10(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetFontBold(ws, row, 1, row, columnRoundHeader - 1);
                            row++;

                            if (listRound != null && listRound.ListRoundItem3.Count() > 0)
                            {
                                foreach (var item in listRound.ListRoundItem3)
                                {
                                    int countRoundData1 = 1;
                                    if (item.RoundName == "Round 2")
                                    {


                                        #region Populate Data for Round 2
                                        //if (item.A2Q2Samples != null && item.A2Q2Samples != string.Empty)
                                        ws.Cells[row, countRoundData1].Value = $"{(item.A2Q2Samples != null && item.A2Q2Samples != string.Empty ? item.A2Q2Samples.ToString() : string.Empty)}";
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer1, item.Note1);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer2, item.Note2);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer3, item.Note3);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer4, item.Note4);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer5, item.Note5);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer6, item.Note6);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer7, item.Note7);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer8, item.Note8);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer9, item.Note9);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer10, item.Note10);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer11, item.Note11);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer12, item.Note12);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer13, item.Note13);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer14, item.Note14);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer15, item.Note15);
                                        countRoundData1++;
                                        #endregion

                                        xlsService.ExcelWrapText(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetBorder(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetArialSize10(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetVerticalAlignCenter(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetHorizontalAlignCenter(ws, row, 1, row, columnRoundHeader - 1);

                                        ws.Row(row).Height = 30;
                                        row++;
                                    }
                                }
                            }

                            #endregion

                            #region Round 3
                            //Round 3
                            ws.Cells[row, 1].Value = "Round 3";
                            ws.Cells[row, 1, row, columnRoundHeader - 1].Merge = true;
                            xlsService.ExcelWrapText(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetBorder(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetBackgroundColorLightGray(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetArialSize10(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetFontBold(ws, row, 1, row, columnRoundHeader - 1);
                            row++;

                            if (listRound != null && listRound.ListRoundItem3.Count() > 0)
                            {
                                foreach (var item in listRound.ListRoundItem3)
                                {
                                    int countRoundData1 = 1;
                                    if (item.RoundName == "Round 3")
                                    {

                                        #region Populate Data for Round 3
                                        //if (item.A2Q2Samples != null && item.A2Q2Samples != string.Empty)
                                        ws.Cells[row, countRoundData1].Value = $"{(item.A2Q2Samples != null && item.A2Q2Samples != string.Empty ? item.A2Q2Samples.ToString() : string.Empty)}";
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer1, item.Note1);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer2, item.Note2);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer3, item.Note3);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer4, item.Note4);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer5, item.Note5);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer6, item.Note6);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer7, item.Note7);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer8, item.Note8);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer9, item.Note9);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer10, item.Note10);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer11, item.Note11);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer12, item.Note12);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer13, item.Note13);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer14, item.Note14);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer15, item.Note15);
                                        countRoundData1++;
                                        #endregion

                                        xlsService.ExcelWrapText(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetBorder(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetArialSize10(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetVerticalAlignCenter(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetHorizontalAlignCenter(ws, row, 1, row, columnRoundHeader - 1);

                                        ws.Row(row).Height = 30;
                                        row++;
                                    }
                                }
                            }

                            #endregion

                            #endregion

                        }

                        #endregion


                        #region Round Header
                        row++;
                        ws.Cells["C" + row].Value = "Round 1";
                        ws.Cells["C" + row + ":D" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "C" + row + ":D" + row);
                        xlsService.ExcelSetBorder(ws, "C" + row + ":D" + row);

                        ws.Cells["E" + row].Value = "Round 2";
                        ws.Cells["E" + row + ":F" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "E" + row + ":F" + row);
                        xlsService.ExcelSetBorder(ws, "E" + row + ":F" + row);

                        ws.Cells["G" + row].Value = "Round 3";
                        ws.Cells["G" + row + ":H" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "G" + row + ":H" + row);
                        xlsService.ExcelSetBorder(ws, "G" + row + ":H" + row);

                        xlsService.ExcelSetBackgroundColorGray(ws, "C" + row + ":H" + row);
                        xlsService.ExcelSetArialSize10(ws, "C" + row + ":H" + row);
                        xlsService.ExcelSetFontBold(ws, "C" + row + ":H" + row);
                        xlsService.ExcelSetVerticalAlignCenter(ws, "C" + row + ":H" + row);
                        xlsService.ExcelSetHorizontalAlignCenter(ws, "C" + row + ":H" + row);
                        row++;
                        #endregion



                        //24.Who performed the review? Enter first and last name.
                        string reviewer = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("who performed the review")).Select(x => x.StrAnswer).FirstOrDefault();
                        ws.Cells["A" + row].Value = "Reviewer:";
                        ws.Cells["C" + row].Value = roundName == "Round 1" ? (reviewer != null ? reviewer : string.Empty) : string.Empty;
                        ws.Cells["E" + row].Value = roundName == "Round 2" ? (reviewer != null ? reviewer : string.Empty) : string.Empty;
                        ws.Cells["G" + row].Value = roundName == "Round 3" ? (reviewer != null ? reviewer : string.Empty) : string.Empty;
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":D" + row].Merge = true;
                        ws.Cells["E" + row + ":F" + row].Merge = true;
                        ws.Cells["G" + row + ":H" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetHorizontalAlignCenter(ws, "A" + row + ":H" + row);
                        row++;

                        //25. When is the date the testing was reviewed? Enter date.
                        string reviewDate = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("when is the date the testing was reviewed")).Select(x => x.StrAnswer).FirstOrDefault();
                        string parsedtReviewDate = string.Empty;
                        if (reviewDate != null && reviewDate != string.Empty)
                        {
                            DateTime tempdtReviewDate;
                            //try to parse string to date time
                            if (DateTime.TryParse(reviewDate, out tempdtReviewDate))
                            {
                                parsedtReviewDate = tempdtReviewDate.ToString("MM/dd/yyyy");
                            }
                            else
                            {
                                //if it failed to parse date time, then we will just remove time string
                                int index = reviewDate.IndexOf(" ");
                                if (index > 0)
                                {
                                    parsedtReviewDate = reviewDate.Substring(0, index);
                                }
                            }
                        }

                        ws.Cells["A" + row].Value = "Review Date:";
                        ws.Cells["C" + row + ":G" + row].Style.Numberformat.Format = "@";
                        ws.Cells["C" + row].Value = roundName == "Round 1" ? (parsedtReviewDate != null ? parsedtReviewDate : string.Empty) : string.Empty;
                        ws.Cells["E" + row].Value = roundName == "Round 2" ? (parsedtReviewDate != null ? parsedtReviewDate : string.Empty) : string.Empty;
                        ws.Cells["G" + row].Value = roundName == "Round 3" ? (parsedtReviewDate != null ? parsedtReviewDate : string.Empty) : string.Empty;
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":D" + row].Merge = true;
                        ws.Cells["E" + row + ":F" + row].Merge = true;
                        ws.Cells["G" + row + ":H" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetHorizontalAlignCenter(ws, "A" + row + ":H" + row);
                        row++;

                        //19. What is the testing result?
                        string testFindings = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("what is the testing status")).Select(x => x.StrAnswer).FirstOrDefault();
                        ws.Cells["A" + row].Value = "Test Findings Assessment:";
                        ws.Cells["C" + row].Value = roundName == "Round 1" ? (testFindings != null ? testFindings : string.Empty) : string.Empty;
                        ws.Cells["E" + row].Value = roundName == "Round 2" ? (testFindings != null ? testFindings : string.Empty) : string.Empty;
                        ws.Cells["G" + row].Value = roundName == "Round 3" ? (testFindings != null ? testFindings : string.Empty) : string.Empty;
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":D" + row].Merge = true;
                        ws.Cells["E" + row + ":F" + row].Merge = true;
                        ws.Cells["G" + row + ":H" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetHorizontalAlignCenter(ws, "A" + row + ":H" + row);
                        row++;


                        #region Legends

                        var uniqueNotes = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.Equals("Unique Notes Reference")).FirstOrDefault();

                        row++;
                        ws.Cells["A" + row].Value = "Legend";
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row + ":B" + row);
                        row++;
                        if (uniqueNotes != null && uniqueNotes.ListNoteItem.Count > 0)
                        {
                            if (questionnaireInput.GeneralNote != null)
                            {
                                if (questionnaireInput.GeneralNote.Display)
                                {
                                    ws.Cells["A" + row].Value = questionnaireInput.GeneralNote.GeneralNoteText;
                                    ws.Cells["C" + row].Value = questionnaireInput.GeneralNote.Description != null ? questionnaireInput.GeneralNote.Description : string.Empty;
                                    ws.Cells["A" + row + ":B" + row].Merge = true;
                                    ws.Cells["C" + row + ":H" + row].Merge = true;
                                    xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                                    xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                                    xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                                    xlsService.ExcelSetFontBold(ws, "A" + row + ":B" + row);
                                    xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":H" + row);
                                    xlsService.ExcelSetHorizontalAlignCenter(ws, "A" + row + ":B" + row);
                                    xlsService.ExcelSetFontColorRed(ws, row, 1, row, 1);
                                    ws.Row(row).Height = 30;
                                    row++;
                                }
                            }


                            foreach (var item in uniqueNotes.ListNoteItem.OrderBy(x => x.Notes))
                            {
                                ws.Cells["A" + row].Value = item.Notes != null && item.Notes != string.Empty ? "{" + item.Notes + "}" : string.Empty;
                                ws.Cells["C" + row].Value = item.Description != null && item.Description != string.Empty ? item.Description : string.Empty;
                                ws.Cells["A" + row + ":B" + row].Merge = true;
                                ws.Cells["C" + row + ":H" + row].Merge = true;
                                xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                                xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                                xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                                xlsService.ExcelSetFontBold(ws, "A" + row + ":B" + row);
                                xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":H" + row);
                                xlsService.ExcelSetHorizontalAlignCenter(ws, "A" + row + ":B" + row);
                                xlsService.ExcelSetFontColorRed(ws, row, 1, row, 1);
                                //set excel row height base on lenght on text
                                xlsService.XlsSetRow(ws, row, item.Description, 380, 30);

                                row++;
                            }

                        }
                        row++;
                        #endregion


                        if ((questionnaireInput.ListIUCSystemGen != null && questionnaireInput.ListIUCSystemGen.Count > 0) ||
                           questionnaireInput.ListIUCNonSystemGen != null && questionnaireInput.ListIUCNonSystemGen.Count > 0)
                        {
                            #region IUC Header
                            row += 2;
                            ws.Cells["A" + row].Value = "IUC";
                            ws.Cells["A" + row + ":F" + row].Merge = true;
                            xlsService.ExcelWrapText(ws, "A" + row + ":F" + row);
                            xlsService.ExcelSetBorder(ws, "A" + row + ":F" + row);
                            xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":F" + row);
                            xlsService.ExcelSetFontBold(ws, "A" + row + ":F" + row);
                            xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":F" + row);
                            xlsService.ExcelSetHorizontalAlignCenter(ws, "A" + row + ":F" + row);
                            row++;

                            ws.Cells["A" + row].Value = "Name of Report";
                            ws.Cells["B" + row].Value = "IUC Type";
                            ws.Cells["C" + row].Value = "Source Control";
                            ws.Cells["D" + row].Value = "IUC Validation";
                            ws.Cells["D" + row + ":F" + row].Merge = true;
                            xlsService.ExcelSetBorder(ws, "A" + row + ":F" + row);
                            xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":F" + row);
                            xlsService.ExcelSetFontBold(ws, "A" + row + ":F" + row);
                            xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":F" + row);
                            xlsService.ExcelSetHorizontalAlignCenter(ws, "A" + row + ":F" + row);
                            row++;

                            #endregion
                        }


                        #region IUC System Generated
                        if (questionnaireInput.ListIUCSystemGen != null && questionnaireInput.ListIUCSystemGen.Count > 0)
                        {
                            foreach (var item in questionnaireInput.ListIUCSystemGen)
                            {
                                string reportName = item.ListQuestionAnswer.Where(x =>
                                    x.Question.ToLower().Contains("what is the report name")).Select(x => x.Answer).FirstOrDefault();

                                string sourceControl = item.ListQuestionAnswer.Where(x =>
                                    x.Question.ToLower().Contains("which system did this report come from, if any? if manually compiled, put spreadsheet")).Select(x => x.Answer).FirstOrDefault();

                                //string general = item.ListQuestionAnswer.Where(x => 
                                //    x.Question.ToLower().Contains("how is the iuc generated")).Select(x => x.Answer).FirstOrDefault();

                                string general = item.ListQuestionAnswer.Where(x =>
                                    x.Question.ToLower().Contains("how is the information used in control (iuc) or what documents are used by the control owner to do their task")).Select(x => x.Answer).FirstOrDefault();

                                //string reportParam = item.ListQuestionAnswer.Where(x => 
                                //    x.Question.ToLower().Contains("what is/are the evidence(s) obtained to address completeness and accuracy of user-entered parameters")).Select(x => x.Answer).FirstOrDefault();

                                string reportParam = item.ListQuestionAnswer.Where(x =>
                                    x.Question.ToLower().Contains("what evidence can you get to show the report parameters used such as report date or data filter? did control owners check if the report parameters are correct")).Select(x => x.Answer).FirstOrDefault();


                                //string sourceData1 = item.ListQuestionAnswer.Where(x => 
                                //    x.Question.ToLower().Contains("how do/does the control(s) address completeness and accuracy")).Select(x => x.Answer).FirstOrDefault();

                                string sourceData1 = item.ListQuestionAnswer.Where(x =>
                                    x.Question.ToLower().Contains("did the control owner check that the report they used is complete and accurate")).Select(x => x.Answer).FirstOrDefault();

                                //string sourceData2 = item.ListQuestionAnswer.Where(x => 
                                //    x.Question.ToLower().Contains("is/are the control(s) designed and operating effectively (completeness)")).Select(x => x.Answer).FirstOrDefault();

                                string sourceData2 = item.ListQuestionAnswer.Where(x =>
                                    x.Question.ToLower().Contains("is/are the control(s) designed and operating effectively for data completeness")).Select(x => x.Answer).FirstOrDefault();


                                //string sourceData3 = item.ListQuestionAnswer.Where(x => 
                                //    x.Question.ToLower().Contains("is/are the control(s) designed and operating effectively (accuracy)")).Select(x => x.Answer).FirstOrDefault();

                                string sourceData3 = item.ListQuestionAnswer.Where(x =>
                                    x.Question.ToLower().Contains("is/are the control(s) designed and operating effectively for data accuracy")).Select(x => x.Answer).FirstOrDefault();


                                //string reportLogic1 = item.ListQuestionAnswer.Where(x => 
                                //    x.Question.ToLower().Contains("how does/do the control(s) address the report logic")).Select(x => x.Answer).FirstOrDefault();

                                string reportLogic1 = item.ListQuestionAnswer.Where(x =>
                                    x.Question.ToLower().Contains("how does/do the control(s) address the report logic or formulas used")).Select(x => x.Answer).FirstOrDefault();


                                //string reportLogic2 = item.ListQuestionAnswer.Where(x => 
                                //    x.Question.ToLower().Contains("is/are the control(s) designed and operating effectively")).Select(x => x.Answer).FirstOrDefault();

                                string reportLogic2 = item.ListQuestionAnswer.Where(x =>
                                    x.Question.ToLower().Contains("is/are the control(s) designed and operating effectively for report logic or formulas used")).Select(x => x.Answer).FirstOrDefault();


                                int startRow = row;

                                ws.Cells["A" + row].Value = reportName != null ? reportName : string.Empty;
                                ws.Cells["B" + row].Value = "System Generated";
                                ws.Cells["C" + row].Value = sourceControl != null ? sourceControl : string.Empty;
                                ws.Cells["D" + row].Value = "General";
                                ws.Cells["E" + row].Value = "Description of How the IUC Is Generated";
                                ws.Cells["F" + row].Value = general != null ? general : string.Empty;
                                xlsService.ExcelWrapText(ws, row, 1, row, 6);
                                xlsService.ExcelSetBorder(ws, row, 1, row, 6);
                                xlsService.ExcelSetArialSize10(ws, row, 1, row, 6);
                                xlsService.ExcelSetVerticalAlignTop(ws, row, 1, row, 6);
                                xlsService.ExcelSetHorizontalAlignLeft(ws, row, 1, row, 6);
                                //ws.Row(row).Height = 70;
                                xlsService.XlsSetRow(ws, row, general, 120, 70);
                                row++;

                                ws.Cells["D" + row].Value = "Report Parameters";
                                ws.Cells["E" + row].Value = "Evidence Obtained to Address Completeness and Accuracy of User-Entered Parameters";
                                ws.Cells["F" + row].Value = reportParam != null ? reportParam : string.Empty;
                                xlsService.ExcelWrapText(ws, row, 4, row, 6);
                                xlsService.ExcelSetBorder(ws, row, 4, row, 6);
                                xlsService.ExcelSetArialSize10(ws, row, 4, row, 6);
                                xlsService.ExcelSetVerticalAlignTop(ws, row, 4, row, 6);
                                xlsService.ExcelSetHorizontalAlignLeft(ws, row, 4, row, 6);
                                //ws.Row(row).Height = 70;
                                xlsService.XlsSetRow(ws, row, reportParam, 120, 70);
                                row++;

                                #region Source Data

                                int starRowSourceData = row;
                                ws.Cells["D" + row].Value = "Source Data";
                                ws.Cells["E" + row].Value = "Control(s) Addressing Completeness and Accuracy (describe how)";
                                ws.Cells["F" + row].Value = sourceData1 != null ? sourceData1 : string.Empty;
                                xlsService.ExcelWrapText(ws, row, 5, row, 6);
                                xlsService.ExcelSetBorder(ws, row, 5, row, 6);
                                xlsService.ExcelSetArialSize10(ws, row, 5, row, 6);
                                xlsService.ExcelSetVerticalAlignTop(ws, row, 5, row, 6);
                                xlsService.ExcelSetHorizontalAlignLeft(ws, row, 5, row, 6);
                                //ws.Row(row).Height = 70;
                                xlsService.XlsSetRow(ws, row, sourceData1, 120, 70);
                                row++;

                                ws.Cells["E" + row].Value = "Is Control(s) Designed and Operating Effectively (completeness)?";
                                ws.Cells["F" + row].Value = sourceData2 != null ? sourceData2 : string.Empty;
                                xlsService.ExcelWrapText(ws, row, 5, row, 6);
                                xlsService.ExcelSetBorder(ws, row, 5, row, 6);
                                xlsService.ExcelSetArialSize10(ws, row, 5, row, 6);
                                xlsService.ExcelSetVerticalAlignTop(ws, row, 5, row, 6);
                                xlsService.ExcelSetHorizontalAlignLeft(ws, row, 5, row, 6);
                                //ws.Row(row).Height = 70;
                                xlsService.XlsSetRow(ws, row, sourceData2, 120, 70);
                                row++;

                                ws.Cells["E" + row].Value = "Is Control(s) Designed and Operating Effectively (accuracy)?";
                                ws.Cells["F" + row].Value = sourceData3 != null ? sourceData3 : string.Empty;
                                xlsService.ExcelWrapText(ws, row, 5, row, 6);
                                xlsService.ExcelSetBorder(ws, row, 5, row, 6);
                                xlsService.ExcelSetArialSize10(ws, row, 5, row, 6);
                                xlsService.ExcelSetVerticalAlignTop(ws, row, 5, row, 6);
                                xlsService.ExcelSetHorizontalAlignLeft(ws, row, 5, row, 6);
                                //ws.Row(row).Height = 70;
                                xlsService.XlsSetRow(ws, row, sourceData3, 120, 70);

                                int endRowSourceData = row;
                                ws.Cells[starRowSourceData, 4, endRowSourceData, 4].Merge = true;
                                xlsService.ExcelSetBorder(ws, starRowSourceData, 4, endRowSourceData, 4);
                                xlsService.ExcelSetVerticalAlignTop(ws, starRowSourceData, 4, starRowSourceData, 4);
                                xlsService.ExcelSetHorizontalAlignLeft(ws, starRowSourceData, 4, starRowSourceData, 4);

                                row++;

                                #endregion

                                #region Report Logic

                                int startRowReportLogic = row;
                                ws.Cells["D" + row].Value = "Report Logic";
                                ws.Cells["E" + row].Value = "Control(s) Addressing the Report Logic";
                                ws.Cells["F" + row].Value = reportLogic1 != null ? reportLogic1 : string.Empty;
                                xlsService.ExcelWrapText(ws, row, 5, row, 6);
                                xlsService.ExcelSetBorder(ws, row, 5, row, 6);
                                xlsService.ExcelSetArialSize10(ws, row, 5, row, 6);
                                xlsService.ExcelSetVerticalAlignTop(ws, row, 5, row, 6);
                                xlsService.ExcelSetHorizontalAlignLeft(ws, row, 5, row, 6);
                                //ws.Row(row).Height = 70;
                                xlsService.XlsSetRow(ws, row, reportLogic1, 120, 70);
                                row++;

                                ws.Cells["E" + row].Value = "Is Control(s) Designed and Operating Effectively?";
                                ws.Cells["F" + row].Value = reportLogic2 != null ? reportLogic2 : string.Empty;
                                xlsService.ExcelWrapText(ws, row, 5, row, 6);
                                xlsService.ExcelSetBorder(ws, row, 5, row, 6);
                                xlsService.ExcelSetArialSize10(ws, row, 5, row, 6);
                                xlsService.ExcelSetVerticalAlignTop(ws, row, 5, row, 6);
                                xlsService.ExcelSetHorizontalAlignLeft(ws, row, 5, row, 6);
                                //ws.Row(row).Height = 70;
                                xlsService.XlsSetRow(ws, row, reportLogic2, 120, 70);

                                int endRowReportLogic = row;
                                ws.Cells[startRowReportLogic, 4, endRowReportLogic, 4].Merge = true;
                                xlsService.ExcelSetBorder(ws, startRowReportLogic, 4, endRowReportLogic, 4);
                                xlsService.ExcelSetVerticalAlignTop(ws, startRowReportLogic, 4, startRowReportLogic, 4);
                                xlsService.ExcelSetHorizontalAlignLeft(ws, startRowReportLogic, 4, startRowReportLogic, 4);

                                #endregion

                                int endRow = row;

                                ws.Cells[startRow, 1, endRow, 1].Merge = true;
                                ws.Cells[startRow, 2, endRow, 2].Merge = true;
                                ws.Cells[startRow, 3, endRow, 3].Merge = true;
                                xlsService.ExcelSetBorder(ws, startRow, 1, endRow, 1);
                                xlsService.ExcelSetBorder(ws, startRow, 2, endRow, 2);
                                xlsService.ExcelSetBorder(ws, startRow, 3, endRow, 3);

                                row++;
                            }

                        }
                        #endregion

                        #region IUC Non System Generated
                        if (questionnaireInput.ListIUCNonSystemGen != null && questionnaireInput.ListIUCNonSystemGen.Count > 0)
                        {
                            foreach (var item in questionnaireInput.ListIUCNonSystemGen)
                            {
                                string reportName = item.ListQuestionAnswer.Where(x => x.Question.ToLower().Contains("what is the report name")).Select(x => x.Answer).FirstOrDefault();
                                //string sourceControl = item.ListQuestionAnswer.Where(x => x.Question.ToLower().Contains("what is the source control")).Select(x => x.Answer).FirstOrDefault();
                                string sourceControl = item.ListQuestionAnswer.Where(x => x.Question.ToLower().Contains("does this control rely on another control for this control to work? what is the source control")).Select(x => x.Answer).FirstOrDefault();
                                string general = item.ListQuestionAnswer.Where(x => x.Question.ToLower().Contains("how do/does the control(s) address the completeness and accuracy of the non-system-generated report")).Select(x => x.Answer).FirstOrDefault();

                                ws.Cells["A" + row].Value = reportName != null ? reportName : string.Empty;
                                ws.Cells["B" + row].Value = "Non System Generated";
                                ws.Cells["C" + row].Value = sourceControl != null ? sourceControl : string.Empty;
                                ws.Cells["D" + row].Value = "General";
                                ws.Cells["E" + row].Value = "Describe how the controls that addresses the completeness and accuracy of the Non-System-Generated Report";
                                ws.Cells["F" + row].Value = general != null ? general : string.Empty;
                                xlsService.ExcelWrapText(ws, row, 1, row, 6);
                                xlsService.ExcelSetBorder(ws, row, 1, row, 6);
                                xlsService.ExcelSetArialSize10(ws, row, 1, row, 6);
                                xlsService.ExcelSetVerticalAlignTop(ws, row, 1, row, 6);
                                xlsService.ExcelSetHorizontalAlignLeft(ws, row, 1, row, 6);
                                //ws.Row(row).Height = 70;
                                xlsService.XlsSetRow(ws, row, general, 120, 70);
                                row++;
                            }
                        }
                        #endregion



                        //ws.Protection.IsProtected = true;
                        //ws.Protection.AllowAutoFilter = false;
                        //ws.Protection.AllowDeleteColumns = false;
                        //ws.Protection.AllowDeleteRows = false;
                        //ws.Protection.AllowEditObject = true;
                        //ws.Protection.AllowEditScenarios = true;
                        //ws.Protection.AllowFormatCells = false;
                        //ws.Protection.AllowFormatColumns = false;
                        //ws.Protection.AllowFormatRows = false;
                        //ws.Protection.AllowInsertColumns = false;
                        //ws.Protection.AllowInsertHyperlinks = false;
                        //ws.Protection.AllowInsertRows = false;
                        //ws.Protection.AllowPivotTables = false;
                        //ws.Protection.AllowSelectLockedCells = true;
                        //ws.Protection.AllowSelectUnlockedCells = true;
                        //ws.Protection.AllowSort = false;
                        //ws.Protection.SetPassword("123qwe!!");

                        //string startupPath = Directory.GetCurrentDirectory();
                        //string strSourceDownload = Path.Combine(startupPath, "include", "questionnaire", "download"); 031421
                        string strSourceDownload = Path.Combine(startupPath, "include", "upload","soxquestionnaire");

                        if (!Directory.Exists(strSourceDownload))
                        {
                            Directory.CreateDirectory(strSourceDownload);
                        }
                        var ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                        //string filename = $"{Guid.NewGuid()}.xlsx";
                        string Fy = questionnaireInput.Rcm.FY;
                        string process = questionnaireInput.Rcm.Process;
                        //string filename = $"{clientName} {Fy} {process} {roundName}.xlsx"; 031421
                        string filename = $"{clientName} {Fy} {process} -  {controlId} {roundName}.xlsx";
                        string strOutput = Path.Combine(strSourceDownload, filename);

                        //Check if file not exists
                        if (System.IO.File.Exists(strOutput))
                        {
                            System.IO.File.Delete(strOutput);
                        }

                        xls.SaveAs(new FileInfo(strOutput));
                        excelFilename = filename;
                    }


                }

            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorCreateQuestionnaireExcel");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "CreateQuestionnaireExcel");
                Debug.WriteLine(ex);
            }

            return excelFilename;
        }

        [HttpPost("excel/create")]
        public string CreateQuestionnaireExcel([FromBody] QuestionnaireRoundSet roundSet)
        {
            //List<string> excelFilename = new List<string>();
            string excelFilename = string.Empty;
            try
            {

                string startupPath = Directory.GetCurrentDirectory();

                //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage xls = new ExcelPackage())
                {

                    var controlName = roundSet.Rcm.ControlId;
                    if (controlName != null)
                    {
                        var ws = xls.Workbook.Worksheets.Add(controlName);

                        ws.View.ShowGridLines = false;


                        ExcelService xlsService = new ExcelService();
                        FormatService txtFormat = new FormatService();
                        ws.Column(1).Width = 30;
                        ws.Column(2).Width = 30;
                        ws.Column(3).Width = 22;
                        ws.Column(4).Width = 22;
                        ws.Column(5).Width = 22;
                        ws.Column(6).Width = 22;
                        ws.Column(7).Width = 22;
                        ws.Column(8).Width = 22;
                        ws.Column(9).Width = 22;
                        ws.Column(10).Width = 22;
                        ws.Column(11).Width = 22;
                        ws.Column(12).Width = 22;
                        ws.Column(13).Width = 22;
                        ws.Column(14).Width = 22;
                        ws.Column(15).Width = 22;


                        int row = 1; //Row 1 Client Name
                        string clientName = roundSet.Rcm.ClientName;
                        ws.Cells["A" + row].Value = clientName ?? string.Empty;
                        xlsService.ExcelSetArialSize10(ws, "A" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);

                        row++; //Row 2 Testing Workpapers
                        ws.Cells["A" + row].Value = "Testing Workpapers";
                        xlsService.ExcelSetFontBold(ws, "A" + row);


                        row++; //Row 3 Process
                        string process = roundSet.Rcm.Process ?? string.Empty;
                        ws.Cells["A" + row].Value = process != null ? process : string.Empty;
                        xlsService.ExcelSetArialSize10(ws, "A" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);


                        row++; //Row 4 SOX FY(yy)
                        string soxFy = $"SOX {roundSet.Rcm.FY}";
                        ws.Cells["A" + row].Value = soxFy != null ? soxFy : string.Empty;
                        xlsService.ExcelSetArialSize10(ws, "A" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);

                        //row++; //Row 4 Control Id
                        string controlId = controlName;
                        //ws.Cells["A" + row].Value = controlId != null ? controlId : string.Empty;
                        //xlsService.ExcelSetArialSize10(ws, "A" + row);
                        //xlsService.ExcelSetFontBold(ws, "A" + row);

                        //Row 
                        row += 2;
                        ws.Cells["A" + row].Value = "Control Details";
                        ws.Cells["A" + row + ":J" + row].Merge = true;
                        xlsService.ExcelSetBorder(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);
                        xlsService.ExcelSetAlignCenter(ws, "A" + row);

                        //Row 
                        row += 1;
                        ws.Cells["A" + row].Value = "Process:";
                        ws.Cells["C" + row].Value = process != null ? process : string.Empty;
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":J" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);
                        //set excel row height base on lenght on text
                        xlsService.XlsSetRow(ws, row, roundSet.Rcm.Process, 380, 15);
                        xlsService.ExcelSetBackgroundColorYellowPaper(ws, "C" + row + ":J" + row);

                        //Row 7 "who is the control owner"
                        row++;
                        //string controlOwner = roundSet.Rcm.ControlOwner;
                        var controlOwner = roundSet.ListUserInputRound1.Where(x => x.StrQuestion.ToLower().Contains("who is the control owner?")).Select(x => x.StrAnswer).FirstOrDefault();
                        if (controlOwner != null)
                        {
                            ws.Cells["A" + row].Value = "Control Owner:";
                            //ws.Cells["C" + row].Value = controlOwner != null ? controlOwner : string.Empty;
                            ws.Cells["C" + row].Value = controlOwner != null ? controlOwner : string.Empty;
                            ws.Cells["A" + row + ":B" + row].Merge = true;
                            ws.Cells["C" + row + ":J" + row].Merge = true;
                            xlsService.ExcelWrapText(ws, "A" + row + ":J" + row);
                            xlsService.ExcelSetBorder(ws, "A" + row + ":J" + row);
                            xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                            xlsService.ExcelSetArialSize10(ws, "A" + row + ":J" + row);
                            xlsService.ExcelSetFontBold(ws, "A" + row);
                            //set excel row height base on lenght on text
                            xlsService.XlsSetRow(ws, row, controlOwner, 380, 15);
                            xlsService.ExcelSetBackgroundColorYellowPaper(ws, "C" + row + ":J" + row);
                        }




                        //Row 
                        row++;
                        ws.Cells["A" + row].Value = "Control ID:";
                        ws.Cells["C" + row].Value = controlId != null ? controlId : string.Empty;
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":J" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);
                        //set excel row height base on lenght on text
                        xlsService.XlsSetRow(ws, row, controlId, 380, 15);
                        xlsService.ExcelSetBackgroundColorYellowPaper(ws, "C" + row + ":J" + row);

                        //New Row  Control Description
                        row++;
                        string controlShortDesc = roundSet.Rcm.ShortDescription;
                        ws.Cells["A" + row].Value = "Control Short Description:";
                        ws.Cells["C" + row].Value = controlShortDesc;
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":J" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);
                        xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBackgroundColorYellowPaper(ws, "C" + row + ":J" + row);
                        xlsService.XlsSetRow(ws, row, controlShortDesc, 380, 15);

                        //Row  "what is the control activity"
                        row++;
                        string controlDesc = roundSet.Rcm.ControlActivityFy19;
                        ws.Cells["A" + row].Value = "Control Description:";
                        ws.Cells["C" + row].Value = controlDesc;
                        
                        //ws.Cells["C" + row].Value = controlDesc != null ? txtFormat.ReplaceTagHtmlParagraph(controlDesc, true) : string.Empty;
                        ws.Cells["C" + row].Value = controlDesc != null ? txtFormat.FormatwithNewLine(controlDesc, true) : string.Empty;
                        
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":J" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);
                        xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBackgroundColorYellowPaper(ws, "C" + row + ":J" + row);
                        //ws.Row(row).Height = 60;

                        //Add freeze pane
                        ws.View.FreezePanes(12, 3);

                        //set excel row height base on lenght on text
                        xlsService.XlsSetRow(ws, row, controlDesc, 300, 60);

                        //New Row  "control type"
                        row++;
                        string controlType = roundSet.Rcm.ControlType;
                        //string controlType = "";
                        ws.Cells["A" + row].Value = "Control Type:";
                        ws.Cells["C" + row].Value = controlType;
                    
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":J" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);
                        xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBackgroundColorYellowPaper(ws, "C" + row + ":J" + row);
                        xlsService.XlsSetRow(ws, row, controlType, 380, 15);



                        //Row  "what are the procedures to test control" or "what are the procedures to test this control"
                        row++;
                        string testValidation = roundSet.Rcm.TestProc;
                        ws.Cells["A" + row].Value = "Test Validation Approach:";
                        
                         ws.Cells["C" + row].Value = testValidation != null ? txtFormat.FormatwithNewLine(testValidation, true) : string.Empty;
                        //ws.Cells["C" + row].Value = testValidation != null ? txtFormat.ReplaceTagHtmlParagraph(testValidation, true) : string.Empty;
                        //ws.Cells["C" + row].Value = testValidation;
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":J" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);
                        xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":J" + row);
                        //ws.Row(row).Height = 60;
                        //set excel row height base on lenght on text
                        xlsService.XlsSetRow(ws, row, testValidation, 300, 60);
                        xlsService.ExcelSetBackgroundColorYellowPaper(ws, "C" + row + ":J" + row);



                        //Row 12
                        //15. What procedures did you use to select the samples?
                        //string methodUsedA = questionnaireInput.ListUserInputItem.Where(x =>
                        //    x.StrQuestion.ToLower().Contains("procedures did you use to select the samples?") ||
                        //    x.StrQuestion.ToLower().Contains("what is the test method used?")
                        //).Select(x => x.StrAnswer).FirstOrDefault();


                        string methodUsedA = string.Empty;
                        if (roundSet.isRound1)
                            methodUsedA = roundSet.ListUserInputRound1.Where(x => x.StrQuestion.ToLower().Contains("procedures did you use to select the samples?") || x.StrQuestion.ToLower().Contains("what is the test method used?")).Select(x => x.StrAnswer).FirstOrDefault();
                        else if (roundSet.isRound2)
                            methodUsedA = roundSet.ListUserInputRound2.Where(x => x.StrQuestion.ToLower().Contains("procedures did you use to select the samples?") || x.StrQuestion.ToLower().Contains("what is the test method used?")).Select(x => x.StrAnswer).FirstOrDefault();
                        else if (roundSet.isRound3)
                            methodUsedA = roundSet.ListUserInputRound3.Where(x => x.StrQuestion.ToLower().Contains("procedures did you use to select the samples?") || x.StrQuestion.ToLower().Contains("what is the test method used?")).Select(x => x.StrAnswer).FirstOrDefault();

                        if (methodUsedA != string.Empty)
                        {
                            row++;
                            ws.Cells["A" + row].Value = "Test Method Used:";
                            ws.Cells["C" + row].Value = methodUsedA != null ? txtFormat.ReplaceTagHtmlParagraph(methodUsedA, true) : string.Empty;
                            ws.Cells["A" + row + ":B" + row].Merge = true;
                            ws.Cells["C" + row + ":J" + row].Merge = true;
                            xlsService.ExcelWrapText(ws, "A" + row + ":J" + row);
                            xlsService.ExcelSetBorder(ws, "A" + row + ":J" + row);
                            xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                            xlsService.ExcelSetArialSize10(ws, "A" + row + ":J" + row);
                            xlsService.ExcelSetFontBold(ws, "A" + row);
                            //set excel row height base on lenght on text
                            xlsService.XlsSetRow(ws, row, roundSet.Rcm.TestProc, 300, 30);
                            xlsService.ExcelSetBackgroundColorYellowPaper(ws, "C" + row + ":J" + row);
                        }


                        //Row 13 "how often does this control happen"
                        row++;
                        string controlFreq = string.Empty;
                        //string controlFreq = roundSet.Rcm.ControlFrequency;
                        /*
                        var controlFreq = roundSet.ListUserInputRound1.Where(x => x.StrQuestion.ToLower().Contains("how often does this control happen")).Select(x => x.StrAnswer).FirstOrDefault();
                        if(controlFreq != null)
                        {
                            ws.Cells["A" + row].Value = "Control Frequency:";
                            //ws.Cells["C" + row].Value = controlFreq != null ? controlFreq : string.Empty;
                            ws.Cells["C" + row].Value = controlFreq != null ? controlFreq : string.Empty;
                            ws.Cells["A" + row + ":B" + row].Merge = true;
                            ws.Cells["C" + row + ":H" + row].Merge = true;
                            xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                            xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                            xlsService.ExcelSetFontBold(ws, "A" + row);
                            xlsService.ExcelSetBackgroundColorYellowPaper(ws, "C" + row + ":H" + row);
                        }
                        */
                        //
                        if (roundSet.isRound1) {
                            controlFreq = roundSet.ListUserInputRound1.Where(x => x.StrQuestion.ToLower().Contains("how often does this control happen")).Select(x => x.StrAnswer).FirstOrDefault();
                        }
                        else if (roundSet.isRound2) {
                            controlFreq = roundSet.ListUserInputRound2.Where(x => x.StrQuestion.ToLower().Contains("how often does this control happen")).Select(x => x.StrAnswer).FirstOrDefault();
                        }
                        else if (roundSet.isRound3) {
                            controlFreq = roundSet.ListUserInputRound3.Where(x => x.StrQuestion.ToLower().Contains("how often does this control happen")).Select(x => x.StrAnswer).FirstOrDefault();
                        }
                        else {
                            controlFreq = roundSet.Rcm.ControlFrequency;
                        }
                        ws.Cells["A" + row].Value = "Control Frequency:";
                        //ws.Cells["C" + row].Value = controlFreq != null ? controlFreq : string.Empty;
                        ws.Cells["C" + row].Value = controlFreq != null ? controlFreq : string.Empty;
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":J" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);
                        xlsService.ExcelSetBackgroundColorYellowPaper(ws, "C" + row + ":H" + row);

                        //



                        //Row 14
                        //6. When was the control first put in place? Enter date like 1/1/xx.
                        row++;
                        string controlPlaceDate = string.Empty;
                        if (roundSet.isRound1)
                            controlPlaceDate = roundSet.ListUserInputRound1.Where(x => x.StrQuestion.ToLower().Contains("control first put in place?")).Select(x => x.StrAnswer).FirstOrDefault();
                        else if (roundSet.isRound2)
                            controlPlaceDate = roundSet.ListUserInputRound2.Where(x => x.StrQuestion.ToLower().Contains("control first put in place?")).Select(x => x.StrAnswer).FirstOrDefault();
                        else if (roundSet.isRound3)
                            controlPlaceDate = roundSet.ListUserInputRound3.Where(x => x.StrQuestion.ToLower().Contains("control first put in place?")).Select(x => x.StrAnswer).FirstOrDefault();
                        string parseControlPlaceDate = string.Empty;

                        if (controlPlaceDate != null && controlPlaceDate != string.Empty)
                        {
                            DateTime tempcontrolPlaceDate;
                            //try to parse string to date time
                            if (DateTime.TryParse(controlPlaceDate, out tempcontrolPlaceDate))
                            {
                                parseControlPlaceDate = tempcontrolPlaceDate.ToString("MM/dd/yyyy");
                            }
                            else
                            {
                                //if it failed to parse date time, then we will just remove time string
                                int index = controlPlaceDate.IndexOf(" ");
                                if (index > 0)
                                {
                                    parseControlPlaceDate = controlPlaceDate.Substring(0, index);
                                }
                            }
                        }

                        ws.Cells["A" + row].Value = "Control In Place Date:";
                        ws.Cells["C" + row].Style.Numberformat.Format = "@";
                        ws.Cells["C" + row].Value = parseControlPlaceDate != null ? parseControlPlaceDate : string.Empty;
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":J" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);
                        xlsService.ExcelSetBackgroundColorYellowPaper(ws, "C" + row + ":J" + row);


                        //Row 15 "what is the level of risk for the control"
                        row++;
                        string levelRisk = string.Empty;

                        if (roundSet.isRound1)
                            levelRisk = roundSet.ListUserInputRound1.Where(x => x.StrQuestion.ToLower().Contains("what is the level of risk for the control")).Select(x => x.StrAnswer).FirstOrDefault();
                        else if (roundSet.isRound2)
                            levelRisk = roundSet.ListUserInputRound2.Where(x => x.StrQuestion.ToLower().Contains("what is the level of risk for the control")).Select(x => x.StrAnswer).FirstOrDefault();
                        else if (roundSet.isRound3)
                            levelRisk = roundSet.ListUserInputRound3.Where(x => x.StrQuestion.ToLower().Contains("what is the level of risk for the control")).Select(x => x.StrAnswer).FirstOrDefault();
                        else
                            levelRisk = roundSet.Rcm.RiskLvl;

                        ws.Cells["A" + row].Value = "Risk Assessment:";
                        ws.Cells["C" + row].Value = levelRisk != null ? levelRisk : string.Empty;
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":J" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);
                        xlsService.ExcelSetBackgroundColorYellowPaper(ws, "C" + row + ":J" + row);

                        row++;
                        //Row 16
                        //12.What date ranges are you selecting samples?
                        string samplePeriod = string.Empty;
                        string samplePeriod2 = string.Empty;

                        if (roundSet.isRound1)
                        {
                            samplePeriod = roundSet.ListUserInputRound1.Where(x => x.StrQuestion.ToLower().Contains("what date ranges are you selecting samples")).Select(x => x.StrAnswer).FirstOrDefault();
                            samplePeriod2 = roundSet.ListUserInputRound1.Where(x => x.StrQuestion.ToLower().Contains("what date ranges are you selecting samples")).Select(x => x.StrAnswer2).FirstOrDefault();
                        }
                        else if (roundSet.isRound2)
                        {
                            samplePeriod = roundSet.ListUserInputRound2.Where(x => x.StrQuestion.ToLower().Contains("what date ranges are you selecting samples")).Select(x => x.StrAnswer).FirstOrDefault();
                            samplePeriod2 = roundSet.ListUserInputRound2.Where(x => x.StrQuestion.ToLower().Contains("what date ranges are you selecting samples")).Select(x => x.StrAnswer2).FirstOrDefault();
                        }
                        else if (roundSet.isRound3)
                        {
                            samplePeriod = roundSet.ListUserInputRound3.Where(x => x.StrQuestion.ToLower().Contains("what date ranges are you selecting samples")).Select(x => x.StrAnswer).FirstOrDefault();
                            samplePeriod2 = roundSet.ListUserInputRound3.Where(x => x.StrQuestion.ToLower().Contains("what date ranges are you selecting samples")).Select(x => x.StrAnswer2).FirstOrDefault();
                        }


                        string parsesamplePeriod = string.Empty;
                        string parsesamplePeriod2 = string.Empty;
                        if (samplePeriod != null && samplePeriod != string.Empty)
                        {
                            //parsesamplePeriod = DateTime.Parse(samplePeriod).ToString("MM/dd/yyyy");
                            DateTime tempsamplePeriod;
                            //try to parse string to date time
                            if (DateTime.TryParse(samplePeriod, out tempsamplePeriod))
                            {
                                parsesamplePeriod = tempsamplePeriod.ToString("MM/dd/yyyy");
                            }
                            else
                            {
                                //if it failed to parse date time, then we will just remove time string
                                int index = samplePeriod.IndexOf(" ");
                                if (index > 0)
                                {
                                    parsesamplePeriod = samplePeriod.Substring(0, index);
                                }
                            }
                        }
                        if (samplePeriod2 != null && samplePeriod2 != string.Empty)
                        {
                            //parsesamplePeriod2 = DateTime.Parse(samplePeriod2).ToString("MM/dd/yyyy");
                            DateTime tempsamplePeriod2;
                            //try to parse string to date time
                            if (DateTime.TryParse(samplePeriod2, out tempsamplePeriod2))
                            {
                                parsesamplePeriod2 = tempsamplePeriod2.ToString("MM/dd/yyyy");
                            }
                            else
                            {
                                //if it failed to parse date time, then we will just remove time string
                                int index = samplePeriod2.IndexOf(" ");
                                if (index > 0)
                                {
                                    parsesamplePeriod2 = samplePeriod2.Substring(0, index);
                                }
                            }
                        }
                        string validateSamplePeriod = string.Empty;
                        if (parsesamplePeriod != null && parsesamplePeriod != string.Empty)
                        {
                            validateSamplePeriod = (parsesamplePeriod != null ? parsesamplePeriod : string.Empty) + " to " + (parsesamplePeriod2 != null ? parsesamplePeriod2 : string.Empty);
                        }

                        ws.Cells["A" + row].Value = "Sample Period:";
                        ws.Cells["C" + row].Style.Numberformat.Format = "@";
                        ws.Cells["C" + row].Value = validateSamplePeriod;
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":J" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);
                        xlsService.ExcelSetBackgroundColorYellowPaper(ws, "C" + row + ":J" + row);



                        //Row 17
                        row++;
                        string populationSize = string.Empty;
                        if (roundSet.isRound1)
                            populationSize = roundSet.ListUserInputRound1.Where(x => x.StrQuestion.ToLower().Contains("what is the population size?")).Select(x => x.StrAnswer).FirstOrDefault();
                        else if (roundSet.isRound2)
                            populationSize = roundSet.ListUserInputRound2.Where(x => x.StrQuestion.ToLower().Contains("what is the population size?")).Select(x => x.StrAnswer).FirstOrDefault();
                        else if (roundSet.isRound3)
                            populationSize = roundSet.ListUserInputRound3.Where(x => x.StrQuestion.ToLower().Contains("what is the population size?")).Select(x => x.StrAnswer).FirstOrDefault();

                        ws.Cells["A" + row].Value = "Population Size:";
                        ws.Cells["C" + row].Value = populationSize != null ? populationSize : string.Empty;
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":J" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);
                        xlsService.ExcelSetBackgroundColorYellowPaper(ws, "C" + row + ":J" + row);

                        //Row 18
                        row++;
                        string sampleDerivation = string.Empty;

                        if (roundSet.isRound1)
                            //sampleDerivation = roundSet.ListUserInputRound1.Where(x => x.StrQuestion.ToLower().Contains("what is the sub-sample size")).Select(x => x.StrAnswer).FirstOrDefault();
                            sampleDerivation = roundSet.ListUserInputRound1.Where(x => x.StrQuestion.ToLower().Contains("what is the sub-sample size") || x.StrQuestion.ToLower().Contains("what is the sample size")).Select(x => x.StrAnswer).FirstOrDefault();
                        else if (roundSet.isRound2)
                            //sampleDerivation = roundSet.ListUserInputRound2.Where(x => x.StrQuestion.ToLower().Contains("what is the sub-sample size")).Select(x => x.StrAnswer).FirstOrDefault();
                            sampleDerivation = roundSet.ListUserInputRound2.Where(x => x.StrQuestion.ToLower().Contains("what is the sub-sample size") || x.StrQuestion.ToLower().Contains("what is the sample size")).Select(x => x.StrAnswer).FirstOrDefault();
                        else if (roundSet.isRound3)
                            //sampleDerivation = roundSet.ListUserInputRound3.Where(x => x.StrQuestion.ToLower().Contains("what is the sub-sample size")).Select(x => x.StrAnswer).FirstOrDefault();
                            sampleDerivation = roundSet.ListUserInputRound3.Where(x => x.StrQuestion.ToLower().Contains("what is the sub-sample size") || x.StrQuestion.ToLower().Contains("what is the sample size")).Select(x => x.StrAnswer).FirstOrDefault();

                        ws.Cells["A" + row].Value = "Sample Derivation:";
                        ws.Cells["C" + row].Value = sampleDerivation != null ? sampleDerivation : string.Empty;
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":J" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);
                        xlsService.ExcelSetBackgroundColorYellowPaper(ws, "C" + row + ":J" + row);

                        //dont display "Electronic Audit Evidence" if answer is blank
                        string auditEvidenceQ = string.Empty;

                        if (roundSet.isRound1)
                            auditEvidenceQ = roundSet.ListUserInputRound1.Where(x => x.StrQuestion.ToLower().Contains("(ipe)")).Select(x => x.StrQuestion).FirstOrDefault();
                        else if (roundSet.isRound2)
                            auditEvidenceQ = roundSet.ListUserInputRound2.Where(x => x.StrQuestion.ToLower().Contains("(ipe)")).Select(x => x.StrQuestion).FirstOrDefault();
                        else if (roundSet.isRound3)
                            auditEvidenceQ = roundSet.ListUserInputRound3.Where(x => x.StrQuestion.ToLower().Contains("(ipe)")).Select(x => x.StrQuestion).FirstOrDefault();

                        if (auditEvidenceQ != null && auditEvidenceQ != string.Empty)
                        {

                            string auditEvidenceA = string.Empty;

                            if (roundSet.isRound1)
                                auditEvidenceA = roundSet.ListUserInputRound1.Where(x => x.StrQuestion.ToLower().Contains("(ipe)")).Select(x => x.StrAnswer).FirstOrDefault();
                            else if (roundSet.isRound2)
                                auditEvidenceA = roundSet.ListUserInputRound2.Where(x => x.StrQuestion.ToLower().Contains("(ipe)")).Select(x => x.StrAnswer).FirstOrDefault();
                            else if (roundSet.isRound3)
                                auditEvidenceA = roundSet.ListUserInputRound3.Where(x => x.StrQuestion.ToLower().Contains("(ipe)")).Select(x => x.StrAnswer).FirstOrDefault();

                            //Row 19
                            row++;
                            ws.Cells["A" + row].Value = "Electronic Audit Evidence:";
                            ws.Cells["C" + row].Value = auditEvidenceA != null ? txtFormat.FormatwithNewLine(auditEvidenceA, true) : string.Empty;
                            ws.Cells["A" + row + ":B" + row].Merge = true;
                            ws.Cells["C" + row + ":J" + row].Merge = true;
                            xlsService.ExcelWrapText(ws, "A" + row + ":J" + row);
                            xlsService.ExcelSetBorder(ws, "A" + row + ":J" + row);
                            xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                            xlsService.ExcelSetArialSize10(ws, "A" + row + ":J" + row);
                            xlsService.ExcelSetFontBold(ws, "A" + row);
                        }


                        row += 2;
                        //Row 21
                        #region Round Header
                        ws.Cells["C" + row].Value = "Walkthrough (WT):";
                        ws.Cells["C" + row + ":D" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "C" + row + ":D" + row);
                        xlsService.ExcelSetBorder(ws, "C" + row + ":D" + row);

                        ws.Cells["E" + row].Value = "Round 1:";
                        ws.Cells["E" + row + ":F" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "E" + row + ":F" + row);
                        xlsService.ExcelSetBorder(ws, "E" + row + ":F" + row);

                        ws.Cells["G" + row].Value = "Round 2:";
                        ws.Cells["G" + row + ":H" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "G" + row + ":H" + row);
                        xlsService.ExcelSetBorder(ws, "G" + row + ":H" + row);

                        ws.Cells["I" + row].Value = "Round 3:";
                        ws.Cells["I" + row + ":J" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "I" + row + ":J" + row);
                        xlsService.ExcelSetBorder(ws, "I" + row + ":J" + row);

                        xlsService.ExcelSetBackgroundColorGray(ws, "C" + row + ":J" + row);
                        xlsService.ExcelSetArialSize10(ws, "C" + row + ":J" + row);
                        xlsService.ExcelSetFontBold(ws, "C" + row + ":J" + row);
                        xlsService.ExcelSetVerticalAlignCenter(ws, "C" + row + ":J" + row);
                        xlsService.ExcelSetHorizontalAlignCenter(ws, "C" + row + ":J" + row);
                        #endregion



                        //Row 
                        string sampleSizeWT = string.Empty;
                        string sampleSize1 = string.Empty;
                        string sampleSize2 = string.Empty;
                        string sampleSize3 = string.Empty;
                        /*
                        sampleSize1 = roundSet.ListUserInputRound1.Where(x => x.StrQuestion.ToLower().Contains("what is the sub-sample size") || x.StrQuestion.ToLower().Contains("what is the sub-sample size?")).Select(x => x.StrAnswer).FirstOrDefault();
                        sampleSize2 = roundSet.ListUserInputRound2.Where(x => x.StrQuestion.ToLower().Contains("what is the sub-sample size") || x.StrQuestion.ToLower().Contains("what is the sub-sample size?")).Select(x => x.StrAnswer).FirstOrDefault();
                        sampleSize3 = roundSet.ListUserInputRound3.Where(x => x.StrQuestion.ToLower().Contains("what is the sub-sample size") || x.StrQuestion.ToLower().Contains("what is the sub-sample size?")).Select(x => x.StrAnswer).FirstOrDefault();
                        */
                        //sampleSizeWT = roundSet.ListUserInputRoundWT.Where(x => x.StrQuestion.ToLower().Contains("what is the sub-sample size") || x.StrQuestion.ToLower().Contains("what is the sample size")).Select(x => x.StrAnswer).FirstOrDefault();
                        sampleSize1 = roundSet.ListUserInputRound1.Where(x => x.StrQuestion.ToLower().Contains("what is the sub-sample size") || x.StrQuestion.ToLower().Contains("what is the sample size")).Select(x => x.StrAnswer).FirstOrDefault();
                        sampleSize2 = roundSet.ListUserInputRound2.Where(x => x.StrQuestion.ToLower().Contains("what is the sub-sample size") || x.StrQuestion.ToLower().Contains("what is the sample size")).Select(x => x.StrAnswer).FirstOrDefault();
                        sampleSize3 = roundSet.ListUserInputRound3.Where(x => x.StrQuestion.ToLower().Contains("what is the sub-sample size") || x.StrQuestion.ToLower().Contains("what is the sample size")).Select(x => x.StrAnswer).FirstOrDefault();

                        row++;
                        ws.Cells["A" + row].Value = "Sample Size:";
                        ws.Cells["C" + row].Value = (sampleSizeWT != null ? sampleSizeWT : string.Empty);
                        ws.Cells["E" + row].Value = (sampleSize1 != null ? sampleSize1 : string.Empty);
                        ws.Cells["G" + row].Value = (sampleSize2 != null ? sampleSize2 : string.Empty);
                        ws.Cells["I" + row].Value = (sampleSize3 != null ? sampleSize3 : string.Empty);
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":D" + row].Merge = true;
                        ws.Cells["E" + row + ":F" + row].Merge = true;
                        ws.Cells["G" + row + ":H" + row].Merge = true;
                        ws.Cells["I" + row + ":J" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);
                        xlsService.ExcelSetVerticalAlignCenter(ws, "C" + row + ":J" + row);
                        xlsService.ExcelSetHorizontalAlignCenter(ws, "C" + row + ":J" + row);

                        row++;
                        //Row 23
                        string str16A1 = string.Empty;
                        string str16B1 = string.Empty;
                        string str16C1 = string.Empty;
                        string sourceFile1 = string.Empty;

                        string str16A2 = string.Empty;
                        string str16B2 = string.Empty;
                        string str16C2 = string.Empty;
                        string sourceFile2 = string.Empty;

                        string str16A3 = string.Empty;
                        string str16B3 = string.Empty;
                        string str16C3 = string.Empty;
                        string sourceFile3 = string.Empty;

                        //FOR WT



                        str16A1 = roundSet.ListUserInputRound1.Where(x => x.StrQuestion.ToLower().Contains("source file")).Select(x => x.StrAnswer).FirstOrDefault();
                        //str16A1 = roundSet.ListUserInputRound1.Where(x => x.StrQuestion.ToLower().Contains("16.a") || x.StrQuestion.ToLower().Contains("16. a")).Select(x => x.StrAnswer).FirstOrDefault();
                        str16B1 = roundSet.ListUserInputRound1.Where(x => x.StrQuestion.ToLower().Contains("who provided the document")).Select(x => x.StrAnswer).FirstOrDefault();
                        //str16B1 = roundSet.ListUserInputRound1.Where(x => x.StrQuestion.ToLower().Contains("16.b") || x.StrQuestion.ToLower().Contains("16. b")).Select(x => x.StrAnswer).FirstOrDefault();
                        str16C1 = roundSet.ListUserInputRound1.Where(x => x.StrQuestion.ToLower().Contains("list down the documen")).Select(x => x.StrAnswer).FirstOrDefault();
                        //str16C1 = roundSet.ListUserInputRound1.Where(x => x.StrQuestion.ToLower().Contains("16.c") || x.StrQuestion.ToLower().Contains("16. c")).Select(x => x.StrAnswer).FirstOrDefault();
                        sourceFile1 = (str16A1 != null && str16A1 != string.Empty ? str16A1 : string.Empty) + Environment.NewLine +
                            (str16B1 != null && str16A1 != string.Empty ? $" provided by {str16B1}" : string.Empty) + Environment.NewLine +
                            (str16C1 != null ? str16C1 : string.Empty);

                        str16A2 = roundSet.ListUserInputRound2.Where(x => x.StrQuestion.ToLower().Contains("source file")).Select(x => x.StrAnswer).FirstOrDefault();
                        //str16A2 = roundSet.ListUserInputRound2.Where(x => x.StrQuestion.ToLower().Contains("16.a") || x.StrQuestion.ToLower().Contains("16. a")).Select(x => x.StrAnswer).FirstOrDefault();
                        str16B2 = roundSet.ListUserInputRound2.Where(x => x.StrQuestion.ToLower().Contains("who provided the document")).Select(x => x.StrAnswer).FirstOrDefault();
                        //str16B2 = roundSet.ListUserInputRound2.Where(x => x.StrQuestion.ToLower().Contains("16.b") || x.StrQuestion.ToLower().Contains("16. b")).Select(x => x.StrAnswer).FirstOrDefault();
                        str16C2 = roundSet.ListUserInputRound2.Where(x => x.StrQuestion.ToLower().Contains("list down the document")).Select(x => x.StrAnswer).FirstOrDefault();
                        //str16C2 = roundSet.ListUserInputRound2.Where(x => x.StrQuestion.ToLower().Contains("16.c") || x.StrQuestion.ToLower().Contains("16. c")).Select(x => x.StrAnswer).FirstOrDefault();
                        sourceFile2 = (str16A2 != null && str16A2 != string.Empty ? str16A2 : string.Empty) + Environment.NewLine +
                            (str16B2 != null && str16A2 != string.Empty ? $" provided by {str16B2}" : string.Empty) + Environment.NewLine +
                            (str16C2 != null ? str16C2 : string.Empty);

                        str16A3 = roundSet.ListUserInputRound3.Where(x => x.StrQuestion.ToLower().Contains("source file")).Select(x => x.StrAnswer).FirstOrDefault();
                        //str16A3 = roundSet.ListUserInputRound3.Where(x => x.StrQuestion.ToLower().Contains("16.a") || x.StrQuestion.ToLower().Contains("16. a")).Select(x => x.StrAnswer).FirstOrDefault();
                        str16B3 = roundSet.ListUserInputRound3.Where(x => x.StrQuestion.ToLower().Contains("who provided the document")).Select(x => x.StrAnswer).FirstOrDefault();
                        //str16B3 = roundSet.ListUserInputRound3.Where(x => x.StrQuestion.ToLower().Contains("16.b") || x.StrQuestion.ToLower().Contains("16. b")).Select(x => x.StrAnswer).FirstOrDefault();
                        str16C3 = roundSet.ListUserInputRound3.Where(x => x.StrQuestion.ToLower().Contains("list down the document")).Select(x => x.StrAnswer).FirstOrDefault();
                        //str16C3 = roundSet.ListUserInputRound3.Where(x => x.StrQuestion.ToLower().Contains("16.c") || x.StrQuestion.ToLower().Contains("16. c")).Select(x => x.StrAnswer).FirstOrDefault();
                        sourceFile3 = (str16A3 != null && str16A3 != string.Empty ? str16A3 : string.Empty) + Environment.NewLine +
                            (str16B3 != null && str16A3 != string.Empty ? $" provided by {str16B3}" : string.Empty) + Environment.NewLine +
                            (str16C3 != null ? str16C3 : string.Empty);


                        ws.Cells["A" + row].Value = "Source File (document name, hardcopy/softcopy & provided by):";
                        ws.Cells["E" + row].Value = (sourceFile1 != null ? txtFormat.FormatwithNewLine(sourceFile1, true) : string.Empty);
                        ws.Cells["G" + row].Value = (sourceFile2 != null ? txtFormat.FormatwithNewLine(sourceFile2, true) : string.Empty);
                        ws.Cells["I" + row].Value = (sourceFile3 != null ? txtFormat.FormatwithNewLine(sourceFile3, true) : string.Empty);
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":D" + row].Merge = true;
                        ws.Cells["E" + row + ":F" + row].Merge = true;
                        ws.Cells["G" + row + ":H" + row].Merge = true;
                        ws.Cells["I" + row + ":J" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);
                        xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":J" + row);
                        //set excel row height base on lenght on text
                        xlsService.XlsSetRow(ws, row, sourceFile1, 130, 150);
                        //ws.Row(row).Height = 320;

                        //int row = 30;
                        //Row 30;
                        row++;
                        var listRound1 = roundSet.ListRoundItem;

                        #region Test Performed By Table
                        //Row 24
                        row++;
                        string testPerformed1 = string.Empty;
                        string testPerformed2 = string.Empty;
                        string testPerformed3 = string.Empty;

                        testPerformed1 = roundSet.ListUserInputRound1.Where(x => x.StrQuestion.ToLower().Contains("performed the testing")).Select(x => x.StrAnswer).FirstOrDefault();
                        testPerformed2 = roundSet.ListUserInputRound2.Where(x => x.StrQuestion.ToLower().Contains("performed the testing")).Select(x => x.StrAnswer).FirstOrDefault();
                        testPerformed3 = roundSet.ListUserInputRound3.Where(x => x.StrQuestion.ToLower().Contains("performed the testing")).Select(x => x.StrAnswer).FirstOrDefault();

                        ws.Cells["A" + row].Value = "Test Performed By:";
                        //ws.Cells["C" + row].Value = (testPerformed1 != null ? txtFormat.ReplaceTagHtmlParagraph(testPerformed1, true) : string.Empty);
                        ws.Cells["E" + row].Value = (testPerformed1 != null ? txtFormat.FormatwithNewLine(testPerformed1, true) : string.Empty);
                        ws.Cells["G" + row].Value = (testPerformed2 != null ? txtFormat.FormatwithNewLine(testPerformed2, true) : string.Empty);
                        ws.Cells["I" + row].Value = (testPerformed3 != null ? txtFormat.FormatwithNewLine(testPerformed3, true) : string.Empty);
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":D" + row].Merge = true;
                        ws.Cells["E" + row + ":F" + row].Merge = true;
                        ws.Cells["G" + row + ":H" + row].Merge = true;
                        ws.Cells["I" + row + ":J" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);
                        xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetHorizontalAlignCenter(ws, "C" + row + ":J" + row);


                        //Row 25
                        row++;
                        string dtTestingPerformed1 = roundSet.ListUserInputRound1.Where(x => x.StrQuestion.ToLower().Contains("when was the testing completed")|| x.StrQuestion.ToLower().Contains("when is the testing was completed")).Select(x => x.StrAnswer).FirstOrDefault();
                        string dtTestingPerformed2 = roundSet.ListUserInputRound2.Where(x => x.StrQuestion.ToLower().Contains("when was the testing completed")|| x.StrQuestion.ToLower().Contains("when is the testing was completed")).Select(x => x.StrAnswer).FirstOrDefault();
                        string dtTestingPerformed3 = roundSet.ListUserInputRound3.Where(x => x.StrQuestion.ToLower().Contains("when was the testing completed")|| x.StrQuestion.ToLower().Contains("when is the testing was completed")).Select(x => x.StrAnswer).FirstOrDefault();

                        string parsedtTestingPerformed1 = string.Empty;
                        string parsedtTestingPerformed2 = string.Empty;
                        string parsedtTestingPerformed3 = string.Empty;

                        if (dtTestingPerformed1 != null && dtTestingPerformed1 != string.Empty)
                        {
                            //parsedtTestingPerformed = DateTime.Parse(dtTestingPerformed).ToString("MM/dd/yyyy");
                            DateTime tempdtTestingPerformed1;
                            //try to parse string to date time
                            if (DateTime.TryParse(dtTestingPerformed1, out tempdtTestingPerformed1))
                            {
                                parsedtTestingPerformed1 = tempdtTestingPerformed1.ToString("MM/dd/yyyy");
                            }
                            else
                            {
                                //if it failed to parse date time, then we will just remove time string
                                int index = dtTestingPerformed1.IndexOf(" ");
                                if (index > 0)
                                {
                                    parsedtTestingPerformed1 = dtTestingPerformed1.Substring(0, index);
                                }
                            }
                        }
                        if (dtTestingPerformed2 != null && dtTestingPerformed2 != string.Empty)
                        {
                            //parsedtTestingPerformed = DateTime.Parse(dtTestingPerformed).ToString("MM/dd/yyyy");
                            DateTime tempdtTestingPerformed2;
                            //try to parse string to date time
                            if (DateTime.TryParse(dtTestingPerformed2, out tempdtTestingPerformed2))
                            {
                                parsedtTestingPerformed2 = tempdtTestingPerformed2.ToString("MM/dd/yyyy");
                            }
                            else
                            {
                                //if it failed to parse date time, then we will just remove time string
                                int index = dtTestingPerformed2.IndexOf(" ");
                                if (index > 0)
                                {
                                    parsedtTestingPerformed2 = dtTestingPerformed2.Substring(0, index);
                                }
                            }
                        }
                        if (dtTestingPerformed3 != null && dtTestingPerformed3 != string.Empty)
                        {
                            //parsedtTestingPerformed = DateTime.Parse(dtTestingPerformed).ToString("MM/dd/yyyy");
                            DateTime tempdtTestingPerformed3;
                            //try to parse string to date time
                            if (DateTime.TryParse(dtTestingPerformed3, out tempdtTestingPerformed3))
                            {
                                parsedtTestingPerformed3 = tempdtTestingPerformed3.ToString("MM/dd/yyyy");
                            }
                            else
                            {
                                //if it failed to parse date time, then we will just remove time string
                                int index = dtTestingPerformed3.IndexOf(" ");
                                if (index > 0)
                                {
                                    parsedtTestingPerformed3 = dtTestingPerformed3.Substring(0, index);
                                }
                            }
                        }

                        ws.Cells["A" + row].Value = "Date Testing Performed:";
                        ws.Cells["C" + row + ":J" + row].Style.Numberformat.Format = "@";
                        ws.Cells["E" + row].Value = (parsedtTestingPerformed1 != null ? parsedtTestingPerformed1 : string.Empty);
                        ws.Cells["G" + row].Value = (parsedtTestingPerformed2 != null ? parsedtTestingPerformed2 : string.Empty);
                        ws.Cells["J" + row].Value = (parsedtTestingPerformed3 != null ? parsedtTestingPerformed3 : string.Empty);
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":D" + row].Merge = true;
                        ws.Cells["E" + row + ":F" + row].Merge = true;
                        ws.Cells["G" + row + ":H" + row].Merge = true;
                        ws.Cells["I" + row + ":J" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);
                        xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetHorizontalAlignCenter(ws, "C" + row + ":J" + row);


                        //Row 26
                        row++;
                        //#18 in questionnaire
                        //McGrath-19. What is the testing result?
                        //string testOfDesign1 = "";
                        string testOfDesign2 = string.Empty;
                        string testOfDesign3 = string.Empty;
                        string testOfDesign1 = string.Empty;

                        //

                        if (clientName == "McGrath")
                       
                        {

                             testOfDesign1 = roundSet.ListUserInputRound1.Where(x => x.StrQuestion.ToLower().Contains("what is the testing result")).Select(x => x.StrAnswer).FirstOrDefault();
                             testOfDesign2 = roundSet.ListUserInputRound2.Where(x => x.StrQuestion.ToLower().Contains("what is the testing result")).Select(x => x.StrAnswer).FirstOrDefault();
                             testOfDesign3 = roundSet.ListUserInputRound3.Where(x => x.StrQuestion.ToLower().Contains("what is the testing result")).Select(x => x.StrAnswer).FirstOrDefault();

                            ws.Cells["A" + row].Value = "Testing Result:";

                        }
                        else
                        {
                             testOfDesign1 = roundSet.ListUserInputRound1.Where(x => x.StrQuestion.ToLower().Contains("(tod)")).Select(x => x.StrAnswer).FirstOrDefault();
                             testOfDesign2 = roundSet.ListUserInputRound2.Where(x => x.StrQuestion.ToLower().Contains("(tod)")).Select(x => x.StrAnswer).FirstOrDefault();
                             testOfDesign3 = roundSet.ListUserInputRound3.Where(x => x.StrQuestion.ToLower().Contains("(tod)")).Select(x => x.StrAnswer).FirstOrDefault();
                            ws.Cells["A" + row].Value = "Test of Design (TOD):";
                        }
                        //

                        //ws.Cells["A" + row].Value = "Test of Design (TOD):";
                       
                        ws.Cells["E" + row].Value = (testOfDesign1 != null ? txtFormat.FormatwithNewLine(testOfDesign1, true) : string.Empty);
                        ws.Cells["G" + row].Value = (testOfDesign2 != null ? txtFormat.FormatwithNewLine(testOfDesign2, true) : string.Empty);
                        ws.Cells["J" + row].Value = (testOfDesign3 != null ? txtFormat.FormatwithNewLine(testOfDesign3, true) : string.Empty);
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":D" + row].Merge = true;
                        ws.Cells["E" + row + ":F" + row].Merge = true;
                        ws.Cells["G" + row + ":H" + row].Merge = true;
                        ws.Cells["I" + row + ":J" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);
                        xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetHorizontalAlignCenter(ws, "C" + row + ":J" + row);
                        xlsService.ExcelSetBackgroundColorLBluePaper(ws, "C" + row + ":J" + row);


                        //Row 27
                        //#19 in questionnaire
                        row++;

                        string testOperatingEffect1 = roundSet.ListUserInputRound1.Where(x => x.StrQuestion.ToLower().Contains("(toe)")).Select(x => x.StrAnswer).FirstOrDefault();
                        string testOperatingEffect2 = roundSet.ListUserInputRound2.Where(x => x.StrQuestion.ToLower().Contains("(toe)")).Select(x => x.StrAnswer).FirstOrDefault();
                        string testOperatingEffect3 = roundSet.ListUserInputRound3.Where(x => x.StrQuestion.ToLower().Contains("(toe)")).Select(x => x.StrAnswer).FirstOrDefault();


                        ws.Cells["A" + row].Value = "Test of Operating Effectiveness (TOE):";
                        ws.Cells["E" + row].Value = testOperatingEffect1 != null ? txtFormat.FormatwithNewLine(testOperatingEffect1, true) : string.Empty;
                        ws.Cells["G" + row].Value = testOperatingEffect2 != null ? txtFormat.FormatwithNewLine(testOperatingEffect2, true) : string.Empty;
                        ws.Cells["I" + row].Value = testOperatingEffect3 != null ? txtFormat.FormatwithNewLine(testOperatingEffect3, true) : string.Empty;
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":D" + row].Merge = true;
                        ws.Cells["E" + row + ":F" + row].Merge = true;
                        ws.Cells["G" + row + ":H" + row].Merge = true;
                        ws.Cells["I" + row + ":J" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row);
                        xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetHorizontalAlignCenter(ws, "C" + row + ":J" + row);
                        xlsService.ExcelSetBackgroundColorLBluePaper(ws, "C" + row + ":J" + row);

                        #endregion  


                        #region Who performed the review Table
                        //24.Who performed the review? Enter first and last name.
                        string reviewer1 = roundSet.ListUserInputRound1.Where(x => x.StrQuestion.ToLower().Contains("who performed the review")).Select(x => x.StrAnswer).FirstOrDefault();
                        string reviewer2 = roundSet.ListUserInputRound2.Where(x => x.StrQuestion.ToLower().Contains("who performed the review")).Select(x => x.StrAnswer).FirstOrDefault();
                        string reviewer3 = roundSet.ListUserInputRound3.Where(x => x.StrQuestion.ToLower().Contains("who performed the review")).Select(x => x.StrAnswer).FirstOrDefault();

                        ws.Cells["A" + row].Value = "Reviewer:";
                        ws.Cells["E" + row].Value = reviewer1 != null ? reviewer1 : string.Empty;
                        ws.Cells["G" + row].Value = reviewer2 != null ? reviewer2 : string.Empty;
                        ws.Cells["I" + row].Value = reviewer3 != null ? reviewer3 : string.Empty;
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":D" + row].Merge = true;
                        ws.Cells["E" + row + ":F" + row].Merge = true;
                        ws.Cells["G" + row + ":H" + row].Merge = true;
                        ws.Cells["I" + row + ":J" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetVerticalAlignCenter(ws, "C" + row + ":J" + row);
                        xlsService.ExcelSetHorizontalAlignCenter(ws, "C" + row + ":J" + row);
                        row++;

                        //25. When is the date the testing was reviewed? Enter date.
                        string reviewDate1 = roundSet.ListUserInputRound1.Where(x => x.StrQuestion.ToLower().Contains("when is the date the testing was reviewed")).Select(x => x.StrAnswer).FirstOrDefault();
                        string reviewDate2 = roundSet.ListUserInputRound2.Where(x => x.StrQuestion.ToLower().Contains("when is the date the testing was reviewed")).Select(x => x.StrAnswer).FirstOrDefault();
                        string reviewDate3 = roundSet.ListUserInputRound3.Where(x => x.StrQuestion.ToLower().Contains("when is the date the testing was reviewed")).Select(x => x.StrAnswer).FirstOrDefault();
                        string parsedtReviewDate1 = string.Empty;
                        string parsedtReviewDate2 = string.Empty;
                        string parsedtReviewDate3 = string.Empty;

                        if (reviewDate1 != null && reviewDate1 != string.Empty)
                        {
                            DateTime tempdtReviewDate1;
                            //try to parse string to date time
                            if (DateTime.TryParse(reviewDate1, out tempdtReviewDate1))
                            {
                                parsedtReviewDate1 = tempdtReviewDate1.ToString("MM/dd/yyyy");
                            }
                            else
                            {
                                //if it failed to parse date time, then we will just remove time string
                                int index = reviewDate1.IndexOf(" ");
                                if (index > 0)
                                {
                                    parsedtReviewDate1 = reviewDate1.Substring(0, index);
                                }
                            }
                        }
                        if (reviewDate2 != null && reviewDate2 != string.Empty)
                        {
                            DateTime tempdtReviewDate2;
                            //try to parse string to date time
                            if (DateTime.TryParse(reviewDate2, out tempdtReviewDate2))
                            {
                                parsedtReviewDate2 = tempdtReviewDate2.ToString("MM/dd/yyyy");
                            }
                            else
                            {
                                //if it failed to parse date time, then we will just remove time string
                                int index = reviewDate2.IndexOf(" ");
                                if (index > 0)
                                {
                                    parsedtReviewDate2 = reviewDate2.Substring(0, index);
                                }
                            }
                        }
                        if (reviewDate3 != null && reviewDate3 != string.Empty)
                        {
                            DateTime tempdtReviewDate3;
                            //try to parse string to date time
                            if (DateTime.TryParse(reviewDate1, out tempdtReviewDate3))
                            {
                                parsedtReviewDate3 = tempdtReviewDate3.ToString("MM/dd/yyyy");
                            }
                            else
                            {
                                //if it failed to parse date time, then we will just remove time string
                                int index = reviewDate3.IndexOf(" ");
                                if (index > 0)
                                {
                                    parsedtReviewDate3 = reviewDate3.Substring(0, index);
                                }
                            }
                        }

                        ws.Cells["A" + row].Value = "Review Date:";
                        ws.Cells["C" + row + ":J" + row].Style.Numberformat.Format = "@";
                        ws.Cells["E" + row].Value = parsedtReviewDate1 != null ? parsedtReviewDate1 : string.Empty;
                        ws.Cells["G" + row].Value = parsedtReviewDate2 != null ? parsedtReviewDate2 : string.Empty;
                        ws.Cells["I" + row].Value = parsedtReviewDate3 != null ? parsedtReviewDate3 : string.Empty;
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":D" + row].Merge = true;
                        ws.Cells["E" + row + ":F" + row].Merge = true;
                        ws.Cells["G" + row + ":H" + row].Merge = true;
                        ws.Cells["I" + row + ":J" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetVerticalAlignCenter(ws, "C" + row + ":J" + row);
                        xlsService.ExcelSetHorizontalAlignCenter(ws, "C" + row + ":J" + row);
                        row++;

                        //19. What is the testing result?
                        string testFindings1 = roundSet.ListUserInputRound1.Where(x => x.StrQuestion.ToLower().Contains("what is the testing status")).Select(x => x.StrAnswer).FirstOrDefault();
                        string testFindings2 = roundSet.ListUserInputRound2.Where(x => x.StrQuestion.ToLower().Contains("what is the testing status")).Select(x => x.StrAnswer).FirstOrDefault();
                        string testFindings3 = roundSet.ListUserInputRound3.Where(x => x.StrQuestion.ToLower().Contains("what is the testing status")).Select(x => x.StrAnswer).FirstOrDefault();

                        ws.Cells["A" + row].Value = "Test Findings Assessment:";
                        ws.Cells["E" + row].Value = testFindings1 != null ? txtFormat.ReplaceTagHtmlParagraph(testFindings1, true) : string.Empty;
                        ws.Cells["G" + row].Value = testFindings2 != null ? txtFormat.ReplaceTagHtmlParagraph(testFindings2, true) : string.Empty;
                        ws.Cells["I" + row].Value = testFindings3 != null ? txtFormat.ReplaceTagHtmlParagraph(testFindings3, true) : string.Empty;
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        ws.Cells["C" + row + ":D" + row].Merge = true;
                        ws.Cells["E" + row + ":F" + row].Merge = true;
                        ws.Cells["G" + row + ":H" + row].Merge = true;
                        ws.Cells["I" + row + ":J" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":J" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetVerticalAlignCenter(ws, "C" + row + ":J" + row);
                        xlsService.ExcelSetHorizontalAlignCenter(ws, "C" + row + ":J" + row);
                        xlsService.ExcelSetBackgroundColorLBluePaper(ws, "C" + row + ":J" + row);
                        row += 2;

                        #endregion


                        #region Sample Selection Table

                        if (listRound1 != null)
                        {

                            if (clientName == "ERI" && controlId == "HRP 2.1")
                            {
                                string title = roundSet.ListUserInputRound1.Where(x => x.StrQuestion.ToLower().Contains("(rt title)")).Select(x => x.StrQuestion.Replace("(RT Title)", "")).FirstOrDefault();
                                ws.Cells["A" + row].Value = $"{title}";
                                ws.Cells["A" + row + ":C" + row].Merge = true;
                                xlsService.ExcelSetBorder(ws, row, 1, row, 3);
                                xlsService.ExcelSetBackgroundColorGray(ws, row, 1, row, 3);
                                xlsService.ExcelSetArialSize12(ws, row, 1, row, 3);
                                xlsService.ExcelSetFontBold(ws, row, 1, row, 3);

                                string answer = roundSet.ListUserInputRound1.Where(x => x.StrQuestion.ToLower().Contains("(rt title)")).Select(x => x.StrAnswer).FirstOrDefault();
                                ws.Cells["D" + row].Value = $"{answer}";
                                ws.Cells["D" + row + ":F" + row].Merge = true;
                                xlsService.ExcelWrapText(ws, row, 4, row, 6);
                                xlsService.ExcelSetBorder(ws, row, 4, row, 6);
                                xlsService.ExcelSetBackgroundColorLightGray(ws, row, 4, row, 6);
                                xlsService.ExcelSetArialSize10(ws, row, 4, row, 6);
                                xlsService.ExcelSetFontBold(ws, row, 4, row, 6);
                                xlsService.ExcelSetVerticalAlignCenter(ws, row, 4, row, 6);
                                row++;
                            }

                            #region Round Sample

                            int columnRoundHeader = 1;

                            #region Round Header
                            var roundHeader = roundSet.ListUserQuestion.Where(x => x.QuestionString.Contains("(RT)"));
                            if (roundHeader != null)
                            {
                                ws.Cells[row, columnRoundHeader].Value = "Sample #";
                                columnRoundHeader++;
                                int countHeaderNotes = 1;
                                foreach (var item in roundHeader)
                                {
                                    //filter policy changes for ELC 1.1
                                    if (!item.QuestionString.Contains("Policy Changes? (R1)") &&
                                        !item.QuestionString.Contains("Policy Changes? (R2)") &&
                                        !item.QuestionString.Contains("Policy Changes? (R3)"))
                                    {
                                        string strHeader = item.QuestionString.Replace("(RT)", "");
                                        string strHeader18 =strHeader.Replace("18.", "");
                                        
                                        string tempNote = string.Empty;

                                        if (roundSet.ListHeaderNote != null && roundSet.ListHeaderNote.Count > 0)
                                        {
                                            var checkHeaderNote = roundSet.ListHeaderNote.Where(x => x.Position.Equals(countHeaderNotes)).FirstOrDefault();
                                            if (checkHeaderNote != null && checkHeaderNote.HeaderNoteText != string.Empty)
                                            {
                                                //strHeader = strHeader + (checkHeaderNote.HeaderNoteText != string.Empty ? $" [{checkHeaderNote.HeaderNoteText}] " : string.Empty);
                                                tempNote = (checkHeaderNote.HeaderNoteText != string.Empty ? checkHeaderNote.HeaderNoteText : string.Empty);
                                            }
                                        }

                                        //ws.Cells[row, columnRoundHeader].Value = strHeader;

                                       // xlsService.TestingAttributeFormat(ws, row, columnRoundHeader, strHeader, tempNote);
                                        xlsService.TestingAttributeFormat(ws, row, columnRoundHeader, strHeader18, tempNote);
                                        columnRoundHeader++;
                                        countHeaderNotes++;
                                    }
                                }

                                xlsService.ExcelWrapText(ws, row, 1, row, columnRoundHeader - 1);
                                xlsService.ExcelSetBorder(ws, row, 1, row, columnRoundHeader - 1);
                                xlsService.ExcelSetBackgroundColorGray(ws, row, 1, row, columnRoundHeader - 1);
                                xlsService.ExcelSetArialSize10(ws, row, 1, row, columnRoundHeader - 1);
                                xlsService.ExcelSetFontBold(ws, row, 1, row, columnRoundHeader - 1);
                                xlsService.ExcelSetVerticalAlignCenter(ws, row, 1, row, columnRoundHeader - 1);
                                xlsService.ExcelSetHorizontalAlignCenter(ws, row, 1, row, columnRoundHeader - 1);
                                row++;
                            }

                            #endregion


                            #region Round 1
                            //Round 1
                            ws.Cells[row, 1].Value = "Round 1";
                            ws.Cells[row, 1, row, columnRoundHeader - 1].Merge = true;
                            xlsService.ExcelWrapText(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetBorder(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetBackgroundColorLightGray(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetArialSize10(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetFontBold(ws, row, 1, row, columnRoundHeader - 1);
                            row++;


                            if (roundSet.ListRoundItem != null && roundSet.ListRoundItem.Count() > 0)
                            {
                                foreach (var item in roundSet.ListRoundItem)
                                {
                                    int countRoundData1 = 1;
                                    if (item.RoundName == "Round 1")
                                    {

                                        //if (questionnaireInput.ListPolicyNote != null && questionnaireInput.ListPolicyNote.Count > 0)
                                        //{
                                        //    string policyQuestion1 = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("policy changes? (r1)")).Select(x => x.StrAnswer).FirstOrDefault();
                                        //    string policyNote1 = questionnaireInput.ListPolicyNote.Where(x => x.Position.Equals(0)).Select(x => x.NoteText).FirstOrDefault();
                                        //    string policy1 = (policyQuestion1 != null ? policyQuestion1 : string.Empty) + (policyNote1 != null && policyNote1 != string.Empty ? $" [{policyNote1}]" : string.Empty);
                                        //    ws.Cells[row, 1].Value = "Policy Changes?";
                                        //    ws.Cells[row, 2].Value = policy1 != null ? policy1 : string.Empty;
                                        //    ws.Cells[row, 2, row, columnRoundHeader - 1].Merge = true;
                                        //    ExcelWrapText2(ws, row, 1, row, columnRoundHeader - 1);
                                        //    ExcelSetBorder2(ws, row, 1, row, columnRoundHeader - 1);
                                        //    ExcelSetHorizontalAlignCenter2(ws, row, 2, row, columnRoundHeader - 1);
                                        //    ExcelSetVerticalAlignCenter2(ws, row, 2, row, columnRoundHeader - 1);
                                        //    ExcelSetArialSize102(ws, row, 1, row, columnRoundHeader - 1);
                                        //    ExcelSetBackgroundColorLightGray2(ws, row, 1, row, columnRoundHeader - 1);
                                        //    row++;
                                        //}

                                        #region Populate Data for Round 1

                                        ws.Cells[row, countRoundData1].Value = $"{(item.A2Q2Samples != null && item.A2Q2Samples != string.Empty ? item.A2Q2Samples.ToString() : string.Empty)}";
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer1, item.Note1);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer2, item.Note2);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer3, item.Note3);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer4, item.Note4);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer5, item.Note5);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer6, item.Note6);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer7, item.Note7);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer8, item.Note8);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer9, item.Note9);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer10, item.Note10);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer11, item.Note11);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer12, item.Note12);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer13, item.Note13);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer14, item.Note14);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer15, item.Note15);
                                        countRoundData1++;
                                        #endregion

                                        xlsService.ExcelWrapText(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetBorder(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetArialSize10(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetVerticalAlignCenter(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetHorizontalAlignCenter(ws, row, 1, row, columnRoundHeader - 1);

                                        ws.Row(row).Height = 30;
                                        row++;
                                    }
                                }
                            }

                            #endregion

                            #region Round 2
                            //Round 2
                            ws.Cells[row, 1].Value = "Round 2";
                            ws.Cells[row, 1, row, columnRoundHeader - 1].Merge = true;
                            xlsService.ExcelWrapText(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetBorder(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetBackgroundColorLightGray(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetArialSize10(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetFontBold(ws, row, 1, row, columnRoundHeader - 1);
                            row++;

                            if (roundSet.ListRoundItem != null && roundSet.ListRoundItem.Count() > 0)
                            {
                                foreach (var item in roundSet.ListRoundItem)
                                {
                                    int countRoundData1 = 1;
                                    if (item.RoundName == "Round 2")
                                    {

                                        //if (questionnaireInput.ListPolicyNote != null && questionnaireInput.ListPolicyNote.Count > 0)
                                        //{
                                        //    string policyQuestion1 = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("policy changes? (r2)")).Select(x => x.StrAnswer).FirstOrDefault();
                                        //    string policyNote1 = questionnaireInput.ListPolicyNote.Where(x => x.Position.Equals(1)).Select(x => x.NoteText).FirstOrDefault();
                                        //    string policy1 = (policyQuestion1 != null ? policyQuestion1 : string.Empty) + (policyNote1 != null && policyNote1 != string.Empty ? $" [{policyNote1}]" : string.Empty);
                                        //    ws.Cells[row, 1].Value = "Policy Changes?";
                                        //    ws.Cells[row, 2].Value = policy1 != null ? policy1 : string.Empty;
                                        //    ws.Cells[row, 2, row, columnRoundHeader - 1].Merge = true;
                                        //    ExcelWrapText2(ws, row, 1, row, columnRoundHeader - 1);
                                        //    ExcelSetBorder2(ws, row, 1, row, columnRoundHeader - 1);
                                        //    ExcelSetHorizontalAlignCenter2(ws, row, 2, row, columnRoundHeader - 1);
                                        //    ExcelSetVerticalAlignCenter2(ws, row, 2, row, columnRoundHeader - 1);
                                        //    ExcelSetArialSize102(ws, row, 1, row, columnRoundHeader - 1);
                                        //    ExcelSetBackgroundColorLightGray2(ws, row, 1, row, columnRoundHeader - 1);
                                        //    row++;
                                        //}


                                        #region Populate Data for Round 2
                                        //if (item.A2Q2Samples != null && item.A2Q2Samples != string.Empty)
                                        ws.Cells[row, countRoundData1].Value = $"{(item.A2Q2Samples != null && item.A2Q2Samples != string.Empty ? item.A2Q2Samples.ToString() : string.Empty)}";
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer1, item.Note1);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer2, item.Note2);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer3, item.Note3);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer4, item.Note4);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer5, item.Note5);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer6, item.Note6);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer7, item.Note7);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer8, item.Note8);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer9, item.Note9);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer10, item.Note10);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer11, item.Note11);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer12, item.Note12);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer13, item.Note13);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer14, item.Note14);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer15, item.Note15);
                                        countRoundData1++;


                                        #endregion

                                        xlsService.ExcelWrapText(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetBorder(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetArialSize10(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetVerticalAlignCenter(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetHorizontalAlignCenter(ws, row, 1, row, columnRoundHeader - 1);

                                        ws.Row(row).Height = 30;
                                        row++;
                                    }
                                }
                            }

                            #endregion

                            #region Round 3
                            //Round 3
                            ws.Cells[row, 1].Value = "Round 3";
                            ws.Cells[row, 1, row, columnRoundHeader - 1].Merge = true;
                            xlsService.ExcelWrapText(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetBorder(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetBackgroundColorLightGray(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetArialSize10(ws, row, 1, row, columnRoundHeader - 1);
                            xlsService.ExcelSetFontBold(ws, row, 1, row, columnRoundHeader - 1);
                            row++;

                            if (roundSet.ListRoundItem != null && roundSet.ListRoundItem.Count() > 0)
                            {
                                foreach (var item in roundSet.ListRoundItem)
                                {
                                    int countRoundData1 = 1;
                                    if (item.RoundName == "Round 3")
                                    {

                                        //if (questionnaireInput.ListPolicyNote != null && questionnaireInput.ListPolicyNote.Count > 0)
                                        //{
                                        //    string policyQuestion1 = questionnaireInput.ListUserInputItem.Where(x => x.StrQuestion.ToLower().Contains("policy changes? (r3)")).Select(x => x.StrAnswer).FirstOrDefault();
                                        //    string policyNote1 = questionnaireInput.ListPolicyNote.Where(x => x.Position.Equals(2)).Select(x => x.NoteText).FirstOrDefault();
                                        //    string policy1 = (policyQuestion1 != null ? policyQuestion1 : string.Empty) + (policyNote1 != null && policyNote1 != string.Empty ? $" [{policyNote1}]" : string.Empty);
                                        //    ws.Cells[row, 1].Value = "Policy Changes?";
                                        //    ws.Cells[row, 2].Value = policy1 != null ? policy1 : string.Empty;
                                        //    ws.Cells[row, 2, row, columnRoundHeader - 1].Merge = true;
                                        //    ExcelWrapText2(ws, row, 1, row, columnRoundHeader - 1);
                                        //    ExcelSetBorder2(ws, row, 1, row, columnRoundHeader - 1);
                                        //    ExcelSetHorizontalAlignCenter2(ws, row, 2, row, columnRoundHeader - 1);
                                        //    ExcelSetVerticalAlignCenter2(ws, row, 2, row, columnRoundHeader - 1);
                                        //    ExcelSetArialSize102(ws, row, 1, row, columnRoundHeader - 1);
                                        //    ExcelSetBackgroundColorLightGray2(ws, row, 1, row, columnRoundHeader - 1);
                                        //    row++;
                                        //}


                                        #region Populate Data for Round 3
                                        //if (item.A2Q2Samples != null && item.A2Q2Samples != string.Empty)
                                        ws.Cells[row, countRoundData1].Value = $"{(item.A2Q2Samples != null && item.A2Q2Samples != string.Empty ? item.A2Q2Samples.ToString() : string.Empty)}";
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer1, item.Note1);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer2, item.Note2);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer3, item.Note3);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer4, item.Note4);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer5, item.Note5);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer6, item.Note6);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer7, item.Note7);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer8, item.Note8);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer9, item.Note9);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer10, item.Note10);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer11, item.Note11);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer12, item.Note12);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer13, item.Note13);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer14, item.Note14);
                                        countRoundData1++;

                                        xlsService.TestingAttributeFormat(ws, row, countRoundData1, item.Answer15, item.Note15);
                                        countRoundData1++;

                                        #endregion

                                        xlsService.ExcelWrapText(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetBorder(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetArialSize10(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetVerticalAlignCenter(ws, row, 1, row, columnRoundHeader - 1);
                                        xlsService.ExcelSetHorizontalAlignCenter(ws, row, 1, row, columnRoundHeader - 1);

                                        ws.Row(row).Height = 30;
                                        row++;
                                    }
                                }
                            }

                            #endregion

                            #endregion

                        }

                        #endregion


                        #region Legends

                        var uniqueNotes = roundSet.ListUniqueNotes;

                        row += 2;
                        ws.Cells["A" + row].Value = "Legend";
                        ws.Cells["A" + row + ":B" + row].Merge = true;
                        xlsService.ExcelWrapText(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetBorder(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetArialSize10(ws, "A" + row + ":B" + row);
                        xlsService.ExcelSetFontBold(ws, "A" + row + ":B" + row);
                        row++;
                        if (uniqueNotes != null && uniqueNotes.Count > 0)
                        {
                            var genNotes = uniqueNotes.Where(x => x.Notes.ToLower().Equals("general note")).FirstOrDefault();
                            if (genNotes != null)
                            {

                                ws.Cells["A" + row].Value = genNotes.Notes != null ? genNotes.Notes : string.Empty;
                                ws.Cells["C" + row].Value = genNotes.Description != null ? txtFormat.ReplaceTagHtmlParagraph(genNotes.Description, true) : string.Empty;
                                ws.Cells["A" + row + ":B" + row].Merge = true;
                                ws.Cells["C" + row + ":H" + row].Merge = true;
                                xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                                xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                                xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                                xlsService.ExcelSetFontBold(ws, "A" + row + ":B" + row);
                                xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":H" + row);
                                xlsService.ExcelSetHorizontalAlignCenter(ws, "A" + row + ":B" + row);
                                xlsService.ExcelSetFontColorRed(ws, row, 1, row, 1);
                                ws.Row(row).Height = 30;
                                row++;

                            }


                            foreach (var item in uniqueNotes.OrderBy(x => x.Notes))
                            {
                                if (!item.Notes.ToLower().Equals("general note"))
                                {
                                    ws.Cells["A" + row].Value = item.Notes != null && item.Notes != string.Empty ? "{" + item.Notes + "}" : string.Empty;
                                    ws.Cells["C" + row].Value = item.Description != null && item.Description != string.Empty ? txtFormat.ReplaceTagHtmlParagraph(item.Description, true) : string.Empty;
                                    ws.Cells["A" + row + ":B" + row].Merge = true;
                                    ws.Cells["C" + row + ":H" + row].Merge = true;
                                    xlsService.ExcelWrapText(ws, "A" + row + ":H" + row);
                                    xlsService.ExcelSetBorder(ws, "A" + row + ":H" + row);
                                    xlsService.ExcelSetArialSize10(ws, "A" + row + ":H" + row);
                                    xlsService.ExcelSetFontBold(ws, "A" + row + ":B" + row);
                                    xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":H" + row);
                                    xlsService.ExcelSetHorizontalAlignCenter(ws, "A" + row + ":B" + row);
                                    xlsService.ExcelSetFontColorRed(ws, row, 1, row, 1);
                                    //set excel row height base on lenght on text
                                    xlsService.XlsSetRow(ws, row, item.Description, 380, 30);

                                    row++;
                                }

                            }



                        }
                        row++;
                        #endregion


                        #region Round Header
                        //row++;
                        //ws.Cells["C" + row].Value = "Round 1";
                        //ws.Cells["C" + row + ":D" + row].Merge = true;
                        //xlsService.ExcelWrapText(ws, "C" + row + ":D" + row);
                        //xlsService.ExcelSetBorder(ws, "C" + row + ":D" + row);

                        //ws.Cells["E" + row].Value = "Round 2";
                        //ws.Cells["E" + row + ":F" + row].Merge = true;
                        //xlsService.ExcelWrapText(ws, "E" + row + ":F" + row);
                        //xlsService.ExcelSetBorder(ws, "E" + row + ":F" + row);

                        //ws.Cells["G" + row].Value = "Round 3";
                        //ws.Cells["G" + row + ":H" + row].Merge = true;
                        //xlsService.ExcelWrapText(ws, "G" + row + ":H" + row);
                        //xlsService.ExcelSetBorder(ws, "G" + row + ":H" + row);

                        //xlsService.ExcelSetBackgroundColorGray(ws, "C" + row + ":H" + row);
                        //xlsService.ExcelSetArialSize10(ws, "C" + row + ":H" + row);
                        //xlsService.ExcelSetFontBold(ws, "C" + row + ":H" + row);
                        //xlsService.ExcelSetVerticalAlignCenter(ws, "C" + row + ":H" + row);
                        //xlsService.ExcelSetHorizontalAlignCenter(ws, "C" + row + ":H" + row);
                        //row++;
                        #endregion



                        List<IUCSystemGenAnswer> listIUCSystemGen = new List<IUCSystemGenAnswer>();
                        List<IUCNonSystemGenAnswer> listIUCNonSystemGen = new List<IUCNonSystemGenAnswer>();
                        if (roundSet.ListIUCSystemGen1 != null)
                            listIUCSystemGen.AddRange(roundSet.ListIUCSystemGen1);
                        if (roundSet.ListIUCSystemGen2 != null)
                            listIUCSystemGen.AddRange(roundSet.ListIUCSystemGen2);
                        if (roundSet.ListIUCSystemGen3 != null)
                            listIUCSystemGen.AddRange(roundSet.ListIUCSystemGen3);

                        if (roundSet.ListIUCNonSystemGen1 != null)
                            listIUCNonSystemGen.AddRange(roundSet.ListIUCNonSystemGen1);
                        if (roundSet.ListIUCNonSystemGen2 != null)
                            listIUCNonSystemGen.AddRange(roundSet.ListIUCNonSystemGen2);
                        if (roundSet.ListIUCNonSystemGen3 != null)
                            listIUCNonSystemGen.AddRange(roundSet.ListIUCNonSystemGen3);

                        if ((listIUCSystemGen != null && listIUCSystemGen.Count > 0) ||
                           listIUCNonSystemGen != null && listIUCNonSystemGen.Count > 0)
                        {
                            #region IUC Header
                            row += 2;
                            ws.Cells["A" + row].Value = "IUC";
                            ws.Cells["A" + row + ":F" + row].Merge = true;
                            xlsService.ExcelWrapText(ws, "A" + row + ":F" + row);
                            xlsService.ExcelSetBorder(ws, "A" + row + ":F" + row);
                            xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":F" + row);
                            xlsService.ExcelSetFontBold(ws, "A" + row + ":F" + row);
                            xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":F" + row);
                            xlsService.ExcelSetHorizontalAlignCenter(ws, "A" + row + ":F" + row);
                            row++;

                            ws.Cells["A" + row].Value = "Name of Report";
                            ws.Cells["B" + row].Value = "IUC Type";
                            ws.Cells["C" + row].Value = "Source Control";
                            ws.Cells["D" + row].Value = "IUC Validation";
                            ws.Cells["D" + row + ":F" + row].Merge = true;
                            xlsService.ExcelSetBorder(ws, "A" + row + ":F" + row);
                            xlsService.ExcelSetBackgroundColorGray(ws, "A" + row + ":F" + row);
                            xlsService.ExcelSetFontBold(ws, "A" + row + ":F" + row);
                            xlsService.ExcelSetVerticalAlignCenter(ws, "A" + row + ":F" + row);
                            xlsService.ExcelSetHorizontalAlignCenter(ws, "A" + row + ":F" + row);
                            row++;

                            #endregion
                        }


                        #region IUC System Generated
                        if (listIUCSystemGen != null && listIUCSystemGen.Count > 0)
                        {
                            foreach (var item in listIUCSystemGen)
                            {
                                string reportName = item.ListQuestionAnswer.Where(x =>
                                    x.Question.ToLower().Contains("what is the report name")).Select(x => x.Answer).FirstOrDefault();

                                string sourceControl = item.ListQuestionAnswer.Where(x =>
                                    x.Question.ToLower().Contains("which system did this report come from, if any? if manually compiled, put spreadsheet")).Select(x => x.Answer).FirstOrDefault();

                                //string general = item.ListQuestionAnswer.Where(x => 
                                //    x.Question.ToLower().Contains("how is the iuc generated")).Select(x => x.Answer).FirstOrDefault();

                                string general = item.ListQuestionAnswer.Where(x =>
                                    x.Question.ToLower().Contains("how is the information used in control (iuc) or what documents are used by the control owner to do their task")).Select(x => x.Answer).FirstOrDefault();

                                //string reportParam = item.ListQuestionAnswer.Where(x => 
                                //    x.Question.ToLower().Contains("what is/are the evidence(s) obtained to address completeness and accuracy of user-entered parameters")).Select(x => x.Answer).FirstOrDefault();

                                string reportParam = item.ListQuestionAnswer.Where(x =>
                                    x.Question.ToLower().Contains("what evidence can you get to show the report parameters used such as report date or data filter? did control owners check if the report parameters are correct")).Select(x => x.Answer).FirstOrDefault();


                                //string sourceData1 = item.ListQuestionAnswer.Where(x => 
                                //    x.Question.ToLower().Contains("how do/does the control(s) address completeness and accuracy")).Select(x => x.Answer).FirstOrDefault();

                                string sourceData1 = item.ListQuestionAnswer.Where(x =>
                                    x.Question.ToLower().Contains("did the control owner check that the report they used is complete and accurate")).Select(x => x.Answer).FirstOrDefault();

                                //string sourceData2 = item.ListQuestionAnswer.Where(x => 
                                //    x.Question.ToLower().Contains("is/are the control(s) designed and operating effectively (completeness)")).Select(x => x.Answer).FirstOrDefault();

                                string sourceData2 = item.ListQuestionAnswer.Where(x =>
                                    x.Question.ToLower().Contains("is/are the control(s) designed and operating effectively for data completeness")).Select(x => x.Answer).FirstOrDefault();


                                //string sourceData3 = item.ListQuestionAnswer.Where(x => 
                                //    x.Question.ToLower().Contains("is/are the control(s) designed and operating effectively (accuracy)")).Select(x => x.Answer).FirstOrDefault();

                                string sourceData3 = item.ListQuestionAnswer.Where(x =>
                                    x.Question.ToLower().Contains("is/are the control(s) designed and operating effectively for data accuracy")).Select(x => x.Answer).FirstOrDefault();


                                //string reportLogic1 = item.ListQuestionAnswer.Where(x => 
                                //    x.Question.ToLower().Contains("how does/do the control(s) address the report logic")).Select(x => x.Answer).FirstOrDefault();

                                string reportLogic1 = item.ListQuestionAnswer.Where(x =>
                                    x.Question.ToLower().Contains("how does/do the control(s) address the report logic or formulas used")).Select(x => x.Answer).FirstOrDefault();


                                //string reportLogic2 = item.ListQuestionAnswer.Where(x => 
                                //    x.Question.ToLower().Contains("is/are the control(s) designed and operating effectively")).Select(x => x.Answer).FirstOrDefault();

                                string reportLogic2 = item.ListQuestionAnswer.Where(x =>
                                    x.Question.ToLower().Contains("is/are the control(s) designed and operating effectively for report logic or formulas used")).Select(x => x.Answer).FirstOrDefault();


                                int startRow = row;

                                ws.Cells["A" + row].Value = reportName != null ? txtFormat.ReplaceTagHtmlParagraph(reportName, true) : string.Empty;
                                ws.Cells["B" + row].Value = "System Generated";
                                ws.Cells["C" + row].Value = sourceControl != null ? txtFormat.ReplaceTagHtmlParagraph(sourceControl, true) : string.Empty;
                                ws.Cells["D" + row].Value = "General";
                                ws.Cells["E" + row].Value = "Description of How the IUC Is Generated";
                                ws.Cells["F" + row].Value = general != null ? txtFormat.ReplaceTagHtmlParagraph(general, true) : string.Empty;
                                xlsService.ExcelWrapText(ws, row, 1, row, 6);
                                xlsService.ExcelSetBorder(ws, row, 1, row, 6);
                                xlsService.ExcelSetArialSize10(ws, row, 1, row, 6);
                                xlsService.ExcelSetVerticalAlignTop(ws, row, 1, row, 6);
                                xlsService.ExcelSetHorizontalAlignLeft(ws, row, 1, row, 6);
                                //ws.Row(row).Height = 70;
                                xlsService.XlsSetRow(ws, row, general, 120, 70);
                                row++;

                                ws.Cells["D" + row].Value = "Report Parameters";
                                ws.Cells["E" + row].Value = "Evidence Obtained to Address Completeness and Accuracy of User-Entered Parameters";
                                ws.Cells["F" + row].Value = reportParam != null ? txtFormat.ReplaceTagHtmlParagraph(reportParam, true) : string.Empty;
                                xlsService.ExcelWrapText(ws, row, 4, row, 6);
                                xlsService.ExcelSetBorder(ws, row, 4, row, 6);
                                xlsService.ExcelSetArialSize10(ws, row, 4, row, 6);
                                xlsService.ExcelSetVerticalAlignTop(ws, row, 4, row, 6);
                                xlsService.ExcelSetHorizontalAlignLeft(ws, row, 4, row, 6);
                                //ws.Row(row).Height = 70;
                                xlsService.XlsSetRow(ws, row, reportParam, 120, 70);
                                row++;

                                #region Source Data

                                int starRowSourceData = row;
                                ws.Cells["D" + row].Value = "Source Data";
                                ws.Cells["E" + row].Value = "Control(s) Addressing Completeness and Accuracy (describe how)";
                                ws.Cells["F" + row].Value = sourceData1 != null ? txtFormat.ReplaceTagHtmlParagraph(sourceData1, true) : string.Empty;
                                xlsService.ExcelWrapText(ws, row, 5, row, 6);
                                xlsService.ExcelSetBorder(ws, row, 5, row, 6);
                                xlsService.ExcelSetArialSize10(ws, row, 5, row, 6);
                                xlsService.ExcelSetVerticalAlignTop(ws, row, 5, row, 6);
                                xlsService.ExcelSetHorizontalAlignLeft(ws, row, 5, row, 6);
                                //ws.Row(row).Height = 70;
                                xlsService.XlsSetRow(ws, row, sourceData1, 120, 70);
                                row++;

                                ws.Cells["E" + row].Value = "Is Control(s) Designed and Operating Effectively (completeness)?";
                                ws.Cells["F" + row].Value = sourceData2 != null ? txtFormat.ReplaceTagHtmlParagraph(sourceData2, true) : string.Empty;
                                xlsService.ExcelWrapText(ws, row, 5, row, 6);
                                xlsService.ExcelSetBorder(ws, row, 5, row, 6);
                                xlsService.ExcelSetArialSize10(ws, row, 5, row, 6);
                                xlsService.ExcelSetVerticalAlignTop(ws, row, 5, row, 6);
                                xlsService.ExcelSetHorizontalAlignLeft(ws, row, 5, row, 6);
                                //ws.Row(row).Height = 70;
                                xlsService.XlsSetRow(ws, row, sourceData2, 120, 70);
                                row++;

                                ws.Cells["E" + row].Value = "Is Control(s) Designed and Operating Effectively (accuracy)?";
                                ws.Cells["F" + row].Value = sourceData3 != null ? txtFormat.ReplaceTagHtmlParagraph(sourceData3, true) : string.Empty;
                                xlsService.ExcelWrapText(ws, row, 5, row, 6);
                                xlsService.ExcelSetBorder(ws, row, 5, row, 6);
                                xlsService.ExcelSetArialSize10(ws, row, 5, row, 6);
                                xlsService.ExcelSetVerticalAlignTop(ws, row, 5, row, 6);
                                xlsService.ExcelSetHorizontalAlignLeft(ws, row, 5, row, 6);
                                //ws.Row(row).Height = 70;
                                xlsService.XlsSetRow(ws, row, sourceData3, 120, 70);

                                int endRowSourceData = row;
                                ws.Cells[starRowSourceData, 4, endRowSourceData, 4].Merge = true;
                                xlsService.ExcelSetBorder(ws, starRowSourceData, 4, endRowSourceData, 4);
                                xlsService.ExcelSetVerticalAlignTop(ws, starRowSourceData, 4, starRowSourceData, 4);
                                xlsService.ExcelSetHorizontalAlignLeft(ws, starRowSourceData, 4, starRowSourceData, 4);

                                row++;

                                #endregion

                                #region Report Logic

                                int startRowReportLogic = row;
                                ws.Cells["D" + row].Value = "Report Logic";
                                ws.Cells["E" + row].Value = "Control(s) Addressing the Report Logic";
                                ws.Cells["F" + row].Value = reportLogic1 != null ? txtFormat.ReplaceTagHtmlParagraph(reportLogic1, true) : string.Empty;
                                xlsService.ExcelWrapText(ws, row, 5, row, 6);
                                xlsService.ExcelSetBorder(ws, row, 5, row, 6);
                                xlsService.ExcelSetArialSize10(ws, row, 5, row, 6);
                                xlsService.ExcelSetVerticalAlignTop(ws, row, 5, row, 6);
                                xlsService.ExcelSetHorizontalAlignLeft(ws, row, 5, row, 6);
                                //ws.Row(row).Height = 70;
                                xlsService.XlsSetRow(ws, row, reportLogic1, 120, 70);
                                row++;

                                ws.Cells["E" + row].Value = "Is Control(s) Designed and Operating Effectively?";
                                ws.Cells["F" + row].Value = reportLogic2 != null ? txtFormat.ReplaceTagHtmlParagraph(reportLogic2, true) : string.Empty;
                                xlsService.ExcelWrapText(ws, row, 5, row, 6);
                                xlsService.ExcelSetBorder(ws, row, 5, row, 6);
                                xlsService.ExcelSetArialSize10(ws, row, 5, row, 6);
                                xlsService.ExcelSetVerticalAlignTop(ws, row, 5, row, 6);
                                xlsService.ExcelSetHorizontalAlignLeft(ws, row, 5, row, 6);
                                //ws.Row(row).Height = 70;
                                xlsService.XlsSetRow(ws, row, reportLogic2, 120, 70);

                                int endRowReportLogic = row;
                                ws.Cells[startRowReportLogic, 4, endRowReportLogic, 4].Merge = true;
                                xlsService.ExcelSetBorder(ws, startRowReportLogic, 4, endRowReportLogic, 4);
                                xlsService.ExcelSetVerticalAlignTop(ws, startRowReportLogic, 4, startRowReportLogic, 4);
                                xlsService.ExcelSetHorizontalAlignLeft(ws, startRowReportLogic, 4, startRowReportLogic, 4);

                                #endregion

                                int endRow = row;

                                ws.Cells[startRow, 1, endRow, 1].Merge = true;
                                ws.Cells[startRow, 2, endRow, 2].Merge = true;
                                ws.Cells[startRow, 3, endRow, 3].Merge = true;
                                xlsService.ExcelSetBorder(ws, startRow, 1, endRow, 1);
                                xlsService.ExcelSetBorder(ws, startRow, 2, endRow, 2);
                                xlsService.ExcelSetBorder(ws, startRow, 3, endRow, 3);

                                row++;
                            }

                        }
                        #endregion

                        #region IUC Non System Generated
                        if (listIUCNonSystemGen != null && listIUCNonSystemGen.Count > 0)
                        {
                            foreach (var item in listIUCNonSystemGen)
                            {
                                string reportName = item.ListQuestionAnswer.Where(x => x.Question.ToLower().Contains("what is the report name")).Select(x => x.Answer).FirstOrDefault();
                                //string sourceControl = item.ListQuestionAnswer.Where(x => x.Question.ToLower().Contains("what is the source control")).Select(x => x.Answer).FirstOrDefault();
                                string sourceControl = item.ListQuestionAnswer.Where(x => x.Question.ToLower().Contains("does this control rely on another control for this control to work? what is the source control")).Select(x => x.Answer).FirstOrDefault();
                                string general = item.ListQuestionAnswer.Where(x => x.Question.ToLower().Contains("how do/does the control(s) address the completeness and accuracy of the non-system-generated report")).Select(x => x.Answer).FirstOrDefault();

                                ws.Cells["A" + row].Value = reportName != null ? reportName : string.Empty;
                                ws.Cells["B" + row].Value = "Non System Generated";
                                ws.Cells["C" + row].Value = sourceControl != null ? txtFormat.ReplaceTagHtmlParagraph(sourceControl, true) : string.Empty;
                                ws.Cells["D" + row].Value = "General";
                                ws.Cells["E" + row].Value = "Describe how the controls that addresses the completeness and accuracy of the Non-System-Generated Report";
                                ws.Cells["F" + row].Value = general != null ? txtFormat.ReplaceTagHtmlParagraph(general, true) : string.Empty;
                                xlsService.ExcelWrapText(ws, row, 1, row, 6);
                                xlsService.ExcelSetBorder(ws, row, 1, row, 6);
                                xlsService.ExcelSetArialSize10(ws, row, 1, row, 6);
                                xlsService.ExcelSetVerticalAlignTop(ws, row, 1, row, 6);
                                xlsService.ExcelSetHorizontalAlignLeft(ws, row, 1, row, 6);
                                //ws.Row(row).Height = 70;
                                xlsService.XlsSetRow(ws, row, general, 120, 70);
                                row++;
                            }
                        }
                        #endregion

                        //string startupPath = Directory.GetCurrentDirectory();
                        //string strSourceDownload = Path.Combine(startupPath, "include", "questionnaire", "download"); 031421
                        string strSourceDownload = Path.Combine(startupPath, "include", "upload", "soxquestionnaire");
                        

                        if (!Directory.Exists(strSourceDownload))
                        {
                            Directory.CreateDirectory(strSourceDownload);
                        }
                        //var ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                        var ts = DateTime.Now.ToString("yyyyMMdd");
                        //string filename = $"{Guid.NewGuid()}.xlsx";
                        string Fy = roundSet.Rcm.FY;
                        //string process = roundSet.Rcm.Process;
                        string roundName = string.Empty;
                        if (roundSet.isRound1)
                            roundName = "Round 1";
                        else if (roundSet.isRound2)
                            roundName = "Round 2";
                        else if (roundSet.isRound3)
                            roundName = "Round 3";

                                               
                        string filename = $"{clientName} {Fy} {process} - {controlId} {roundName}-{ts}.xlsx";
                        string strOutput = Path.Combine(strSourceDownload, filename);

                        //Check if file not exists
                        if (System.IO.File.Exists(strOutput))
                        {
                            System.IO.File.Delete(strOutput);
                        }

                        xls.SaveAs(new FileInfo(strOutput));
                        excelFilename = filename;
                    }


                }

            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorCreateQuestionnaireExcel");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "CreateQuestionnaireExcel");
                Debug.WriteLine(ex);
            }

            return excelFilename;
        }

        [HttpPost("field")]
        public IEnumerable<QuestionnaireQuestion> GetQuestionnaireField([FromBody] QuestionnaireFieldParam QuestionnaireFieldParam)
        {
            List<QuestionnaireQuestion> QuestionnaireQuestion = null;
            try
            {
                //QuestionnaireQuestion = _soxContext.QuestionnaireQuestion
                //    .Where(x => x.AppId == QuestionnaireFieldParam.AppId)
                //    .Include(x => x.Options)
                //    .ToList();

                QuestionnaireQuestion = _soxContext.QuestionnaireQuestion
                    .Where(x =>
                        x.ClientName.Equals(QuestionnaireFieldParam.ClientName) &&
                        x.ControlName.Equals(QuestionnaireFieldParam.ControlName) &&
                        x.AppId.Equals(QuestionnaireFieldParam.AppKey.AppId))
                    .Include(x => x.Options)
                    .AsNoTracking()
                    .ToList();
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetQuestionnaireField {ex}", "ErrorGetQuestionnaireField");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetQuestionnaireField");
            }

            return QuestionnaireQuestion.ToArray();

        }


        [HttpPost("appId")]
        public IActionResult GetQuestionnaireAppId([FromBody] QuestionnaireFieldParam QuestionnaireFieldParam)
        {

            string AppId = string.Empty;
            try
            {
                AppId = _soxContext.QuestionnaireQuestion
                    .Where(x => x.ClientName.Equals(QuestionnaireFieldParam.ClientName) && x.ControlName.Equals(QuestionnaireFieldParam.ControlName))
                    .Select(x => x.AppId)
                    .AsNoTracking()
                    .FirstOrDefault();
            }
            catch (Exception ex)
            {
                //ErrorLog.Write(ex);
                FileLog.Write($"Error GetQuestionnaireAppId {ex}", "ErrorGetQuestionnaireAppId");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetQuestionnaireAppId");
                return BadRequest($"");
            }

            return Ok(AppId);

        }

        //[AllowAnonymous]
        [HttpPost("iucsystem")]
        public IActionResult GetIUCSystemGenField([FromBody] QuestionnaireFieldParam QuestionnaireFieldParam)
        {
            IUCSystemGen IUCSystemGen = null;
            try
            {

                IUCSystemGen = _soxContext.IUCSystemGen
                    .Where(x =>
                        x.AppId.Equals(QuestionnaireFieldParam.AppKey.AppId))
                    .Include(x => x.ListQuestionAnswer).AsNoTracking().FirstOrDefault();
            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorGetIUCSystemGenField");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetIUCSystemGenField");
            }

            if (IUCSystemGen != null)
                return Ok(IUCSystemGen);
            else
                return BadRequest();

        }

        //[AllowAnonymous]
        [HttpPost("iucnonsystem")]
        public IActionResult GetIUCNonSystemGenField([FromBody] QuestionnaireFieldParam QuestionnaireFieldParam)
        {
            IUCNonSystemGen IUCNonSystemGen = null;
            try
            {

                IUCNonSystemGen = _soxContext.IUCNonSystemGen
                    .Where(x =>
                        x.AppId.Equals(QuestionnaireFieldParam.AppKey.AppId))
                    .Include(x => x.ListQuestionAnswer).AsNoTracking().FirstOrDefault();

            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorGetIUCNonSystemGenField");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetIUCNonSystemGenField");
            }

            if (IUCNonSystemGen != null)
                return Ok(IUCNonSystemGen);
            else
                return BadRequest();

        }


        [HttpPost("podio/create/questionnaire")]
        public async Task<IActionResult> CreatePodioQuestionnaire([FromBody] List<QuestionnaireUserInput> ListQuestionnaireInput)
        {
            bool status = false;
            try
            {
                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = ListQuestionnaireInput.Where(x => x.AppId != string.Empty).Select(x => x.AppId).FirstOrDefault();
                //PodioAppKey.AppToken = _config.GetSection("QuestionnaireApp").AsEnumerable().Where(x => x.Key.Contains(PodioAppKey.AppId)).Select(x => x.Value).FirstOrDefault();
                PodioAppKey.AppToken = GetAppToken(PodioAppKey.AppId);

                if (PodioAppKey.AppId != string.Empty && PodioAppKey.AppToken != string.Empty)
                {
                    var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                    ServicePointManager.Expect100Continue = true;
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                    await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                    if (podio.IsAuthenticated() && ListQuestionnaireInput.Count != 0)
                    {
                        Item questionnaireItem = new Item();

                        foreach (var item in ListQuestionnaireInput)
                        {

                            if (item.FieldId != null)
                            {
                                switch (item.Type)
                                {
                                    case "text":
                                        if (item.StrAnswer != null && item.StrAnswer != string.Empty)
                                        {
                                            var textItem = questionnaireItem.Field<TextItemField>(item.FieldId.Value);
                                            textItem.Value = item.StrAnswer;
                                        }
                                        break;
                                    case "category":
                                        if (item.StrAnswer != null && item.StrAnswer != string.Empty)
                                        {
                                            var categoryItem = questionnaireItem.Field<CategoryItemField>(item.FieldId.Value);
                                            categoryItem.OptionText = item.StrAnswer;
                                        }
                                        break;
                                    case "date":
                                        if (item.StrAnswer != null && item.StrAnswer != string.Empty)
                                        {


                                            DateTime dtValue, dtValue2;
                                            if (DateTime.TryParse(item.StrAnswer, out dtValue))
                                            {
                                                var dateField = questionnaireItem.Field<DateItemField>(item.FieldId.Value);
                                                dateField.Start = dtValue;
                                                if (item.StrAnswer2 != null && item.StrAnswer2 != string.Empty)
                                                {
                                                    if (DateTime.TryParse(item.StrAnswer2, out dtValue2))
                                                        dateField.End = dtValue2;
                                                    else
                                                        dateField.End = null;
                                                }
                                            }

                                        }
                                        break;
                                    case "app":
                                        if (item.StrQuestion == "Round Reference")
                                        {
                                            if (item.ListRoundItem != null && item.ListRoundItem.Count > 0)
                                            {
                                                var appReference = questionnaireItem.Field<AppItemField>(item.FieldId.Value);
                                                List<int> listRoundItem = new List<int>();

                                                //table 1
                                                if (item.ListRoundItem != null && item.ListRoundItem.Count > 0)
                                                {
                                                    foreach (var round in item.ListRoundItem)
                                                    {
                                                        if (round.PodioItemId != 0)
                                                            listRoundItem.Add(round.PodioItemId);
                                                    }
                                                }

                                                //table 2
                                                if (item.ListRoundItem2 != null && item.ListRoundItem2.Count > 0)
                                                {
                                                    foreach (var round in item.ListRoundItem2)
                                                    {
                                                        if (round.PodioItemId != 0)
                                                            listRoundItem.Add(round.PodioItemId);
                                                    }
                                                }

                                                //table 3
                                                if (item.ListRoundItem3 != null && item.ListRoundItem3.Count > 0)
                                                {
                                                    foreach (var round in item.ListRoundItem3)
                                                    {
                                                        if (round.PodioItemId != 0)
                                                            listRoundItem.Add(round.PodioItemId);
                                                    }
                                                }

                                                appReference.ItemIds = listRoundItem;
                                            }
                                        }
                                        else if (item.StrQuestion == "Unique Notes Reference")
                                        {
                                            if (item.ListNoteItem != null && item.ListNoteItem.Count > 0)
                                            {
                                                var appReference = questionnaireItem.Field<AppItemField>(item.FieldId.Value);
                                                List<int> listRoundItem = new List<int>();
                                                foreach (var round in item.ListNoteItem)
                                                {
                                                    if (round.PodioItemId != 0)
                                                        listRoundItem.Add(round.PodioItemId);
                                                }
                                                appReference.ItemIds = listRoundItem;
                                            }
                                        }
                                        else if (item.StrQuestion == "IUC System Generated")
                                        {
                                            if (item.ListIUCSystemGen != null && item.ListIUCSystemGen.Count > 0)
                                            {
                                                var appReference = questionnaireItem.Field<AppItemField>(item.FieldId.Value);
                                                List<int> listRoundItem = new List<int>();
                                                foreach (var round in item.ListIUCSystemGen)
                                                {
                                                    if (round.PodioItemId != 0)
                                                        listRoundItem.Add(round.PodioItemId);
                                                }
                                                appReference.ItemIds = listRoundItem;
                                            }
                                        }
                                        else if (item.StrQuestion == "IUC Non System Generated")
                                        {
                                            if (item.ListIUCNonSystemGen != null && item.ListIUCNonSystemGen.Count > 0)
                                            {
                                                var appReference = questionnaireItem.Field<AppItemField>(item.FieldId.Value);
                                                List<int> listRoundItem = new List<int>();
                                                foreach (var round in item.ListIUCNonSystemGen)
                                                {
                                                    if (round.PodioItemId != 0)
                                                        listRoundItem.Add(round.PodioItemId);
                                                }
                                                appReference.ItemIds = listRoundItem;
                                            }
                                        }
                                        break;
                                    case "duration":
                                        if (item.StrAnswer != null && item.StrAnswer != string.Empty)
                                        {
                                            TimeSpan ts;
                                            //var timeSpan = TimeSpan.TryParse(item.StrAnswer, CultureInfo.CurrentCulture, out ts);
                                            if (TimeSpan.TryParse(item.StrAnswer, CultureInfo.CurrentCulture, out ts))
                                            {
                                                DurationItemField durationField = questionnaireItem.Field<DurationItemField>(item.FieldId.Value);
                                                durationField.Value = ts;
                                            }
                                        }
                                        break;
                                }
                            }

                        }

                        var roundId = await podio.ItemService.AddNewItem(Int32.Parse(PodioAppKey.AppId), questionnaireItem);

                        foreach (var item in ListQuestionnaireInput)
                        {
                            item.ItemId = int.Parse(roundId.ToString());
                        }

                        status = true;
                    }
                }

            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorCreatePodioQuestionnaire");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "CreatePodioQuestionnaire");
                return BadRequest(ex.ToString());
            }
            if (status)
            {
                return Ok(ListQuestionnaireInput.ToArray());
            }
            else
            {
                return NoContent();
            }



        }


        [HttpPost("podio/create/questionnaire2")]
        public async Task<IActionResult> CreatePodioQuestionnaire2([FromBody] QuestionnaireRoundSet RoundSet)
        {
            bool status = false;
            try
            {
                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                if(RoundSet.isRound1)
                    PodioAppKey.AppId = RoundSet.ListUserInputRound1.Where(x => x.AppId != string.Empty).Select(x => x.AppId).FirstOrDefault();
                if (RoundSet.isRound2)
                    PodioAppKey.AppId = RoundSet.ListUserInputRound2.Where(x => x.AppId != string.Empty).Select(x => x.AppId).FirstOrDefault();
                if (RoundSet.isRound3)
                    PodioAppKey.AppId = RoundSet.ListUserInputRound3.Where(x => x.AppId != string.Empty).Select(x => x.AppId).FirstOrDefault();
                PodioAppKey.AppToken = GetAppToken(PodioAppKey.AppId);

                if (PodioAppKey.AppId != string.Empty && PodioAppKey.AppToken != string.Empty)
                {
                    var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                    ServicePointManager.Expect100Continue = true;
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                    await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                    if (podio.IsAuthenticated())
                    {
                        if(RoundSet.isRound1)
                        {
                            var tempRoundSet1 = await AddPodioQuestionnaire(podio, PodioAppKey, RoundSet, "Round 1");
                            if(tempRoundSet1 != null && tempRoundSet1.Count > 0)
                            {
                                RoundSet.ListUserInputRound1 = tempRoundSet1;
                            }
                        }

                        if (RoundSet.isRound2)
                        {
                            var tempRoundSet2 = await AddPodioQuestionnaire(podio, PodioAppKey, RoundSet, "Round 2");
                            if (tempRoundSet2 != null && tempRoundSet2.Count > 0)
                            {
                                RoundSet.ListUserInputRound2 = tempRoundSet2;
                            }
                        }

                        if (RoundSet.isRound3)
                        {
                            var tempRoundSet3 = await AddPodioQuestionnaire(podio, PodioAppKey, RoundSet, "Round 3");
                            if (tempRoundSet3 != null && tempRoundSet3.Count > 0)
                            {
                                RoundSet.ListUserInputRound3 = tempRoundSet3;
                            }
                        }

                        status = true;
                    }
                }

            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorCreatePodioQuestionnaire2");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "CreatePodioQuestionnaire2");
                return BadRequest(ex.ToString());
            }
            if (status)
            {
                return Ok(RoundSet);
            }
            else
            {
                return NoContent();
            }



        }


        [HttpPost("podio/create/questionnaire/notes")]
        public async Task<IActionResult> CreatePodioNotes([FromBody] List<NotesItem> ListNotesItem)
        {
            bool status = false;
            List<NotesItem> listNotes = new List<NotesItem>();
            try
            {
                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("QuestionnaireOtherApp").GetSection("NotesAppId").Value;
                PodioAppKey.AppToken = _config.GetSection("QuestionnaireOtherApp").GetSection("NotesAppToken").Value;

                if (PodioAppKey.AppId != string.Empty && PodioAppKey.AppToken != string.Empty)
                {
                    var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                    ServicePointManager.Expect100Continue = true;
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                    await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                    if (podio.IsAuthenticated() && ListNotesItem.Count != 0)
                    {

                        int q1Field = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("NotesAppField").GetSection("Field1").Value);
                        int q2Field = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("NotesAppField").GetSection("Field2").Value);
                        foreach (var item in ListNotesItem)
                        {
                            Item questionnaireItem = new Item();

                            if (item.Notes != string.Empty)
                            {
                                var textNotes = questionnaireItem.Field<TextItemField>(q1Field);
                                textNotes.Value = item.Notes;
                            }

                            if (item.Description != string.Empty)
                            {
                                var textDesc = questionnaireItem.Field<TextItemField>(q2Field);
                                textDesc.Value = item.Description;
                            }

                            var roundId = await podio.ItemService.AddNewItem(Int32.Parse(PodioAppKey.AppId), questionnaireItem);
                            item.PodioItemId = int.Parse(roundId.ToString());

                            listNotes.Add(item);
                            status = true;
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorCreatePodioNotes");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "CreatePodioNotes");
                return BadRequest(ex.ToString());
            }
            if (status)
            {
                return Ok(listNotes.ToArray());
            }
            else
            {
                return NoContent();
            }

        }


        [HttpPost("podio/create/questionnaire/testround")]
        public async Task<IActionResult> CreatePodioTestRound([FromBody] List<RoundItem> ListRoundItem)
        {
            bool status = false;
            List<RoundItem> listTestRound = new List<RoundItem>();
            try
            {
                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppId").Value;
                PodioAppKey.AppToken = _config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppToken").Value;

                if (PodioAppKey.AppId != string.Empty && PodioAppKey.AppToken != string.Empty)
                {
                    var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                    ServicePointManager.Expect100Continue = true;
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                    await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                    if (podio.IsAuthenticated() && ListRoundItem.Count != 0)
                    {

                        int Field1 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field1").Value);
                        int Field2 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field2").Value);
                        int Field3 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field3").Value);
                        int Field4 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field4").Value);
                        int Field5 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field5").Value);
                        int Field6 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field6").Value);
                        int Field7 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field7").Value);
                        int Field8 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field8").Value);
                        int Field9 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field9").Value);
                        int Field10 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field10").Value);
                        int Field11 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field11").Value);
                        int Field12 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field12").Value);
                        int Field13 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field13").Value);
                        int Field14 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field14").Value);
                        int Field15 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field15").Value);
                        int Field16 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field16").Value);
                        int Field17 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field17").Value);
                        int Field18 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field18").Value);
                        int Field19 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field19").Value);
                        int Field20 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field20").Value);
                        int Field21 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field21").Value);
                        int Field22 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field22").Value);
                        int Field23 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field23").Value);
                        int Field24 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field24").Value);
                        int Field25 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field25").Value);
                        int Field26 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field26").Value);
                        int Field27 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field27").Value);
                        int Field28 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field28").Value);
                        int Field29 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field29").Value);
                        int Field30 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field30").Value);
                        int Field31 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field31").Value);
                        int Field32 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field32").Value);
                        int Field33 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field33").Value);
                        int Field34 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field34").Value);
                        int Field35 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field35").Value);
                        int Field36 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field36").Value);
                        int Field37 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field37").Value);
                        int Field38 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field38").Value);
                        int Field39 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field39").Value);
                        int Field40 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field40").Value);
                        int Field41 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field41").Value);
                        int Field42 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field42").Value);
                        int Field43 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field43").Value);
                        int Field44 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field44").Value);
                        int Field45 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field45").Value);
                        int Field46 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field46").Value);
                        int Field47 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("Field47").Value);

                        int FieldOther1 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("OtherAnswer1").Value);
                        int FieldOther2 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("OtherAnswer2").Value);
                        int FieldOther3 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("OtherAnswer3").Value);
                        int FieldOther4 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("OtherAnswer4").Value);
                        int FieldOther5 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("OtherAnswer5").Value);
                        int FieldOther6 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("OtherAnswer6").Value);
                        int FieldOther7 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("OtherAnswer7").Value);
                        int FieldOther8 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("OtherAnswer8").Value);
                        int FieldOther9 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("OtherAnswer9").Value);
                        int FieldOther10 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("OtherAnswer10").Value);
                        int FieldOther11 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("OtherAnswer11").Value);
                        int FieldOther12 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("OtherAnswer12").Value);
                        int FieldOther13 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("OtherAnswer13").Value);
                        int FieldOther14 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("OtherAnswer14").Value);
                        int FieldOther15 = int.Parse(_config.GetSection("QuestionnaireOtherApp").GetSection("RoundAppField").GetSection("OtherAnswer15").Value);



                        foreach (var item in ListRoundItem)
                        {
                            Item questionnaireItem = new Item();

                            #region Check Round 
                            if (item.RoundName != null && item.RoundName != string.Empty)
                            {
                                var textRoundName = questionnaireItem.Field<TextItemField>(Field1);
                                textRoundName.Value = item.RoundName;
                            }

                            if (item.Position != null && item.Position != string.Empty)
                            {
                                var textPosition = questionnaireItem.Field<TextItemField>(Field2);
                                textPosition.Value = item.Position;
                            }

                            if (item.Question1 != null && item.Question1 != string.Empty)
                            {
                                var textQ1 = questionnaireItem.Field<TextItemField>(Field3);
                                textQ1.Value = item.Question1;
                            }

                            if (item.Answer1 != null && item.Answer1 != string.Empty)
                            {
                                var textAns1 = questionnaireItem.Field<TextItemField>(Field4);
                                textAns1.Value = item.Answer1;
                            }

                            if (item.Note1 != null && item.Note1 != string.Empty)
                            {
                                var textNote1 = questionnaireItem.Field<TextItemField>(Field5);
                                textNote1.Value = item.Note1;
                            }

                            if (item.Question2 != null && item.Question2 != string.Empty)
                            {
                                var textQ2 = questionnaireItem.Field<TextItemField>(Field6);
                                textQ2.Value = item.Question2;
                            }

                            if (item.Answer2 != null && item.Answer2 != string.Empty)
                            {
                                var textAns2 = questionnaireItem.Field<TextItemField>(Field7);
                                textAns2.Value = item.Answer2;
                            }

                            if (item.Note2 != null && item.Note2 != string.Empty)
                            {
                                var textNote2 = questionnaireItem.Field<TextItemField>(Field8);
                                textNote2.Value = item.Note2;
                            }

                            if (item.Question3 != null && item.Question3 != string.Empty)
                            {
                                var textQ3 = questionnaireItem.Field<TextItemField>(Field9);
                                textQ3.Value = item.Question3;
                            }

                            if (item.Answer3 != null && item.Answer3 != string.Empty)
                            {
                                var textAns3 = questionnaireItem.Field<TextItemField>(Field10);
                                textAns3.Value = item.Answer3;
                            }

                            if (item.Note3 != null && item.Note3 != string.Empty)
                            {
                                var textNote3 = questionnaireItem.Field<TextItemField>(Field11);
                                textNote3.Value = item.Note3;
                            }

                            if (item.Question4 != null && item.Question4 != string.Empty)
                            {
                                var textQ4 = questionnaireItem.Field<TextItemField>(Field12);
                                textQ4.Value = item.Question4;
                            }

                            if (item.Answer4 != null && item.Answer4 != string.Empty)
                            {
                                var textAns4 = questionnaireItem.Field<TextItemField>(Field13);
                                textAns4.Value = item.Answer4;
                            }

                            if (item.Note4 != null && item.Note4 != string.Empty)
                            {
                                var textNote4 = questionnaireItem.Field<TextItemField>(Field14);
                                textNote4.Value = item.Note4;
                            }

                            if (item.Question5 != null && item.Question5 != string.Empty)
                            {
                                var textQ5 = questionnaireItem.Field<TextItemField>(Field15);
                                textQ5.Value = item.Question5;
                            }

                            if (item.Answer5 != null && item.Answer5 != string.Empty)
                            {
                                var textAns5 = questionnaireItem.Field<TextItemField>(Field16);
                                textAns5.Value = item.Answer5;
                            }

                            if (item.Note5 != null && item.Note5 != string.Empty)
                            {
                                var textNote5 = questionnaireItem.Field<TextItemField>(Field17);
                                textNote5.Value = item.Note5;
                            }

                            if (item.Question6 != null && item.Question6 != string.Empty)
                            {
                                var textQ6 = questionnaireItem.Field<TextItemField>(Field18);
                                textQ6.Value = item.Question6;
                            }

                            if (item.Answer6 != null && item.Answer6 != string.Empty)
                            {
                                var textAns6 = questionnaireItem.Field<TextItemField>(Field19);
                                textAns6.Value = item.Answer6;
                            }

                            if (item.Note6 != null && item.Note6 != string.Empty)
                            {
                                var textNote6 = questionnaireItem.Field<TextItemField>(Field20);
                                textNote6.Value = item.Note6;
                            }

                            if (item.Question7 != null && item.Question7 != string.Empty)
                            {
                                var textQ7 = questionnaireItem.Field<TextItemField>(Field21);
                                textQ7.Value = item.Question7;
                            }

                            if (item.Answer7 != null && item.Answer7 != string.Empty)
                            {
                                var textAns7 = questionnaireItem.Field<TextItemField>(Field22);
                                textAns7.Value = item.Answer7;
                            }

                            if (item.Note7 != null && item.Note7 != string.Empty)
                            {
                                var textNote7 = questionnaireItem.Field<TextItemField>(Field23);
                                textNote7.Value = item.Note7;
                            }

                            if (item.Question8 != null && item.Question8 != string.Empty)
                            {
                                var textQ8 = questionnaireItem.Field<TextItemField>(Field24);
                                textQ8.Value = item.Question8;
                            }

                            if (item.Answer8 != null && item.Answer8 != string.Empty)
                            {
                                var textAns8 = questionnaireItem.Field<TextItemField>(Field25);
                                textAns8.Value = item.Answer8;
                            }

                            if (item.Note8 != null && item.Note8 != string.Empty)
                            {
                                var textNote8 = questionnaireItem.Field<TextItemField>(Field26);
                                textNote8.Value = item.Note8;
                            }

                            if (item.Question9 != null && item.Question9 != string.Empty)
                            {
                                var textQ9 = questionnaireItem.Field<TextItemField>(Field27);
                                textQ9.Value = item.Question9;
                            }

                            if (item.Answer9 != null && item.Answer9 != string.Empty)
                            {
                                var textAns9 = questionnaireItem.Field<TextItemField>(Field28);
                                textAns9.Value = item.Answer9;
                            }

                            if (item.Note9 != null && item.Note9 != string.Empty)
                            {
                                var textNote9 = questionnaireItem.Field<TextItemField>(Field29);
                                textNote9.Value = item.Note9;
                            }

                            if (item.Question10 != null && item.Question10 != string.Empty)
                            {
                                var textQ10 = questionnaireItem.Field<TextItemField>(Field30);
                                textQ10.Value = item.Question10;
                            }

                            if (item.Answer10 != null && item.Answer10 != string.Empty)
                            {
                                var textAns10 = questionnaireItem.Field<TextItemField>(Field31);
                                textAns10.Value = item.Answer10;
                            }

                            if (item.Note10 != null && item.Note10 != string.Empty)
                            {
                                var textNote10 = questionnaireItem.Field<TextItemField>(Field32);
                                textNote10.Value = item.Note10;
                            }

                            if (item.Question11 != null && item.Question11 != string.Empty)
                            {
                                var textQ11 = questionnaireItem.Field<TextItemField>(Field33);
                                textQ11.Value = item.Question11;
                            }

                            if (item.Answer11 != null && item.Answer11 != string.Empty)
                            {
                                var textAns11 = questionnaireItem.Field<TextItemField>(Field34);
                                textAns11.Value = item.Answer11;
                            }

                            if (item.Note11 != null && item.Note11 != string.Empty)
                            {
                                var textNote11 = questionnaireItem.Field<TextItemField>(Field35);
                                textNote11.Value = item.Note11;
                            }

                            if (item.Question12 != null && item.Question12 != string.Empty)
                            {
                                var textQ12 = questionnaireItem.Field<TextItemField>(Field36);
                                textQ12.Value = item.Question12;
                            }

                            if (item.Answer12 != null && item.Answer12 != string.Empty)
                            {
                                var textAns12 = questionnaireItem.Field<TextItemField>(Field37);
                                textAns12.Value = item.Answer12;
                            }

                            if (item.Note12 != null && item.Note12 != string.Empty)
                            {
                                var textNote12 = questionnaireItem.Field<TextItemField>(Field38);
                                textNote12.Value = item.Note12;
                            }

                            if (item.Question13 != null && item.Question13 != string.Empty)
                            {
                                var textQ13 = questionnaireItem.Field<TextItemField>(Field39);
                                textQ13.Value = item.Question13;
                            }

                            if (item.Answer13 != null && item.Answer13 != string.Empty)
                            {
                                var textAns13 = questionnaireItem.Field<TextItemField>(Field40);
                                textAns13.Value = item.Answer13;
                            }

                            if (item.Note13 != null && item.Note13 != string.Empty)
                            {
                                var textNote13 = questionnaireItem.Field<TextItemField>(Field41);
                                textNote13.Value = item.Note13;
                            }

                            if (item.Question14 != null && item.Question14 != string.Empty)
                            {
                                var textQ14 = questionnaireItem.Field<TextItemField>(Field42);
                                textQ14.Value = item.Question14;
                            }

                            if (item.Answer14 != null && item.Answer14 != string.Empty)
                            {
                                var textAns14 = questionnaireItem.Field<TextItemField>(Field43);
                                textAns14.Value = item.Answer14;
                            }

                            if (item.Note14 != null && item.Note14 != string.Empty)
                            {
                                var textNote14 = questionnaireItem.Field<TextItemField>(Field44);
                                textNote14.Value = item.Note14;
                            }

                            if (item.Question15 != null && item.Question15 != string.Empty)
                            {
                                var textQ15 = questionnaireItem.Field<TextItemField>(Field45);
                                textQ15.Value = item.Question15;
                            }

                            if (item.Answer15 != null && item.Answer15 != string.Empty)
                            {
                                var textAns15 = questionnaireItem.Field<TextItemField>(Field46);
                                textAns15.Value = item.Answer15;
                            }

                            if (item.Note15 != null && item.Note15 != string.Empty)
                            {
                                var textNote15 = questionnaireItem.Field<TextItemField>(Field47);
                                textNote15.Value = item.Note15;
                            }

                            #region Other Answer
                            if (item.OtherAnswer1 != null && item.OtherAnswer1 != string.Empty)
                            {
                                var textOther1 = questionnaireItem.Field<TextItemField>(FieldOther1);
                                textOther1.Value = item.OtherAnswer1;
                            }
                            if (item.OtherAnswer2 != null && item.OtherAnswer2 != string.Empty)
                            {
                                var textOther2 = questionnaireItem.Field<TextItemField>(FieldOther2);
                                textOther2.Value = item.OtherAnswer2;
                            }
                            if (item.OtherAnswer3 != null && item.OtherAnswer3 != string.Empty)
                            {
                                var textOther3 = questionnaireItem.Field<TextItemField>(FieldOther3);
                                textOther3.Value = item.OtherAnswer3;
                            }
                            if (item.OtherAnswer4 != null && item.OtherAnswer4 != string.Empty)
                            {
                                var textOther4 = questionnaireItem.Field<TextItemField>(FieldOther4);
                                textOther4.Value = item.OtherAnswer4;
                            }
                            if (item.OtherAnswer5 != null && item.OtherAnswer5 != string.Empty)
                            {
                                var textOther5 = questionnaireItem.Field<TextItemField>(FieldOther5);
                                textOther5.Value = item.OtherAnswer5;
                            }
                            if (item.OtherAnswer6 != null && item.OtherAnswer6 != string.Empty)
                            {
                                var textOther6 = questionnaireItem.Field<TextItemField>(FieldOther6);
                                textOther6.Value = item.OtherAnswer6;
                            }
                            if (item.OtherAnswer7 != null && item.OtherAnswer7 != string.Empty)
                            {
                                var textOther7 = questionnaireItem.Field<TextItemField>(FieldOther7);
                                textOther7.Value = item.OtherAnswer7;
                            }
                            if (item.OtherAnswer8 != null && item.OtherAnswer8 != string.Empty)
                            {
                                var textOther8 = questionnaireItem.Field<TextItemField>(FieldOther8);
                                textOther8.Value = item.OtherAnswer8;
                            }
                            if (item.OtherAnswer9 != null && item.OtherAnswer9 != string.Empty)
                            {
                                var textOther9 = questionnaireItem.Field<TextItemField>(FieldOther9);
                                textOther9.Value = item.OtherAnswer9;
                            }
                            if (item.OtherAnswer10 != null && item.OtherAnswer10 != string.Empty)
                            {
                                var textOther10 = questionnaireItem.Field<TextItemField>(FieldOther10);
                                textOther10.Value = item.OtherAnswer10;
                            }
                            if (item.OtherAnswer11 != null && item.OtherAnswer11 != string.Empty)
                            {
                                var textOther11 = questionnaireItem.Field<TextItemField>(FieldOther11);
                                textOther11.Value = item.OtherAnswer11;
                            }
                            if (item.OtherAnswer12 != null && item.OtherAnswer12 != string.Empty)
                            {
                                var textOther12 = questionnaireItem.Field<TextItemField>(FieldOther12);
                                textOther12.Value = item.OtherAnswer12;
                            }
                            if (item.OtherAnswer13 != null && item.OtherAnswer13 != string.Empty)
                            {
                                var textOther13 = questionnaireItem.Field<TextItemField>(FieldOther13);
                                textOther13.Value = item.OtherAnswer13;
                            }
                            if (item.OtherAnswer14 != null && item.OtherAnswer14 != string.Empty)
                            {
                                var textOther14 = questionnaireItem.Field<TextItemField>(FieldOther14);
                                textOther14.Value = item.OtherAnswer14;
                            }
                            if (item.OtherAnswer15 != null && item.OtherAnswer15 != string.Empty)
                            {
                                var textOther15 = questionnaireItem.Field<TextItemField>(FieldOther15);
                                textOther15.Value = item.OtherAnswer15;
                            }
                            #endregion

                            #endregion
                            var roundId = await podio.ItemService.AddNewItem(Int32.Parse(PodioAppKey.AppId), questionnaireItem);

                            status = true;


                            RoundItem createdItemId = new RoundItem();
                            createdItemId = item;
                            createdItemId.PodioItemId = int.Parse(roundId.ToString());
                            listTestRound.Add(createdItemId);

                            status = true;
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorCreatePodioTestRound");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "CreatePodioTestRound");
                return BadRequest(ex.ToString());
            }
            if (status)
            {
                return Ok(listTestRound.ToArray());
            }
            else
            {
                return NoContent();
            }

        }

        //[AllowAnonymous]
        [HttpPost("podio/create/questionnaire/systemgen")]
        public async Task<IActionResult> CreatePodioIUCSystemGen([FromBody] List<IUCSystemGen> listIUCSystemGen)
        {
            bool status = false;
            List<IUCSystemGen> listCreatedIUCSystemGen = new List<IUCSystemGen>();

            try
            {
                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("QuestionnaireOtherApp").GetSection("SystemGenAppId").Value;
                PodioAppKey.AppToken = _config.GetSection("QuestionnaireOtherApp").GetSection("SystemGenToken").Value;

                if (PodioAppKey.AppId != string.Empty && PodioAppKey.AppToken != string.Empty)
                {
                    var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                    ServicePointManager.Expect100Continue = true;
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                    await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                    if (podio.IsAuthenticated() && listIUCSystemGen.Count != 0)
                    {
                        foreach (var item in listIUCSystemGen)
                        {
                            Item questionnaireItem = new Item();

                            #region Check Round 
                            //if (item.Position.ToString() != string.Empty)
                            //{
                            //    var textRoundName = questionnaireItem.Field<TextItemField>(213064215);
                            //    textRoundName.Value = item.Position.ToString();
                            //}

                            if (item.ListQuestionAnswer.Count > 0)
                            {
                                foreach (var innerField in item.ListQuestionAnswer)
                                {
                                    if (innerField.Answer != null && innerField.Answer != string.Empty)
                                    {
                                        var textQ1 = questionnaireItem.Field<TextItemField>(innerField.FieldId.Value);
                                        textQ1.Value = innerField.Answer;
                                    }
                                }
                            }


                            #endregion
                            var roundId = await podio.ItemService.AddNewItem(Int32.Parse(PodioAppKey.AppId), questionnaireItem);

                            status = true;


                            IUCSystemGen createdItemId = new IUCSystemGen();
                            createdItemId = item;
                            createdItemId.PodioItemId = int.Parse(roundId.ToString());
                            listCreatedIUCSystemGen.Add(createdItemId);

                            status = true;
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorCreatePodioIUCSystemGen");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "CreatePodioIUCSystemGen");
                return BadRequest(ex.ToString());
            }
            if (status)
            {
                return Ok(listCreatedIUCSystemGen.ToArray());
            }
            else
            {
                return NoContent();
            }

        }

        //[AllowAnonymous]
        [HttpPost("podio/create/questionnaire/nonsystemgen")]
        public async Task<IActionResult> CreatePodioIUCNonSystemGen([FromBody] List<IUCNonSystemGen> listIUCNonSystemGen)
        {
            bool status = false;
            List<IUCNonSystemGen> listCreatedIUCNonSystemGen = new List<IUCNonSystemGen>();

            try
            {
                PodioApiKey PodioKey = new PodioApiKey();
                PodioAppKey PodioAppKey = new PodioAppKey();
                PodioKey.ClientId = _config.GetSection("PodioApi").GetSection("ClientId").Value;
                PodioKey.ClientSecret = _config.GetSection("PodioApi").GetSection("ClientSecret").Value;
                PodioAppKey.AppId = _config.GetSection("QuestionnaireOtherApp").GetSection("NonSystemGenAppId").Value;
                PodioAppKey.AppToken = _config.GetSection("QuestionnaireOtherApp").GetSection("NonSystemGenToken").Value;

                if (PodioAppKey.AppId != string.Empty && PodioAppKey.AppToken != string.Empty)
                {
                    var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                    ServicePointManager.Expect100Continue = true;
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                    await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                    if (podio.IsAuthenticated() && listIUCNonSystemGen.Count != 0)
                    {
                        foreach (var item in listIUCNonSystemGen)
                        {
                            Item questionnaireItem = new Item();

                            #region Check Round 
                            //if (item.Position.ToString() != string.Empty)
                            //{
                            //    var textRoundName = questionnaireItem.Field<TextItemField>(213064215);
                            //    textRoundName.Value = item.Position.ToString();
                            //}

                            if (item.ListQuestionAnswer.Count > 0)
                            {
                                foreach (var innerField in item.ListQuestionAnswer)
                                {
                                    if (innerField.Answer != null && innerField.Answer != string.Empty)
                                    {
                                        var textQ1 = questionnaireItem.Field<TextItemField>(innerField.FieldId.Value);
                                        textQ1.Value = innerField.Answer;
                                    }
                                }
                            }


                            #endregion
                            var roundId = await podio.ItemService.AddNewItem(Int32.Parse(PodioAppKey.AppId), questionnaireItem);

                            status = true;


                            IUCNonSystemGen createdItemId = new IUCNonSystemGen();
                            createdItemId = item;
                            createdItemId.PodioItemId = int.Parse(roundId.ToString());
                            listCreatedIUCNonSystemGen.Add(createdItemId);

                            status = true;
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex.ToString());
                //ErrorLog.Write(ex);
                FileLog.Write(ex.ToString(), "ErrorCreatePodioIUCNonSystemGen");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "CreatePodioIUCNonSystemGen");
                return BadRequest(ex.ToString());
            }
            if (status)
            {
                return Ok(listCreatedIUCNonSystemGen.ToArray());
            }
            else
            {
                return NoContent();
            }

        }


        [HttpPost("listparam")]
        public IActionResult GetQuestionnaireListFields(string ClientName)
        {
            //Console.WriteLine("Triggered process");

            List<QuestionnaireFieldParam> listParam = new List<QuestionnaireFieldParam>();

            try
            {

                using (var clientContext = _soxContext.Database.BeginTransaction())
                {
                    var checClient = _soxContext.QuestionnaireFieldParam.Where(x => x.ClientName.Equals(ClientName)).Include(x => x.AppKey);
                    if (checClient != null && checClient.Count() > 0)
                    {
                        foreach (var item in checClient)
                        {
                            listParam.Add(item);
                        }
                    }

                }

            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorGetQuestionnaireListFields");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetQuestionnaireListFields");
                return BadRequest(ex.ToString());
            }

            return Ok(listParam);

        }


        [HttpPost("load/clientapp")]
        public async Task<IActionResult> GetClientParamExcel(string fileName)
        {
            List<QuestionnaireFieldParam> listParam = new List<QuestionnaireFieldParam>();
            try
            {
                #region Read Excel File
                //default filename "ClientApp.xlsx"
                string startupPath = Directory.GetCurrentDirectory();
                string path = Path.Combine(startupPath, "include", "questionnaire", fileName); 

                var fi = new FileInfo(path);

                //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage xls = new ExcelPackage(fi))
                {
                    ExcelWorksheet ws = xls.Workbook.Worksheets[0]; //load first sheet

                    int colCount = ws.Dimension.End.Column; //get total column
                    int rowCount = ws.Dimension.End.Row; // get total row
                    bool isEmpty;

                    //loop row
                    for (int row = 2; row <= rowCount; row++)
                    {
                        isEmpty = true;
                        QuestionnaireFieldParam questionnaireParam = new QuestionnaireFieldParam();
                        PodioAppKey appkey = new PodioAppKey();
                        //loop column
                        for (int col = 1; col <= colCount; col++)
                        {
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty && ws.Cells[row, col].Value?.ToString() != null)
                            {
                                isEmpty = false;
                            }
                            if (col == 1) { questionnaireParam.ClientName = ws.Cells[row, col].Value?.ToString(); }
                            else if (col == 2) { questionnaireParam.ControlName = ws.Cells[row, col].Value?.ToString(); }
                            else if (col == 3) { appkey.AppId = ws.Cells[row, col].Value?.ToString(); }
                            else if (col == 4) { appkey.AppToken = ws.Cells[row, col].Value?.ToString(); }
                        }

                        if (!isEmpty && appkey.AppId != string.Empty && questionnaireParam.ClientName != string.Empty)
                        {
                            questionnaireParam.AppKey = appkey;
                            listParam.Add(questionnaireParam);
                        }
                    }

                }

                #endregion

                #region Save Data


                using (var clientContext = _soxContext.Database.BeginTransaction())
                {

                    _soxContext.Database.ExecuteSqlRaw("TRUNCATE TABLE QuestionnaireFieldParam");

                    foreach (var item in listParam)
                    {
                        var checkAppKey = _soxContext.PodioAppKey.Where(x => x.AppId.Equals(item.AppKey.AppId)).FirstOrDefault();
                        if (checkAppKey != null)
                        {
                            item.AppKey = checkAppKey;
                        }
                        _soxContext.Add(item);
                        await _soxContext.SaveChangesAsync();

                    }

                    clientContext.Commit();

                }


                #endregion


            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorGetClientParamExcel");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetClientParamExcel");
                Debug.WriteLine(ex);
                return BadRequest();
            }

            return Ok(listParam);
        }


        [HttpPost("client/list")]
        public IActionResult GetClientQuestionnaireParam()
        {
            List<string> listClient = new List<string>();
            try
            {

                using (var clientContext = _soxContext.Database.BeginTransaction())
                {


                    var checkAppKey = _soxContext.QuestionnaireFieldParam.Select(x => x.ClientName).Distinct();
                    if (checkAppKey != null)
                    {
                        foreach (var item in checkAppKey)
                        {
                            listClient.Add(item);
                        }
                    }
                }


            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorGetClientQuestionnaireParam");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "GetClientQuestionnaireParam");
                Debug.WriteLine(ex);
                return BadRequest();
            }

            return Ok(listClient);
        }


        //[AllowAnonymous]
        [HttpPost("save")]
        public async Task<IActionResult> SaveQuestionnaire([FromBody] QuestionnaireRoundSet roundSet)
        {
            bool status = false;
            try
            {
                using (var context = _soxContext.Database.BeginTransaction())
                {
                    //QuestionnaireRoundSet _roundSet = new QuestionnaireRoundSet();
                    //_roundSet = roundSet;

                    //bool isExists = false;

                    var rcmCheck = _soxContext.Rcm.AsNoTracking().FirstOrDefault(id => id.PodioItemId == roundSet.Rcm.PodioItemId);
                    if (rcmCheck != null)
                    {
                        roundSet.Rcm.Id = rcmCheck.Id;
                    }


                    var questionnaireCheck = _soxContext.QuestionnaireRoundSet
                        .Where(id => id.UniqueId.Equals(roundSet.UniqueId))     
                        .Include(x => x.ListIUCSystemGen1).ThenInclude(x => x.ListQuestionAnswer)
                        .Include(x => x.ListIUCSystemGen2).ThenInclude(x => x.ListQuestionAnswer)
                        .Include(x => x.ListIUCSystemGen3).ThenInclude(x => x.ListQuestionAnswer)
                        .Include(x => x.ListIUCNonSystemGen1).ThenInclude(x => x.ListQuestionAnswer)
                        .Include(x => x.ListIUCNonSystemGen2).ThenInclude(x => x.ListQuestionAnswer)
                        .Include(x => x.ListIUCNonSystemGen3).ThenInclude(x => x.ListQuestionAnswer)
                        .Include(x => x.sampleSel1)
                        .Include(x => x.sampleSel2)
                        .Include(x => x.sampleSel3)
                        .FirstOrDefault();
                    if (questionnaireCheck != null)
                    {
                        //isExists = true;
                        //roundSet.Id = questionnaireCheck.Id;
                        //_soxContext.Entry(roundSet).State = EntityState.Modified;
                        //_soxContext.QuestionnaireRoundSet.Attach(roundSet);
                        //_soxContext.Update(roundSet);

                        //Update
                        roundSet.UpdatedOn = DateTime.Now;
                        _soxContext.Entry(questionnaireCheck).CurrentValues.SetValues(roundSet);
                        await _soxContext.SaveChangesAsync();

                        //check sample selection
                        if (roundSet.sampleSel1 != null && roundSet.sampleSel1.Id == 0)
                        {
                            _soxContext.Add(roundSet.sampleSel1);
                            await _soxContext.SaveChangesAsync();

                            questionnaireCheck.sampleSel1 = roundSet.sampleSel1;
                            await _soxContext.SaveChangesAsync();
                        }
                        else
                        {
                            //check if sample selection is null
                            if(questionnaireCheck.sampleSel1 != null)
                            {
                                //update
                                 _soxContext.Entry(questionnaireCheck.sampleSel1).CurrentValues.SetValues(roundSet.sampleSel1);
                                await _soxContext.SaveChangesAsync();
                            }
                           
                        }

                        if (roundSet.sampleSel2 != null && roundSet.sampleSel2.Id == 0)
                        {
                            //_soxContext.Add(roundSet.sampleSel2);
                            //await _soxContext.SaveChangesAsync();

                            //var entity = _soxContext.QuestionnaireRoundSet.Where(id => id.UniqueId.Equals(roundSet.UniqueId)).Include(x => x.sampleSel2).FirstOrDefault();
                            //if (entity != null)
                            //{
                            //    _soxContext.Entry(entity.sampleSel2).CurrentValues.SetValues(roundSet.sampleSel2);
                            //    await _soxContext.SaveChangesAsync();
                            //}

                            //_soxContext.Entry(questionnaireCheck.sampleSel1).CurrentValues.SetValues(roundSet.sampleSel2);
                            //await _soxContext.SaveChangesAsync();
                            _soxContext.Add(roundSet.sampleSel2);
                            await _soxContext.SaveChangesAsync();

                            questionnaireCheck.sampleSel2 = roundSet.sampleSel2;
                            await _soxContext.SaveChangesAsync();
                        }
                        else
                        {
                            if(questionnaireCheck.sampleSel2 != null)
                            {
                                _soxContext.Entry(questionnaireCheck.sampleSel2).CurrentValues.SetValues(roundSet.sampleSel2);
                                await _soxContext.SaveChangesAsync();
                            }
                            
                        }

                        if (roundSet.sampleSel3 != null && roundSet.sampleSel3.Id == 0)
                        {
                            _soxContext.Add(roundSet.sampleSel3);
                            await _soxContext.SaveChangesAsync();

                            questionnaireCheck.sampleSel3 = roundSet.sampleSel3;
                            await _soxContext.SaveChangesAsync();

                        }
                        else
                        {
                            if(questionnaireCheck.sampleSel3 != null)
                            {
                                _soxContext.Entry(questionnaireCheck.sampleSel3).CurrentValues.SetValues(roundSet.sampleSel3);
                                await _soxContext.SaveChangesAsync();
                            }
                            
                        }


                        foreach (var item in roundSet.ListUserInputRound1)
                        {
                            if (item.IsDisabled)
                                item.StrAnswer = item.StrDefaultAnswer;

                            if (item.Id != 0)
                            {
                                var entity = _soxContext.QuestionnaireUserAnswer.Where(x => x.Id.Equals(item.Id)).FirstOrDefault();
                                if (entity != null)
                                {
                                    _soxContext.Entry(entity).CurrentValues.SetValues(item);
                                    await _soxContext.SaveChangesAsync();
                                }
                            }
                        }

                        foreach (var item in roundSet.ListUserInputRound2)
                        {
                            if (item.IsDisabled)
                                item.StrAnswer = item.StrDefaultAnswer;
                            if (item.Id != 0)
                            {
                                var entity = _soxContext.QuestionnaireUserAnswer.Where(x => x.Id.Equals(item.Id)).FirstOrDefault();
                                if (entity != null)
                                {
                                    _soxContext.Entry(entity).CurrentValues.SetValues(item);
                                    await _soxContext.SaveChangesAsync();
                                }
                                //_soxContext.QuestionnaireUserAnswer.Attach(item);
                                //_soxContext.Update(item);
                                //await _soxContext.SaveChangesAsync();
                            }
                        }

                        foreach (var item in roundSet.ListUserInputRound3)
                        {
                            if (item.IsDisabled)
                                item.StrAnswer = item.StrDefaultAnswer;
                            if (item.Id != 0)
                            {
                                var entity = _soxContext.QuestionnaireUserAnswer.Where(x => x.Id.Equals(item.Id)).FirstOrDefault();
                                if (entity != null)
                                {
                                    _soxContext.Entry(entity).CurrentValues.SetValues(item);
                                    await _soxContext.SaveChangesAsync();
                                }
                                //_soxContext.QuestionnaireUserAnswer.Attach(item);
                                //_soxContext.Update(item);
                                //await _soxContext.SaveChangesAsync();
                            }

                        }

                        foreach (var item in roundSet.ListUniqueNotes)
                        {
                            if (item.Id != 0)
                            {
                                var entity = _soxContext.NotesItem.Where(x => x.Id.Equals(item.Id)).FirstOrDefault();
                                if (entity != null)
                                {
                                    _soxContext.Entry(entity).CurrentValues.SetValues(item);
                                    await _soxContext.SaveChangesAsync();
                                }
                            }
                            else
                            {
                                item.QuestionnaireRoundSetId = questionnaireCheck.Id;
                                //_soxContext.Entry(item).State = EntityState.Added;
                                _soxContext.Add(item);
                                await _soxContext.SaveChangesAsync();
                            }
                        }

                        //validate deleted items in roundSet.ListUniqueNotes
                        var tempListNotesId = roundSet.ListUniqueNotes.Select(x => x.Id).ToList();
                        var allNotesItem = _soxContext.NotesItem.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_notesitem`({roundSet.Id});").ToList();
                        if (allNotesItem.Any())
                        {
                            //compare roundSet.ListUniqueNotes and allNotesItem
                            var tempAllNotesId = allNotesItem.Select(x => x.Id).ToList();

                            foreach (var roundId in tempAllNotesId)
                            {
                                if (!tempListNotesId.Contains(roundId))
                                {
                                    Debug.WriteLine($"Removing Round Item: {roundId}");
                                    var res = _soxContext.Database.ExecuteSqlRaw($"CALL `sox`.`sp_questionnaire_remove_notesitem`({roundId});");
                                    if (res == 0)
                                    {
                                        var roundSetRound = roundSet.ListUniqueNotes.Where(x => x.Id.Equals(roundId)).FirstOrDefault();
                                        if (roundSetRound != null)
                                            roundSet.ListUniqueNotes.Remove(roundSetRound);
                                    }
                                }
                            }
                        }


                        foreach (var item in roundSet.ListRoundItem)
                        {
                            if (item.Id != 0)
                            {
                                var entity = _soxContext.RoundItem.Where(x => x.Id.Equals(item.Id)).FirstOrDefault();
                                if (entity != null)
                                {
                                    _soxContext.Entry(entity).CurrentValues.SetValues(item);
                                    await _soxContext.SaveChangesAsync();
                                }
                            }
                            else
                            {
                                item.QuestionnaireRoundSetId = questionnaireCheck.Id;
                                //_soxContext.Entry(item).State = EntityState.Added;
                                _soxContext.Add(item);
                                await _soxContext.SaveChangesAsync();
                            }

                        }

                        //validate deleted items in roundSet.ListRoundItem
                        var tempListRoundId = roundSet.ListRoundItem.Select(x => x.Id).ToList();
                        var allRoundItem = _soxContext.RoundItem.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_rounditem`({roundSet.Id});").ToList();
                        if (allRoundItem.Any())
                        {
                            //compare roundSet.ListRoundItem and allRoundItem
                            var tempRoundListId = allRoundItem.Select(x => x.Id).ToList();

                            foreach (var roundId in tempRoundListId)
                            {
                                if (!tempListRoundId.Contains(roundId))
                                {
                                    Debug.WriteLine($"Removing Round Item: {roundId}");
                                    var res = _soxContext.Database.ExecuteSqlRaw($"CALL `sox`.`sp_questionnaire_remove_rounditem`({roundId});");
                                    if (res == 0)
                                    {
                                        var roundSetRound = roundSet.ListRoundItem.Where(x => x.Id.Equals(roundId)).FirstOrDefault();
                                        if (roundSetRound != null)
                                            roundSet.ListRoundItem.Remove(roundSetRound);
                                    }
                                }
                            }

                        }


                        foreach (var item in roundSet.ListHeaderNote)
                        {
                            if (item.Id != 0)
                            {
                                var entity = _soxContext.HeaderNote.Where(x => x.Id.Equals(item.Id)).FirstOrDefault();
                                if (entity != null)
                                {
                                    _soxContext.Entry(entity).CurrentValues.SetValues(item);
                                    await _soxContext.SaveChangesAsync();
                                }
                            }

                        }




                        #region IUC Processes

                        foreach (var item in roundSet.ListIUCSystemGen1)
                        {
                            if (item.Id != 0)
                            {
                                var entity = _soxContext.IUCSystemGenAnswer.Where(x => x.Id.Equals(item.Id)).FirstOrDefault();
                                if (entity != null)
                                {
                                    if (item.ListQuestionAnswer.Any())
                                    {
                                        foreach (var innerItem in item.ListQuestionAnswer)
                                        {
                                            var entity2 = _soxContext.IUCQuestionUserAnswer.Where(x => x.Id.Equals(innerItem.Id)).FirstOrDefault();
                                            if (entity2 != null)
                                            {
                                                _soxContext.Entry(entity2).CurrentValues.SetValues(innerItem);
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }
                                    }
                                    _soxContext.Entry(entity).CurrentValues.SetValues(item);
                                    await _soxContext.SaveChangesAsync();
                                }
                            }
                            else
                            {
                                //_soxContext.Add(item);
                                //await _soxContext.SaveChangesAsync();
                                questionnaireCheck.ListIUCSystemGen1.Add(item);
                                await _soxContext.SaveChangesAsync();

                            }
                        }

                        //validate deleted items in roundSet.ListIUCSystemGen1
                        var tempListIUCSystemGen1Id = roundSet.ListIUCSystemGen1.Select(x => x.Id).ToList();
                        var allListIUCSystemGen1 = _soxContext.IUCNonSystemGenAnswer.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_iuc_system_gen`({roundSet.Id},'Round 1');").ToList();
                        if (allListIUCSystemGen1.Any())
                        {
                            var tempListId = allListIUCSystemGen1.Select(x => x.Id).ToList();

                            foreach (var roundId in tempListId)
                            {
                                if (!tempListIUCSystemGen1Id.Contains(roundId))
                                {
                                    Debug.WriteLine($"Removing IUC System Item: {roundId}");
                                    var res = _soxContext.Database.ExecuteSqlRaw($"CALL `sox`.`sp_questionnaire_remove_iucsys`({roundId});");
                                    if (res == 0)
                                    {
                                        var roundSetRound = roundSet.ListIUCSystemGen1.Where(x => x.Id.Equals(roundId)).FirstOrDefault();
                                        if (roundSetRound != null)
                                            roundSet.ListIUCSystemGen1.Remove(roundSetRound);
                                    }
                                }
                            }
                        }

                        foreach (var item in roundSet.ListIUCSystemGen2)
                        {
                            if (item.Id != 0)
                            {
                                var entity = _soxContext.IUCSystemGenAnswer.Where(x => x.Id.Equals(item.Id)).FirstOrDefault();
                                if (entity != null)
                                {
                                    if (item.ListQuestionAnswer.Any())
                                    {
                                        foreach (var innerItem in item.ListQuestionAnswer)
                                        {
                                            var entity2 = _soxContext.IUCQuestionUserAnswer.Where(x => x.Id.Equals(innerItem.Id)).FirstOrDefault();
                                            if (entity2 != null)
                                            {
                                                _soxContext.Entry(entity2).CurrentValues.SetValues(innerItem);
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }
                                    }

                                    _soxContext.Entry(entity).CurrentValues.SetValues(item);
                                    await _soxContext.SaveChangesAsync();
                                }
                            }
                            else
                            {
                                //item.QuestionnaireRoundSetId2 = questionnaireCheck.Id;
                                //_soxContext.Add(item);
                                //await _soxContext.SaveChangesAsync();

                                //_soxContext.Entry(item).State = EntityState.Added;
                                //_soxContext.Entry(roundSet).State = EntityState.Modified;
                                //_soxContext.Add(item);
                                questionnaireCheck.ListIUCSystemGen2.Add(item);
                                await _soxContext.SaveChangesAsync();
                            }

                        }

                        //validate deleted items in roundSet.ListIUCSystemGen2
                        var tempListIUCSystemGen2Id = roundSet.ListIUCSystemGen2.Select(x => x.Id).ToList();
                        var allListIUCSystemGen2 = _soxContext.IUCNonSystemGenAnswer.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_iuc_system_gen`({roundSet.Id},'Round 2');").ToList();
                        if (allListIUCSystemGen2.Any())
                        {
                            var tempListId = allListIUCSystemGen2.Select(x => x.Id).ToList();

                            foreach (var roundId in tempListId)
                            {
                                if (!tempListIUCSystemGen2Id.Contains(roundId))
                                {
                                    Debug.WriteLine($"Removing IUC System Item: {roundId}");
                                    var res = _soxContext.Database.ExecuteSqlRaw($"CALL `sox`.`sp_questionnaire_remove_iucsys`({roundId});");
                                    if (res == 0)
                                    {
                                        var roundSetRound = roundSet.ListIUCSystemGen2.Where(x => x.Id.Equals(roundId)).FirstOrDefault();
                                        if (roundSetRound != null)
                                            roundSet.ListIUCSystemGen2.Remove(roundSetRound);
                                    }
                                }
                            }
                        }

                        foreach (var item in roundSet.ListIUCSystemGen3)
                        {
                            if (item.Id != 0)
                            {
                                var entity = _soxContext.IUCSystemGenAnswer.Where(x => x.Id.Equals(item.Id)).FirstOrDefault();
                                if (entity != null)
                                {
                                    if (item.ListQuestionAnswer.Any())
                                    {
                                        foreach (var innerItem in item.ListQuestionAnswer)
                                        {
                                            var entity2 = _soxContext.IUCQuestionUserAnswer.Where(x => x.Id.Equals(innerItem.Id)).FirstOrDefault();
                                            if (entity2 != null)
                                            {
                                                _soxContext.Entry(entity2).CurrentValues.SetValues(innerItem);
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }
                                    }

                                    _soxContext.Entry(entity).CurrentValues.SetValues(item);
                                    await _soxContext.SaveChangesAsync();
                                }
                            }
                            else
                            {
                                questionnaireCheck.ListIUCSystemGen3.Add(item);
                                await _soxContext.SaveChangesAsync();
                            }

                        }

                        //validate deleted items in roundSet.ListIUCSystemGen3
                        var tempListIUCSystemGen3Id = roundSet.ListIUCSystemGen3.Select(x => x.Id).ToList();
                        var allListIUCSystemGen3 = _soxContext.IUCNonSystemGenAnswer.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_iuc_system_gen`({roundSet.Id},'Round 3');").ToList();
                        if (allListIUCSystemGen3.Any())
                        {
                            var tempListId = allListIUCSystemGen3.Select(x => x.Id).ToList();

                            foreach (var roundId in tempListId)
                            {
                                if (!tempListIUCSystemGen3Id.Contains(roundId))
                                {
                                    Debug.WriteLine($"Removing IUC System Item: {roundId}");
                                    var res = _soxContext.Database.ExecuteSqlRaw($"CALL `sox`.`sp_questionnaire_remove_iucsys`({roundId});");
                                    if (res == 0)
                                    {
                                        var roundSetRound = roundSet.ListIUCSystemGen3.Where(x => x.Id.Equals(roundId)).FirstOrDefault();
                                        if (roundSetRound != null)
                                            roundSet.ListIUCSystemGen3.Remove(roundSetRound);
                                    }
                                }
                            }
                        }

                        foreach (var item in roundSet.ListIUCNonSystemGen1)
                        {
                            if (item.Id != 0)
                            {
                                var entity = _soxContext.IUCNonSystemGenAnswer.Where(x => x.Id.Equals(item.Id)).FirstOrDefault();
                                if (entity != null)
                                {
                                    if (item.ListQuestionAnswer.Any())
                                    {
                                        foreach (var innerItem in item.ListQuestionAnswer)
                                        {
                                            var entity2 = _soxContext.IUCQuestionUserAnswer.Where(x => x.Id.Equals(innerItem.Id)).FirstOrDefault();
                                            if (entity2 != null)
                                            {
                                                _soxContext.Entry(entity2).CurrentValues.SetValues(innerItem);
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }
                                    }

                                    _soxContext.Entry(entity).CurrentValues.SetValues(item);
                                    await _soxContext.SaveChangesAsync();
                                }
                            }
                            else
                            {
                                questionnaireCheck.ListIUCNonSystemGen1.Add(item);
                                await _soxContext.SaveChangesAsync();
                     
                            }

                        }

                        //validate deleted items in roundSet.ListIUCNonSystemGen1
                        var tempListIUCNonSystemGen1Id = roundSet.ListIUCNonSystemGen1.Select(x => x.Id).ToList();
                        var allListIUCNonSystemGen1 = _soxContext.IUCNonSystemGenAnswer.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_iuc_non_system_gen`({roundSet.Id},'Round 1');").ToList();
                        if (allListIUCNonSystemGen1.Any())
                        {
                            var tempListId = allListIUCNonSystemGen1.Select(x => x.Id).ToList();

                            foreach (var roundId in tempListId)
                            {
                                if (!tempListIUCNonSystemGen1Id.Contains(roundId))
                                {
                                    Debug.WriteLine($"Removing IUC Non System Item: {roundId}");
                                    var res = _soxContext.Database.ExecuteSqlRaw($"CALL `sox`.`sp_questionnaire_remove_iucnonsys`({roundId});");
                                    if (res == 0)
                                    {
                                        var roundSetRound = roundSet.ListIUCNonSystemGen1.Where(x => x.Id.Equals(roundId)).FirstOrDefault();
                                        if (roundSetRound != null)
                                            roundSet.ListIUCNonSystemGen1.Remove(roundSetRound);
                                    }
                                }
                            }
                        }


                        foreach (var item in roundSet.ListIUCNonSystemGen2)
                        {
                            if (item.Id != 0)
                            {
                                var entity = _soxContext.IUCNonSystemGenAnswer.Where(x => x.Id.Equals(item.Id)).FirstOrDefault();
                                if (entity != null)
                                {
                                    if (item.ListQuestionAnswer.Any())
                                    {
                                        foreach (var innerItem in item.ListQuestionAnswer)
                                        {
                                            var entity2 = _soxContext.IUCQuestionUserAnswer.Where(x => x.Id.Equals(innerItem.Id)).FirstOrDefault();
                                            if (entity2 != null)
                                            {
                                                _soxContext.Entry(entity2).CurrentValues.SetValues(innerItem);
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }
                                    }

                                    _soxContext.Entry(entity).CurrentValues.SetValues(item);
                                    await _soxContext.SaveChangesAsync();
                                }
                            }
                            else
                            {
                                //item.QuestionnaireRoundSetId2 = questionnaireCheck.Id;
                                //_soxContext.Add(item);
                                //await _soxContext.SaveChangesAsync();
                                questionnaireCheck.ListIUCNonSystemGen2.Add(item);
                                await _soxContext.SaveChangesAsync();
                            }

                        }

                        //validate deleted items in roundSet.ListIUCNonSystemGen2
                        var tempListIUCNonSystemGen2Id = roundSet.ListIUCNonSystemGen2.Select(x => x.Id).ToList();
                        var allListIUCNonSystemGen2 = _soxContext.IUCNonSystemGenAnswer.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_iuc_non_system_gen`({roundSet.Id},'Round 2');").ToList();
                        if (allListIUCNonSystemGen2.Any())
                        {
                            var tempListId = allListIUCNonSystemGen2.Select(x => x.Id).ToList();

                            foreach (var roundId in tempListId)
                            {
                                if (!tempListIUCNonSystemGen2Id.Contains(roundId))
                                {
                                    Debug.WriteLine($"Removing IUC Non System Item: {roundId}");
                                    var res = _soxContext.Database.ExecuteSqlRaw($"CALL `sox`.`sp_questionnaire_remove_iucnonsys`({roundId});");
                                    if (res == 0)
                                    {
                                        var roundSetRound = roundSet.ListIUCNonSystemGen2.Where(x => x.Id.Equals(roundId)).FirstOrDefault();
                                        if (roundSetRound != null)
                                            roundSet.ListIUCNonSystemGen2.Remove(roundSetRound);
                                    }
                                }
                            }
                        }

                        foreach (var item in roundSet.ListIUCNonSystemGen3)
                        {
                            if (item.Id != 0)
                            {
                                var entity = _soxContext.IUCNonSystemGenAnswer.Where(x => x.Id.Equals(item.Id)).FirstOrDefault();
                                if (entity != null)
                                {
                                    if (item.ListQuestionAnswer.Any())
                                    {
                                        foreach (var innerItem in item.ListQuestionAnswer)
                                        {
                                            var entity2 = _soxContext.IUCQuestionUserAnswer.Where(x => x.Id.Equals(innerItem.Id)).FirstOrDefault();
                                            if (entity2 != null)
                                            {
                                                _soxContext.Entry(entity2).CurrentValues.SetValues(innerItem);
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }
                                    }

                                    _soxContext.Entry(entity).CurrentValues.SetValues(item);
                                    await _soxContext.SaveChangesAsync();
                                }
                            }
                            else
                            {
                                //item.QuestionnaireRoundSetId3 = questionnaireCheck.Id;
                                //_soxContext.Add(item);
                                //await _soxContext.SaveChangesAsync();
                                questionnaireCheck.ListIUCNonSystemGen3.Add(item);
                                await _soxContext.SaveChangesAsync();
                            }

                        }

                        //validate deleted items in roundSet.ListIUCNonSystemGen3
                        var tempListIUCNonSystemGen3Id = roundSet.ListIUCNonSystemGen3.Select(x => x.Id).ToList();
                        var allListIUCNonSystemGen3 = _soxContext.IUCNonSystemGenAnswer.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_iuc_non_system_gen`({roundSet.Id},'Round 3');").ToList();
                        if (allListIUCNonSystemGen3.Any())
                        {
                            var tempListId = allListIUCNonSystemGen3.Select(x => x.Id).ToList();

                            foreach (var roundId in tempListId)
                            {
                                if (!tempListIUCNonSystemGen3Id.Contains(roundId))
                                {
                                    Debug.WriteLine($"Removing IUC Non System Item: {roundId}");
                                    var res = _soxContext.Database.ExecuteSqlRaw($"CALL `sox`.`sp_questionnaire_remove_iucnonsys`({roundId});");
                                    if (res == 0)
                                    {
                                        var roundSetRound = roundSet.ListIUCNonSystemGen3.Where(x => x.Id.Equals(roundId)).FirstOrDefault();
                                        if (roundSetRound != null)
                                            roundSet.ListIUCNonSystemGen3.Remove(roundSetRound);
                                    }
                                }
                            }
                        }

                        #endregion



                        context.Commit();
                    }
                    else
                    {
                        roundSet.CreatedOn = DateTime.Now;
                        roundSet.UpdatedOn = DateTime.Now;
                        _soxContext.Entry(roundSet.Rcm).State = EntityState.Unchanged;
                        _soxContext.Add(roundSet);
                        await _soxContext.SaveChangesAsync();
                        context.Commit();
                    }



                }

                status = true;
            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSaveQuestionnaire");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SaveQuestionnaire");
                return BadRequest(ex.ToString());
            }
            if (status)
            {
                return Ok(roundSet);
            }
            else
            {
                return NoContent();
            }



        }


        //[AllowAnonymous]
        [HttpPost("saveold")]
        public async Task<IActionResult> SaveQuestionnaireOLD([FromBody] QuestionnaireRoundSet roundSet)
        {
            bool status = false;
            try
            {
                using (var context = _soxContext.Database.BeginTransaction())
                {
                    //QuestionnaireRoundSet _roundSet = new QuestionnaireRoundSet();
                    //_roundSet = roundSet;

                    //bool isExists = false;

                    var rcmCheck = _soxContext.Rcm.AsNoTracking().FirstOrDefault(id => id.PodioItemId == roundSet.Rcm.PodioItemId);
                    if (rcmCheck != null)
                    {
                        roundSet.Rcm.Id = rcmCheck.Id;
                    }


                    var questionnaireCheck = _soxContext.QuestionnaireRoundSet
                        .Where(id => id.UniqueId.Equals(roundSet.UniqueId))
                        .Include(x => x.ListIUCNonSystemGen1).ThenInclude(x => x.ListQuestionAnswer)
                        .FirstOrDefault();
                    if (questionnaireCheck != null)
                    {
                        //isExists = true;
                        //roundSet.Id = questionnaireCheck.Id;
                        //_soxContext.Entry(roundSet).State = EntityState.Modified;
                        //_soxContext.QuestionnaireRoundSet.Attach(roundSet);
                        //_soxContext.Update(roundSet);
                        _soxContext.Entry(questionnaireCheck).CurrentValues.SetValues(roundSet);
                        await _soxContext.SaveChangesAsync();


                        foreach (var item in roundSet.ListUserInputRound1)
                        {
                            if (item.Id != 0)
                            {
                                var entity = _soxContext.QuestionnaireUserAnswer.Where(x => x.Id.Equals(item.Id)).FirstOrDefault(); 
                                if(entity != null)
                                {
                                    _soxContext.Entry(entity).CurrentValues.SetValues(item);
                                    await _soxContext.SaveChangesAsync();
                                }
                            }
                        }

                        foreach (var item in roundSet.ListUserInputRound2)
                        {
                            if (item.Id != 0)
                            {
                                var entity = _soxContext.QuestionnaireUserAnswer.Where(x => x.Id.Equals(item.Id)).FirstOrDefault();
                                if (entity != null)
                                {
                                    _soxContext.Entry(entity).CurrentValues.SetValues(item);
                                    await _soxContext.SaveChangesAsync();
                                }
                                //_soxContext.QuestionnaireUserAnswer.Attach(item);
                                //_soxContext.Update(item);
                                //await _soxContext.SaveChangesAsync();
                            }
                        }

                        foreach (var item in roundSet.ListUserInputRound3)
                        {
                            if (item.Id != 0)
                            {
                                var entity = _soxContext.QuestionnaireUserAnswer.Where(x => x.Id.Equals(item.Id)).FirstOrDefault();
                                if (entity != null)
                                {
                                    _soxContext.Entry(entity).CurrentValues.SetValues(item);
                                    await _soxContext.SaveChangesAsync();
                                }
                                //_soxContext.QuestionnaireUserAnswer.Attach(item);
                                //_soxContext.Update(item);
                                //await _soxContext.SaveChangesAsync();
                            }

                        }

                        foreach (var item in roundSet.ListUniqueNotes)
                        {
                            if (item.Id != 0)
                            {
                                var entity = _soxContext.NotesItem.Where(x => x.Id.Equals(item.Id)).FirstOrDefault();
                                if (entity != null)
                                {
                                    _soxContext.Entry(entity).CurrentValues.SetValues(item);
                                    await _soxContext.SaveChangesAsync();
                                }
                            }
                            else
                            {
                                item.QuestionnaireRoundSetId = questionnaireCheck.Id;
                                //_soxContext.Entry(item).State = EntityState.Added;
                                _soxContext.Add(item);
                                await _soxContext.SaveChangesAsync();
                            }
                        }

                        //validate deleted items in roundSet.ListUniqueNotes
                        var tempListNotesId = roundSet.ListUniqueNotes.Select(x => x.Id).ToList();
                        var allNotesItem = _soxContext.NotesItem.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_notesitem`({roundSet.Id});").ToList();
                        if (allNotesItem.Any())
                        {
                            //compare roundSet.ListUniqueNotes and allNotesItem
                            var tempAllNotesId = allNotesItem.Select(x => x.Id).ToList();

                            foreach (var roundId in tempAllNotesId)
                            {
                                if (!tempListNotesId.Contains(roundId))
                                {
                                    Debug.WriteLine($"Removing Round Item: {roundId}");
                                    var res = _soxContext.Database.ExecuteSqlRaw($"CALL `sox`.`sp_questionnaire_remove_notesitem`({roundId});");
                                    if (res == 0)
                                    {
                                        var roundSetRound = roundSet.ListUniqueNotes.Where(x => x.Id.Equals(roundId)).FirstOrDefault();
                                        if (roundSetRound != null)
                                            roundSet.ListUniqueNotes.Remove(roundSetRound);
                                    }
                                }
                            }
                        }


                        foreach (var item in roundSet.ListRoundItem)
                        {
                            if(item.Id != 0)
                            {
                                var entity = _soxContext.RoundItem.Where(x => x.Id.Equals(item.Id)).FirstOrDefault();
                                if (entity != null)
                                {
                                    _soxContext.Entry(entity).CurrentValues.SetValues(item);
                                    await _soxContext.SaveChangesAsync();
                                }
                            }
                            else
                            {
                                item.QuestionnaireRoundSetId = questionnaireCheck.Id;
                                //_soxContext.Entry(item).State = EntityState.Added;
                                _soxContext.Add(item);
                                await _soxContext.SaveChangesAsync();
                            }

                        }

                        //validate deleted items in roundSet.ListRoundItem
                        var tempListRoundId = roundSet.ListRoundItem.Select(x => x.Id).ToList();
                        var allRoundItem = _soxContext.RoundItem.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_rounditem`({roundSet.Id});").ToList();
                        if (allRoundItem.Any())
                        {
                            //compare roundSet.ListRoundItem and allRoundItem
                            var tempRoundListId = allRoundItem.Select(x => x.Id).ToList();

                            foreach (var roundId in tempRoundListId)
                            {
                                if (!tempListRoundId.Contains(roundId))
                                {
                                    Debug.WriteLine($"Removing Round Item: {roundId}");
                                    var res = _soxContext.Database.ExecuteSqlRaw($"CALL `sox`.`sp_questionnaire_remove_rounditem`({roundId});");
                                    if (res == 0)
                                    {
                                        var roundSetRound = roundSet.ListRoundItem.Where(x => x.Id.Equals(roundId)).FirstOrDefault();
                                        if (roundSetRound != null)
                                            roundSet.ListRoundItem.Remove(roundSetRound);
                                    }
                                }
                            }
                        
                        }


                        #region IUC Processes

                        foreach (var item in roundSet.ListIUCSystemGen1)
                        {
                            if (item.Id != 0)
                            {
                                var entity = _soxContext.IUCSystemGenAnswer.Where(x => x.Id.Equals(item.Id)).FirstOrDefault();
                                if (entity != null)
                                {
                                    if (item.ListQuestionAnswer.Any())
                                    {
                                        foreach (var innerItem in item.ListQuestionAnswer)
                                        {
                                            var entity2 = _soxContext.IUCQuestionUserAnswer.Where(x => x.Id.Equals(innerItem.Id)).FirstOrDefault();
                                            if (entity2 != null)
                                            {
                                                _soxContext.Entry(entity2).CurrentValues.SetValues(innerItem);
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }
                                    }
                                    _soxContext.Entry(entity).CurrentValues.SetValues(item);
                                    await _soxContext.SaveChangesAsync();
                                }
                            }
                            else
                            {
                                //_soxContext.Add(item);
                                //await _soxContext.SaveChangesAsync();
                                

                            }
                        }

                        //validate deleted items in roundSet.ListIUCSystemGen1
                        var tempListIUCSystemGen1Id = roundSet.ListIUCSystemGen1.Select(x => x.Id).ToList();
                        var allListIUCSystemGen1 = _soxContext.IUCNonSystemGenAnswer.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_iuc_system_gen`({roundSet.Id},'Round 1');").ToList();
                        if (allListIUCSystemGen1.Any())
                        {
                            var tempListId = allListIUCSystemGen1.Select(x => x.Id).ToList();

                            foreach (var roundId in tempListId)
                            {
                                if (!tempListIUCSystemGen1Id.Contains(roundId))
                                {
                                    Debug.WriteLine($"Removing IUC System Item: {roundId}");
                                    var res = _soxContext.Database.ExecuteSqlRaw($"CALL `sox`.`sp_questionnaire_remove_iucsys`({roundId});");
                                    if (res == 0)
                                    {
                                        var roundSetRound = roundSet.ListIUCSystemGen1.Where(x => x.Id.Equals(roundId)).FirstOrDefault();
                                        if (roundSetRound != null)
                                            roundSet.ListIUCSystemGen1.Remove(roundSetRound);
                                    }
                                }
                            }
                        }

                        foreach (var item in roundSet.ListIUCSystemGen2)
                        {
                            if (item.Id != 0)
                            {
                                var entity = _soxContext.IUCSystemGenAnswer.Where(x => x.Id.Equals(item.Id)).FirstOrDefault();
                                if (entity != null)
                                {
                                    if (item.ListQuestionAnswer.Any())
                                    {
                                        foreach (var innerItem in item.ListQuestionAnswer)
                                        {
                                            var entity2 = _soxContext.IUCQuestionUserAnswer.Where(x => x.Id.Equals(innerItem.Id)).FirstOrDefault();
                                            if (entity2 != null)
                                            {
                                                _soxContext.Entry(entity2).CurrentValues.SetValues(innerItem);
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }
                                    }

                                    _soxContext.Entry(entity).CurrentValues.SetValues(item);
                                    await _soxContext.SaveChangesAsync();
                                }
                            }
                            else
                            {
                                //item.QuestionnaireRoundSetId2 = questionnaireCheck.Id;
                                //_soxContext.Add(item);
                                //await _soxContext.SaveChangesAsync();

                                _soxContext.Entry(item).State = EntityState.Added;
                                _soxContext.Entry(roundSet).State = EntityState.Modified;
                                //_soxContext.Add(item);
                                await _soxContext.SaveChangesAsync();
                            }

                        }

                        //validate deleted items in roundSet.ListIUCSystemGen2
                        var tempListIUCSystemGen2Id = roundSet.ListIUCSystemGen2.Select(x => x.Id).ToList();
                        var allListIUCSystemGen2 = _soxContext.IUCNonSystemGenAnswer.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_iuc_system_gen`({roundSet.Id},'Round 2');").ToList();
                        if (allListIUCSystemGen2.Any())
                        {
                            var tempListId = allListIUCSystemGen2.Select(x => x.Id).ToList();

                            foreach (var roundId in tempListId)
                            {
                                if (!tempListIUCSystemGen2Id.Contains(roundId))
                                {
                                    Debug.WriteLine($"Removing IUC System Item: {roundId}");
                                    var res = _soxContext.Database.ExecuteSqlRaw($"CALL `sox`.`sp_questionnaire_remove_iucsys`({roundId});");
                                    if (res == 0)
                                    {
                                        var roundSetRound = roundSet.ListIUCSystemGen2.Where(x => x.Id.Equals(roundId)).FirstOrDefault();
                                        if (roundSetRound != null)
                                            roundSet.ListIUCSystemGen2.Remove(roundSetRound);
                                    }
                                }
                            }
                        }

                        foreach (var item in roundSet.ListIUCSystemGen3)
                        {
                            if (item.Id != 0)
                            {
                                var entity = _soxContext.IUCSystemGenAnswer.Where(x => x.Id.Equals(item.Id)).FirstOrDefault();
                                if (entity != null)
                                {
                                    if (item.ListQuestionAnswer.Any())
                                    {
                                        foreach (var innerItem in item.ListQuestionAnswer)
                                        {
                                            var entity2 = _soxContext.IUCQuestionUserAnswer.Where(x => x.Id.Equals(innerItem.Id)).FirstOrDefault();
                                            if (entity2 != null)
                                            {
                                                _soxContext.Entry(entity2).CurrentValues.SetValues(innerItem);
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }
                                    }

                                    _soxContext.Entry(entity).CurrentValues.SetValues(item);
                                    await _soxContext.SaveChangesAsync();
                                }
                            }
                            else
                            {
                                //item.QuestionnaireRoundSetId3 = questionnaireCheck.Id;
                                //_soxContext.Add(item);
                                //await _soxContext.SaveChangesAsync();
                            }

                        }

                        //validate deleted items in roundSet.ListIUCSystemGen3
                        var tempListIUCSystemGen3Id = roundSet.ListIUCSystemGen3.Select(x => x.Id).ToList();
                        var allListIUCSystemGen3 = _soxContext.IUCNonSystemGenAnswer.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_iuc_system_gen`({roundSet.Id},'Round 3');").ToList();
                        if (allListIUCSystemGen3.Any())
                        {
                            var tempListId = allListIUCSystemGen3.Select(x => x.Id).ToList();

                            foreach (var roundId in tempListId)
                            {
                                if (!tempListIUCSystemGen3Id.Contains(roundId))
                                {
                                    Debug.WriteLine($"Removing IUC System Item: {roundId}");
                                    var res = _soxContext.Database.ExecuteSqlRaw($"CALL `sox`.`sp_questionnaire_remove_iucsys`({roundId});");
                                    if (res == 0)
                                    {
                                        var roundSetRound = roundSet.ListIUCSystemGen3.Where(x => x.Id.Equals(roundId)).FirstOrDefault();
                                        if (roundSetRound != null)
                                            roundSet.ListIUCSystemGen3.Remove(roundSetRound);
                                    }
                                }
                            }
                        }

                        foreach (var item in roundSet.ListIUCNonSystemGen1)
                        {
                            if (item.Id != 0)
                            {
                                var entity = _soxContext.IUCNonSystemGenAnswer.Where(x => x.Id.Equals(item.Id)).FirstOrDefault();
                                if (entity != null)
                                {
                                    if (item.ListQuestionAnswer.Any())
                                    {
                                        foreach (var innerItem in item.ListQuestionAnswer)
                                        {
                                            var entity2 = _soxContext.IUCQuestionUserAnswer.Where(x => x.Id.Equals(innerItem.Id)).FirstOrDefault();
                                            if(entity2 != null)
                                            {
                                                _soxContext.Entry(entity2).CurrentValues.SetValues(innerItem);
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }
                                    }

                                    _soxContext.Entry(entity).CurrentValues.SetValues(item);
                                    await _soxContext.SaveChangesAsync();
                                }  
                            }
                            else
                            {
                                //item.QuestionnaireRoundSetId = questionnaireCheck.Id;
                                //_soxContext.Add(item);
                                //_soxContext.Entry(roundSet.list).State = EntityState.Added;
                                //_soxContext.Entry(roundSet).State = EntityState.Modified;
                                questionnaireCheck.ListIUCNonSystemGen1.Add(item);
                                await _soxContext.SaveChangesAsync();
                            }

                        }

                        //validate deleted items in roundSet.ListIUCNonSystemGen1
                        var tempListIUCNonSystemGen1Id = roundSet.ListIUCNonSystemGen1.Select(x => x.Id).ToList();
                        var allListIUCNonSystemGen1 = _soxContext.IUCNonSystemGenAnswer.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_iuc_non_system_gen`({roundSet.Id},'Round 1');").ToList();
                        if (allListIUCNonSystemGen1.Any())
                        {
                            var tempListId = allListIUCNonSystemGen1.Select(x => x.Id).ToList();

                            foreach (var roundId in tempListId)
                            {
                                if (!tempListIUCNonSystemGen1Id.Contains(roundId))
                                {
                                    Debug.WriteLine($"Removing IUC Non System Item: {roundId}");
                                    var res = _soxContext.Database.ExecuteSqlRaw($"CALL `sox`.`sp_questionnaire_remove_iucnonsys`({roundId});");
                                    if (res == 0)
                                    {
                                        var roundSetRound = roundSet.ListIUCNonSystemGen1.Where(x => x.Id.Equals(roundId)).FirstOrDefault();
                                        if (roundSetRound != null)
                                            roundSet.ListIUCNonSystemGen1.Remove(roundSetRound);
                                    }
                                }
                            }
                        }


                        foreach (var item in roundSet.ListIUCNonSystemGen2)
                        {
                            if (item.Id != 0)
                            {
                                var entity = _soxContext.IUCNonSystemGenAnswer.Where(x => x.Id.Equals(item.Id)).FirstOrDefault();
                                if (entity != null)
                                {
                                    if (item.ListQuestionAnswer.Any())
                                    {
                                        foreach (var innerItem in item.ListQuestionAnswer)
                                        {
                                            var entity2 = _soxContext.IUCQuestionUserAnswer.Where(x => x.Id.Equals(innerItem.Id)).FirstOrDefault();
                                            if (entity2 != null)
                                            {
                                                _soxContext.Entry(entity2).CurrentValues.SetValues(innerItem);
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }
                                    }

                                    _soxContext.Entry(entity).CurrentValues.SetValues(item);
                                    await _soxContext.SaveChangesAsync();
                                }
                            }
                            else
                            {
                                //item.QuestionnaireRoundSetId2 = questionnaireCheck.Id;
                                //_soxContext.Add(item);
                                //await _soxContext.SaveChangesAsync();
                            }

                        }

                        //validate deleted items in roundSet.ListIUCNonSystemGen2
                        var tempListIUCNonSystemGen2Id = roundSet.ListIUCNonSystemGen2.Select(x => x.Id).ToList();
                        var allListIUCNonSystemGen2 = _soxContext.IUCNonSystemGenAnswer.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_iuc_non_system_gen`({roundSet.Id},'Round 2');").ToList();
                        if (allListIUCNonSystemGen2.Any())
                        {
                            var tempListId = allListIUCNonSystemGen2.Select(x => x.Id).ToList();

                            foreach (var roundId in tempListId)
                            {
                                if (!tempListIUCNonSystemGen2Id.Contains(roundId))
                                {
                                    Debug.WriteLine($"Removing IUC Non System Item: {roundId}");
                                    var res = _soxContext.Database.ExecuteSqlRaw($"CALL `sox`.`sp_questionnaire_remove_iucnonsys`({roundId});");
                                    if (res == 0)
                                    {
                                        var roundSetRound = roundSet.ListIUCNonSystemGen2.Where(x => x.Id.Equals(roundId)).FirstOrDefault();
                                        if (roundSetRound != null)
                                            roundSet.ListIUCNonSystemGen2.Remove(roundSetRound);
                                    }
                                }
                            }
                        }

                        foreach (var item in roundSet.ListIUCNonSystemGen3)
                        {
                            if (item.Id != 0)
                            {
                                var entity = _soxContext.IUCNonSystemGenAnswer.Where(x => x.Id.Equals(item.Id)).FirstOrDefault();
                                if (entity != null)
                                {
                                    if (item.ListQuestionAnswer.Any())
                                    {
                                        foreach (var innerItem in item.ListQuestionAnswer)
                                        {
                                            var entity2 = _soxContext.IUCQuestionUserAnswer.Where(x => x.Id.Equals(innerItem.Id)).FirstOrDefault();
                                            if (entity2 != null)
                                            {
                                                _soxContext.Entry(entity2).CurrentValues.SetValues(innerItem);
                                                await _soxContext.SaveChangesAsync();
                                            }
                                        }
                                    }

                                    _soxContext.Entry(entity).CurrentValues.SetValues(item);
                                    await _soxContext.SaveChangesAsync();
                                }
                            }
                            else
                            {
                                //item.QuestionnaireRoundSetId3 = questionnaireCheck.Id;
                                //_soxContext.Add(item);
                                //await _soxContext.SaveChangesAsync();
                            }

                        }

                        //validate deleted items in roundSet.ListIUCNonSystemGen3
                        var tempListIUCNonSystemGen3Id = roundSet.ListIUCNonSystemGen3.Select(x => x.Id).ToList();
                        var allListIUCNonSystemGen3 = _soxContext.IUCNonSystemGenAnswer.FromSqlRaw($"CALL `sox`.`sp_get_questionnaire_iuc_non_system_gen`({roundSet.Id},'Round 3');").ToList();
                        if (allListIUCNonSystemGen3.Any())
                        {
                            var tempListId = allListIUCNonSystemGen3.Select(x => x.Id).ToList();

                            foreach (var roundId in tempListId)
                            {
                                if (!tempListIUCNonSystemGen3Id.Contains(roundId))
                                {
                                    Debug.WriteLine($"Removing IUC Non System Item: {roundId}");
                                    var res = _soxContext.Database.ExecuteSqlRaw($"CALL `sox`.`sp_questionnaire_remove_iucnonsys`({roundId});");
                                    if (res == 0)
                                    {
                                        var roundSetRound = roundSet.ListIUCNonSystemGen3.Where(x => x.Id.Equals(roundId)).FirstOrDefault();
                                        if (roundSetRound != null)
                                            roundSet.ListIUCNonSystemGen3.Remove(roundSetRound);
                                    }
                                }
                            }
                        }

                        #endregion


                        foreach (var item in roundSet.ListHeaderNote)
                        {
                            if (item.Id != 0)
                            {
                                var entity = _soxContext.HeaderNote.Where(x => x.Id.Equals(item.Id)).FirstOrDefault();
                                if (entity != null)
                                {
                                    _soxContext.Entry(entity).CurrentValues.SetValues(item);
                                    await _soxContext.SaveChangesAsync();
                                }
                            }

                        }


                        //_soxContext.Entry(questionnaireCheck).CurrentValues.SetValues(roundSet);
                        //_soxContext.QuestionnaireRoundSet.Attach(roundSet);
                        //_soxContext.Entry(roundSet.ListUserInputRound1).State = EntityState.Modified;
                        //questionnaireCheck = roundSet;
                        //questionnaireCheck.ListUserInputRound1 = roundSet.ListUserInputRound1;
                        //questionnaireCheck.ListUserInputRound2 = roundSet.ListUserInputRound2;
                        //questionnaireCheck.ListUserInputRound3 = roundSet.ListUserInputRound3;

                        //_soxContext.Update(questionnaireCheck);
                        //_soxContext.Entry(questionnaireCheck).State = EntityState.Modified;

                        //_soxContext.QuestionnaireRoundSet.Attach(roundSet);
                        //_soxContext.Update(roundSet);
                        //var entry = _soxContext.Entry(roundSet);
                        //entry.State = EntityState.Modified;

                        ////ListUniqueNotes
                        //foreach (var item in roundSet.ListUniqueNotes)
                        //{
                        //    var checkUniqueNotes = _soxContext.NotesItem.AsNoTracking().FirstOrDefault(x => x.Id.Equals(item.Id));
                        //    if (checkUniqueNotes == null)
                        //    {
                        //        //var res = _soxContext.Database.ExecuteSqlRaw("call rv_test_param(@id)", new MySqlParameter("@id", 1));
                        //        var res = _soxContext.Database.ExecuteSqlRaw($"CALL `sox`.`sp_insert_notesitem`({item.Notes}, {item.Description}, {item.PodioItemId}, {roundSet.Id});");
                        //    }
                        //}

                        context.Commit();
                    }
                    else
                    {
                        _soxContext.Entry(roundSet.Rcm).State = EntityState.Unchanged;
                        _soxContext.Add(roundSet);
                        await _soxContext.SaveChangesAsync();
                        context.Commit();
                    }



                }
                   
                status = true;
            }
            catch (Exception ex)
            {
                //Console.WriteLine(ex.ToString());
                FileLog.Write(ex.ToString(), "ErrorSaveQuestionnaire");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SaveQuestionnaire");
                return BadRequest(ex.ToString());
            }
            if (status)
            {
                return Ok(roundSet);
            }
            else
            {
                return NoContent();
            }



        }


        //[AllowAnonymous]
        [HttpPost("sox/tracker/create")]
        public void CreateSOXTrackerRecord([FromBody] QuestionnaireRoundSet roundSet)
        {
            try
            {
                // TODO sir Levin some data are static because can't find where to get those data in questionnaire input
                SoxTracker tracker = new SoxTracker(); // Tracker DATA
                                                                             // Defaults
                tracker.R1Tester = tracker.R11LReviewer = tracker.R12LReviewer = tracker.R2Tester = tracker.WTTester = tracker.WT1LReviewer = tracker.WT2LReviewer =
                    tracker.R21LReviewer = tracker.R22LReviewer = tracker.R3Tester = tracker.R31LReviewer = tracker.R32LReviewer = "N/A";
                tracker.RCRWT = tracker.RCRR1 = tracker.RCRR2 = tracker.RCRR3 = 1;
                tracker.R1TestingStatus = tracker.R2TestingStatus = tracker.R3TestingStatus = "N/A";
                tracker.Subprocess = "Control Activities";
                tracker.PBC = "test"; // TODO WHERE TO FIND THIS
                tracker.PBCOwner = "A2BQ"; // TODO WHERE TO FIND THIS
                tracker.FY = "FY"; // TODO WHERE TO FIND THIS
                tracker.PopulationFileRequest = "yes"; // TODO WHERE TO FIND THIS
                tracker.R3Sample = "1"; // TODO WHERE TO FIND THIS
                tracker.WTPBC = "N/A Non-Key"; // TODO WHERE TO FIND THIS
                tracker.R1PBC = tracker.R2PBC = tracker.R3PBC = "Not Tested (Year End Control)";// TODO WHERE TO FIND THIS
                tracker.WTTestingStatus = "In Progress"; // TODO WHERE TO FIND THIS
                tracker.ExternalAuditorSample = "Yes"; // TODO WHERE TO FIND THIS

                string clientName = roundSet.Rcm.ClientName;
                string controlId = roundSet.Rcm.ControlId;
                string controlOwner = roundSet.Rcm.ControlOwner;
                string testPerformed1 = string.Empty;
                string testPerformed2 = string.Empty;
                string testPerformed3 = string.Empty;
                string sampleSize1 = string.Empty;
                string sampleSize2 = string.Empty;
                string sampleSize3 = string.Empty;
                string reviewer1 = string.Empty;
                string reviewer2 = string.Empty;
                string reviewer3 = string.Empty;
                string testFindings1 = string.Empty;
                string testFindings2 = string.Empty;
                string testFindings3 = string.Empty;

                if (roundSet.isRound1)
                {
                    testFindings1 = roundSet.ListUserInputRound1.Where(x => x.StrQuestion.ToLower().Contains("what is the testing status")).Select(x => x.StrAnswer).FirstOrDefault();
                    reviewer1 = roundSet.ListUserInputRound1.Where(x => x.StrQuestion.ToLower().Contains("who performed the review")).Select(x => x.StrAnswer).FirstOrDefault();
                    sampleSize1 = roundSet.ListUserInputRound1.Where(x => x.StrQuestion.ToLower().Contains("what is the sample size") || x.StrQuestion.ToLower().Contains("what is the sub-sample size?")).Select(x => x.StrAnswer).FirstOrDefault();
                    testPerformed1 = roundSet.ListUserInputRound1.Where(x => x.StrQuestion.ToLower().Contains("performed the testing")).Select(x => x.StrAnswer).FirstOrDefault();
                }
                else if (roundSet.isRound2)
                {
                    testFindings2 = roundSet.ListUserInputRound2.Where(x => x.StrQuestion.ToLower().Contains("what is the testing status")).Select(x => x.StrAnswer).FirstOrDefault();
                    reviewer2 = roundSet.ListUserInputRound2.Where(x => x.StrQuestion.ToLower().Contains("who performed the review")).Select(x => x.StrAnswer).FirstOrDefault();
                    sampleSize2 = roundSet.ListUserInputRound2.Where(x => x.StrQuestion.ToLower().Contains("what is the sample size") || x.StrQuestion.ToLower().Contains("what is the sub-sample size?")).Select(x => x.StrAnswer).FirstOrDefault();
                    testPerformed2 = roundSet.ListUserInputRound2.Where(x => x.StrQuestion.ToLower().Contains("performed the testing")).Select(x => x.StrAnswer).FirstOrDefault();
                }
                else if (roundSet.isRound3)
                {
                    reviewer3 = roundSet.ListUserInputRound3.Where(x => x.StrQuestion.ToLower().Contains("who performed the review")).Select(x => x.StrAnswer).FirstOrDefault();
                    testFindings3 = roundSet.ListUserInputRound3.Where(x => x.StrQuestion.ToLower().Contains("what is the testing status")).Select(x => x.StrAnswer).FirstOrDefault();
                    sampleSize3 = roundSet.ListUserInputRound3.Where(x => x.StrQuestion.ToLower().Contains("what is the sample size") || x.StrQuestion.ToLower().Contains("what is the sub-sample size?")).Select(x => x.StrAnswer).FirstOrDefault();
                    testPerformed3 = roundSet.ListUserInputRound3.Where(x => x.StrQuestion.ToLower().Contains("performed the testing")).Select(x => x.StrAnswer).FirstOrDefault();
                }

                tracker.ClientName = tracker.Process = clientName != null ? clientName : string.Empty;
                tracker.ControlId = controlId != null ? controlId : string.Empty;
                tracker.PBCOwner = controlOwner != null ? controlOwner : string.Empty;
                tracker.SampleSelection = (sampleSize1 != null || sampleSize1 != null || sampleSize1 != null) ? "Yes" : "No";

                tracker.R1Tester = (testPerformed1 != null ? testPerformed1 : string.Empty);
                tracker.R2Tester = (testPerformed2 != null ? testPerformed2 : string.Empty);
                tracker.R3Tester = (testPerformed3 != null ? testPerformed3 : string.Empty);

                tracker.R11LReviewer = reviewer1;
                tracker.R21LReviewer = reviewer2;
                tracker.R31LReviewer = reviewer3;

                tracker.R1TestingStatus = testFindings1;
                tracker.R2TestingStatus = testFindings2;
                tracker.R3TestingStatus = testFindings3;

                this.Save2SOXTracker(tracker);
            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorCreateSOXTrackerRecord");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "CreateSOXTrackerRecord");
                Debug.WriteLine(ex);
            }
        }

        private string GetAppToken(string appId)
        {
            using (var context = _soxContext.Database.BeginTransaction())
            {
                var podioKey = _soxContext.PodioAppKey.AsNoTracking().Where(x => x.AppId.Equals(appId)).FirstOrDefault();
                if (podioKey != null)
                {
                    return podioKey.AppToken;
                }
                else
                {
                    return string.Empty;
                }
            }
        }

        private async Task<List<QuestionnaireUserAnswer>> AddPodioQuestionnaire(
            Podio podio,
            PodioAppKey podioKey,
            QuestionnaireRoundSet roundSet, 
            string roundName)
        {
            Item questionnaireItem = new Item();
            List<QuestionnaireUserAnswer> ListQuestionnaireInput = new List<QuestionnaireUserAnswer>();

            try
            {
                switch (roundName)
                {
                    case "Round 1":
                        ListQuestionnaireInput = roundSet.ListUserInputRound1.ToList();
                        break;
                    case "Round 2":
                        ListQuestionnaireInput = roundSet.ListUserInputRound2.ToList();
                        break;
                    case "Round 3":
                        ListQuestionnaireInput = roundSet.ListUserInputRound3.ToList();
                        break;

                }

                foreach (var item in ListQuestionnaireInput)
                {

                    if (item.FieldId != null)
                    {
                        switch (item.Type)
                        {
                            case "text":
                                if (item.StrAnswer != null && item.StrAnswer != string.Empty)
                                {
                                    var textItem = questionnaireItem.Field<TextItemField>(item.FieldId.Value);
                                    textItem.Value = item.StrAnswer;
                                }
                                break;
                            case "category":
                                if (item.StrAnswer != null && item.StrAnswer != string.Empty)
                                {
                                    var categoryItem = questionnaireItem.Field<CategoryItemField>(item.FieldId.Value);
                                    categoryItem.OptionText = item.StrAnswer;
                                }
                                break;
                            case "date":
                                if (item.StrAnswer != null && item.StrAnswer != string.Empty)
                                {


                                    DateTime dtValue, dtValue2;
                                    if (DateTime.TryParse(item.StrAnswer, out dtValue))
                                    {
                                        var dateField = questionnaireItem.Field<DateItemField>(item.FieldId.Value);
                                        dateField.Start = dtValue;
                                        if (item.StrAnswer2 != null && item.StrAnswer2 != string.Empty)
                                        {
                                            if (DateTime.TryParse(item.StrAnswer2, out dtValue2))
                                                dateField.End = dtValue2;
                                            else
                                                dateField.End = null;
                                        }
                                    }

                                }
                                break;
                            case "app":
                                if (item.StrQuestion == "Round Reference")
                                {
                                    if (roundSet.ListRoundItem != null && roundSet.ListRoundItem.Count > 0)
                                    {
                                        var appReference = questionnaireItem.Field<AppItemField>(item.FieldId.Value);
                                        List<int> listRoundItem = new List<int>();
                                        foreach (var round in roundSet.ListRoundItem)
                                        {
                                            if (round.PodioItemId != 0)
                                                listRoundItem.Add(round.PodioItemId);
                                        }


                                        ////table 2
                                        //if (item.ListRoundItem2 != null && item.ListRoundItem2.Count > 0)
                                        //{
                                        //    foreach (var round in item.ListRoundItem2)
                                        //    {
                                        //        if (round.PodioItemId != 0)
                                        //            listRoundItem.Add(round.PodioItemId);
                                        //    }
                                        //}

                                        ////table 3
                                        //if (item.ListRoundItem3 != null && item.ListRoundItem3.Count > 0)
                                        //{
                                        //    foreach (var round in item.ListRoundItem3)
                                        //    {
                                        //        if (round.PodioItemId != 0)
                                        //            listRoundItem.Add(round.PodioItemId);
                                        //    }
                                        //}
                                        if (listRoundItem.Count > 0)
                                        {
                                            appReference.ItemIds = listRoundItem;
                                        }

                                    }
                                }
                                else if (item.StrQuestion == "Unique Notes Reference")
                                {
                                    if (roundSet.ListUniqueNotes != null && roundSet.ListUniqueNotes.Count > 0)
                                    {
                                        var appReference = questionnaireItem.Field<AppItemField>(item.FieldId.Value);
                                        List<int> listNotesItem = new List<int>();
                                        foreach (var round in roundSet.ListUniqueNotes)
                                        {
                                            if (round.PodioItemId != 0)
                                                listNotesItem.Add(round.PodioItemId);
                                        }
                                        if (listNotesItem.Count > 0)
                                        {
                                            appReference.ItemIds = listNotesItem;
                                        }

                                    }

                                    //if (item.ListNoteItem != null && item.ListNoteItem.Count > 0)
                                    //{
                                    //    var appReference = questionnaireItem.Field<AppItemField>(item.FieldId.Value);
                                    //    List<int> listRoundItem = new List<int>();
                                    //    foreach (var round in item.ListNoteItem)
                                    //    {
                                    //        if (round.PodioItemId != 0)
                                    //            listRoundItem.Add(round.PodioItemId);
                                    //    }
                                    //    appReference.ItemIds = listRoundItem;
                                    //}


                                }
                                else if (item.StrQuestion == "IUC System Generated")
                                {
                                    var appReference = questionnaireItem.Field<AppItemField>(item.FieldId.Value);
                                    List<int> listIUCSystemGen = new List<int>();
                                    if (roundSet.ListIUCSystemGen1 != null && roundSet.ListIUCSystemGen1.Count > 0)
                                    {
                                        foreach (var round in roundSet.ListIUCSystemGen1)
                                        {
                                            if (round.PodioItemId != 0)
                                                listIUCSystemGen.Add(round.PodioItemId);
                                        }
                                    }
                                    if (roundSet.ListIUCSystemGen2 != null && roundSet.ListIUCSystemGen2.Count > 0)
                                    {
                                        foreach (var round in roundSet.ListIUCSystemGen2)
                                        {
                                            if (round.PodioItemId != 0)
                                                listIUCSystemGen.Add(round.PodioItemId);
                                        }
                                    }
                                    if (roundSet.ListIUCSystemGen3 != null && roundSet.ListIUCSystemGen3.Count > 0)
                                    {
                                        foreach (var round in roundSet.ListIUCSystemGen3)
                                        {
                                            if (round.PodioItemId != 0)
                                                listIUCSystemGen.Add(round.PodioItemId);
                                        }
                                    }
                                    if (listIUCSystemGen.Count > 0)
                                    {
                                        appReference.ItemIds = listIUCSystemGen;
                                    }

                                }
                                else if (item.StrQuestion == "IUC Non System Generated")
                                {
                                    var appReference = questionnaireItem.Field<AppItemField>(item.FieldId.Value);
                                    List<int> listUICNonSystemItem = new List<int>();
                                    if (roundSet.ListIUCNonSystemGen1 != null && roundSet.ListIUCNonSystemGen1.Count > 0)
                                    {
                                        foreach (var round in roundSet.ListIUCNonSystemGen1)
                                        {
                                            if (round.PodioItemId != 0)
                                                listUICNonSystemItem.Add(round.PodioItemId);
                                        }
                                    }
                                    if (roundSet.ListIUCNonSystemGen2 != null && roundSet.ListIUCNonSystemGen2.Count > 0)
                                    {
                                        foreach (var round in roundSet.ListIUCNonSystemGen2)
                                        {
                                            if (round.PodioItemId != 0)
                                                listUICNonSystemItem.Add(round.PodioItemId);
                                        }
                                    }
                                    if (roundSet.ListIUCNonSystemGen3 != null && roundSet.ListIUCNonSystemGen3.Count > 0)
                                    {
                                        foreach (var round in roundSet.ListIUCNonSystemGen3)
                                        {
                                            if (round.PodioItemId != 0)
                                                listUICNonSystemItem.Add(round.PodioItemId);
                                        }
                                    }
                                    if (listUICNonSystemItem.Count > 0)
                                    {
                                        appReference.ItemIds = listUICNonSystemItem;
                                    }
                                }
                                break;
                            case "duration":
                                if (item.StrAnswer != null && item.StrAnswer != string.Empty)
                                {
                                    TimeSpan ts;
                                    //var timeSpan = TimeSpan.TryParse(item.StrAnswer, CultureInfo.CurrentCulture, out ts);
                                    if (TimeSpan.TryParse(item.StrAnswer, CultureInfo.CurrentCulture, out ts))
                                    {
                                        DurationItemField durationField = questionnaireItem.Field<DurationItemField>(item.FieldId.Value);
                                        durationField.Value = ts;
                                    }
                                }
                                break;
                            case "image":
                                if (item.StrAnswer != null && item.StrAnswer != string.Empty)
                                {
                                    // Upload file
                                    //var filePath = Server.MapPath("\\files\\report.pdf");
                                    //var uploadedFile = podio.FileService.UploadFile(filePath, "report.pdf");
                                    string startupPath = Directory.GetCurrentDirectory();
                                    string filename = item.StrAnswer;
                                    string path = Path.Combine(startupPath, "include", "upload", "image", filename);
                                    var uploadedFile = await podio.FileService.UploadFile(path, filename);

                                    // Set FileIds
                                    if (uploadedFile != null)
                                    {
                                        //Debug.WriteLine($"uploadedFile: {uploadedFile.Result.FileId}");
                                        ImageItemField imageField = questionnaireItem.Field<ImageItemField>(item.FieldId.Value);
                                        imageField.FileIds = new List<int> { uploadedFile.FileId };
                                    }
                                    
                                }
                               
                                break;
                        }
                    }

                }

                var roundId = await podio.ItemService.AddNewItem(Int32.Parse(podioKey.AppId), questionnaireItem);

                foreach (var item in ListQuestionnaireInput)
                {
                    item.ItemId = int.Parse(roundId.ToString());
                }
            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorAddPodioWorkpaper");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "AddPodioWorkpaper");
            }

            

            return ListQuestionnaireInput;
        }

        private void Save2SOXTracker(SoxTracker tracker)
        {
            Debug.WriteLine("Saving Questionnaire DATA to SOX Tracker...");
            //using (var context = _soxContext.Database.BeginTransaction())
            //{
            //    _soxContext.Add(tracker);
            //    await _soxContext.SaveChangesAsync();
            //    context.Commit();
            //}
        }

        private Dictionary<string, string> GetMimeTypes()
        {
            return new Dictionary<string, string>
            {
                {".txt", "text/plain"},
                {".pdf", "application/pdf"},
                {".doc", "application/vnd.ms-word"},
                {".docx", "application/vnd.ms-word"},
                {".xls", "application/vnd.ms-excel"},
                {".xlsx", "application/vnd.ms-excel"},
                {".xlsm", "application/vnd.ms-excel.addin.macroEnabled.12"},
                {".png", "image/png"},
                {".jpg", "image/jpeg"},
            };
        }

        

    }
}
