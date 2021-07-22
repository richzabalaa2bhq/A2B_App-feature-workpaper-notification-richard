using A2B_App.Server.Data;
//using A2B_App.Server.Log;
using A2B_App.Shared.Podio;
using A2B_App.Shared.Sox;
using Microsoft.EntityFrameworkCore;
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
    public class RcmService
    {
        private SoxContext _soxContext;
        private readonly IConfiguration _config;
        public RcmService(SoxContext soxContext, IConfiguration config)
        {
            _soxContext = soxContext;
            _config = config;
        }


        /// <summary>
        /// Save webhook Create or Update
        /// </summary>
        /// <param name="podioHook">Podio Hook</param>
        /// <returns>bool</returns>
        public bool SaveWebhookAsync(PodioHook podioHook)
        {
            bool result = false;
            try
            {

                podioHook.DateCreated = DateTime.Now.ToUniversalTime();
                //podioHook.DateUpdated = DateTime.Now.ToUniversalTime();
                _soxContext.Add<PodioHook>(podioHook);
                _soxContext.SaveChanges();
                result = true;
            }
            catch (Exception ex)
            {
                FileLog.Write($"Error SaveWebhookRcmCta {ex}", "ErrorSaveWebhookRcmCta");
            }

            return result;
        }

        public async Task<bool> SavePodioRcmCtaAsync(int ItemId)
        {
            bool result = false;

            using (var context = _soxContext.Database.BeginTransaction())
            {
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

                        var podioItem = podio.ItemService.GetItem(ItemId);
                        var rawResponse = Newtonsoft.Json.JsonConvert.SerializeObject(podioItem.Result);
                        FileLog.Write($"{rawResponse}", "SavePodioRcmCtaAsync");

                        RcmCta rcmCta = new RcmCta();

                        //check if item_id already exists in the database
                        bool isExists = false;
                        var rcmCtaCheck = _soxContext.RcmCta.FirstOrDefault(id => id.PodioItemId == (int)podioItem.Result.ItemId);
                        if (rcmCtaCheck != null)
                        {
                            rcmCta = rcmCtaCheck;
                            isExists = true;
                        }

                        #region Podio Item
                        rcmCta.PodioItemId = (int)podioItem.Result.ItemId;
                        rcmCta.PodioUniqueId = podioItem.Result.AppItemIdFormatted.ToString();
                        rcmCta.PodioRevision = podioItem.Result.CurrentRevision.Revision;
                        rcmCta.PodioLink = podioItem.Result.Link.ToString();
                        rcmCta.CreatedBy = podioItem.Result.CreatedBy.Name.ToString();
                        rcmCta.CreatedOn = DateTime.Parse(podioItem.Result.CreatedOn.ToString());

                        #region sample selection fields

                        //1. Client Name
                        AppItemField clientApp = podioItem.Result.Field<AppItemField>("1-client-name");
                        IEnumerable<Item> clientAppRef = clientApp.Items;
                        rcmCta.ClientName = clientAppRef.Select(x => x.Title).FirstOrDefault();

                        //1. Client Code
                        CalculationItemField calcClientCode = podioItem.Result.Field<CalculationItemField>("1-client-code");
                        rcmCta.ClientCode = calcClientCode.ValueAsString;


                        //2. What is the sub-process?
                        CategoryItemField catSubProcess = podioItem.Result.Field<CategoryItemField>("what-is-the-sub-process");
                        IEnumerable<CategoryItemField.Option> listSubProcess = catSubProcess.Options;
                        rcmCta.SubProcess = listSubProcess.Select(x => x.Text).FirstOrDefault();

                        //3. What is the financial statement element?
                        TextItemField txtFinStatement = podioItem.Result.Field<TextItemField>("what-is-the-financial-statement-element-2");
                        rcmCta.FinancialStatementElement = txtFinStatement.Value;

                        //4. What is the specific risk?
                        TextItemField txtSpecificRisk = podioItem.Result.Field<TextItemField>("what-is-the-specific-risk");
                        rcmCta.SpecificRisk = txtSpecificRisk.Value;

                        //5.a. Does it show completeness & accuracy?
                        CategoryItemField q5a = podioItem.Result.Field<CategoryItemField>("5a-does-it-show-completeness-accuracy");
                        IEnumerable<CategoryItemField.Option> listQ5A = q5a.Options;
                        rcmCta.Q5ACompletenessAccuracy = listQ5A.Select(x => x.Text).FirstOrDefault();
                        //test 051921
                       
                       
                        //5.b. Does it show existence & occurrence?
                        CategoryItemField q5b = podioItem.Result.Field<CategoryItemField>("5b-does-it-show-existence-occurrence");
                        IEnumerable<CategoryItemField.Option> listQ5B = q5b.Options;
                        rcmCta.Q5BExistenceOccur = listQ5B.Select(x => x.Text).FirstOrDefault();

                        //5.c. Does it show presentation & disclosure?
                        CategoryItemField q5c = podioItem.Result.Field<CategoryItemField>("5c-does-it-show-presentation-disclosure");
                        IEnumerable<CategoryItemField.Option> listQ5C = q5c.Options;
                        rcmCta.Q5CPresentationDisclose = listQ5C.Select(x => x.Text).FirstOrDefault();

                        //5.d. Does it show rights & obligations?
                        CategoryItemField q5d = podioItem.Result.Field<CategoryItemField>("5d-does-it-show-rights-obligations");
                        IEnumerable<CategoryItemField.Option> listQ5D = q5d.Options;
                        rcmCta.Q5DRightObligation = listQ5D.Select(x => x.Text).FirstOrDefault();

                        //5.e. Does it show valuation & allocation?
                        CategoryItemField q5e = podioItem.Result.Field<CategoryItemField>("5e-does-it-show-valuation-allocation");
                        IEnumerable<CategoryItemField.Option> listQ5E = q5e.Options;
                        rcmCta.Q5EValuationAlloc = listQ5E.Select(x => x.Text).FirstOrDefault();

                        //6. What is the control ID?
                        CategoryItemField controlId = podioItem.Result.Field<CategoryItemField>("6-what-is-the-control-id-3");
                        IEnumerable<CategoryItemField.Option> listControlId = controlId.Options;
                        rcmCta.ControlId = listControlId.Select(x => x.Text).FirstOrDefault();

                        //7. What is the control activity?
                        TextItemField txtControlActivity = podioItem.Result.Field<TextItemField>("what-is-the-control-activity");
                        rcmCta.ControlActivity = txtControlActivity.Value;

                        //8. What is the control objective?
                        TextItemField txtControlObj = podioItem.Result.Field<TextItemField>("8-what-is-the-control-objective");
                        rcmCta.ControlObjective = txtControlObj.Value;

                        //9. When is the control in place date?
                        TextItemField txtControlPlace = podioItem.Result.Field<TextItemField>("when-is-the-control-in-place-2");
                        rcmCta.ControlPlaceDate = txtControlPlace.Value;

                        //10. Who is the control owner?
                        CategoryItemField controlOwner = podioItem.Result.Field<CategoryItemField>("who-is-the-control-owner");
                        IEnumerable<CategoryItemField.Option> listControlOwner = controlOwner.Options;
                        rcmCta.ControlOwner = listControlOwner.Select(x => x.Text).FirstOrDefault();

                        //11. What is the control frequency?
                        CategoryItemField frequency = podioItem.Result.Field<CategoryItemField>("what-is-the-control-frequency");
                        IEnumerable<CategoryItemField.Option> listFrequency = frequency.Options;
                        rcmCta.Frequency = listFrequency.Select(x => x.Text).FirstOrDefault();

                        //12. What is the entity?
                        CategoryItemField entity = podioItem.Result.Field<CategoryItemField>("what-is-the-entity-2");
                        IEnumerable<CategoryItemField.Option> listEntity = entity.Options;
                        rcmCta.Entity = listEntity.Select(x => x.Text).FirstOrDefault();

                        //12.a. Notes
                        TextItemField txtNotes = podioItem.Result.Field<TextItemField>("notes");
                        rcmCta.Notes = txtNotes.Value;

                        //13. Is it a Key/ Non-Key control?
                        CategoryItemField keyControl = podioItem.Result.Field<CategoryItemField>("is-it-a-key-non-key-control");
                        IEnumerable<CategoryItemField.Option> listKeyControl = keyControl.Options;
                        rcmCta.KeyControl = listKeyControl.Select(x => x.Text).FirstOrDefault();

                        //14. What is the control type?
                        CategoryItemField controlType = podioItem.Result.Field<CategoryItemField>("what-is-the-control-type");
                        IEnumerable<CategoryItemField.Option> listControlType = controlType.Options;
                        rcmCta.ControlType = listControlType.Select(x => x.Text).FirstOrDefault();

                        //15. What is the nature of procedure?
                        CategoryItemField natureProc = podioItem.Result.Field<CategoryItemField>("what-is-the-nature-of-procedure");
                        IEnumerable<CategoryItemField.Option> listNatureProc = natureProc.Options;
                        rcmCta.NatureProcedure = listNatureProc.Select(x => x.Text).FirstOrDefault();

                        //16. Is it a fraud control?
                        CategoryItemField fraudControl = podioItem.Result.Field<CategoryItemField>("fraud-control");
                        IEnumerable<CategoryItemField.Option> listFraudControl = fraudControl.Options;
                        rcmCta.FraudControl = listFraudControl.Select(x => x.Text).FirstOrDefault();

                        //17. What is the risk level?
                        CategoryItemField risk = podioItem.Result.Field<CategoryItemField>("risk-level");
                        IEnumerable<CategoryItemField.Option> listRisk = risk.Options;
                        rcmCta.Risk = listRisk.Select(x => x.Text).FirstOrDefault();

                        //18. Is it a management review control?
                        CategoryItemField revControl = podioItem.Result.Field<CategoryItemField>("management-review-control");
                        IEnumerable<CategoryItemField.Option> listRevControl = revControl.Options;
                        rcmCta.ReviewControl = listEntity.Select(x => x.Text).FirstOrDefault();

                        //19. What is the Testing Procedure?
                        TextItemField txtTestProc = podioItem.Result.Field<TextItemField>("test-procedures");
                        rcmCta.TestingProcedure = txtTestProc.Value;

                        //Population File Required
                        CategoryItemField populationFileReq = podioItem.Result.Field<CategoryItemField>("populationfilerequired");
                        IEnumerable<CategoryItemField.Option> listpopulationFileReq = populationFileReq.Options;
                        rcmCta.PopulationFileRequired = listpopulationFileReq.Select(x => x.Text).FirstOrDefault();


                        #endregion

                        #endregion

                        if (isExists)
                        {
                            //if exists we update database row
                            _soxContext.Update(rcmCta);
                        }
                        else
                        {
                            //create new item
                            _soxContext.Add(rcmCta);
                        }

                        await _soxContext.SaveChangesAsync();
                        context.Commit();
                    }
                }

                catch (Exception ex)
                {
                    await context.RollbackAsync();
                    FileLog.Write($"Error SaveWebhookRcmCta {ex}", "ErrorSaveWebhookRcmCta");
                }

            }

            result = true;

            return result;
        }

        public async Task<bool> DeleteRcmCtaAsync(int ItemId)
        {
            bool result = false;

            using (var context = _soxContext.Database.BeginTransaction())
            {
                try
                {
                    var rcmCtaCheck = _soxContext.RcmCta.FirstOrDefault(id => id.PodioItemId == (int)ItemId);

                    if (rcmCtaCheck != null)
                    {

                        _soxContext.Remove(rcmCtaCheck);

                        await _soxContext.SaveChangesAsync();
                        context.Commit();
                    }
                }

                catch (Exception ex)
                {
                    await context.RollbackAsync();
                    FileLog.Write($"Error DeleteRcmCtaAsync {ex}", "ErrorDeleteRcmCtaAsync");
                }

            }

            result = true;

            return result;
        }

        public async Task<bool> SavePodioRcmAsync(int ItemId)
        {
            bool result = false;

            using (var context = _soxContext.Database.BeginTransaction())
            {
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

                    if (podio.IsAuthenticated())
                    {

                        var podioItem = podio.ItemService.GetItem(ItemId);
                        var rawResponse = Newtonsoft.Json.JsonConvert.SerializeObject(podioItem.Result);
                        FileLog.Write($"{rawResponse}", "SavePodioRcmAsync");

                        Rcm rcm = new Rcm();

                        //check if item_id already exists in the database
                        bool isExists = false;
                        var rcmCheck = _soxContext.Rcm.FirstOrDefault(id => id.PodioItemId == (int)podioItem.Result.ItemId);
                        if (rcmCheck != null)
                        {
                            rcm = rcmCheck;
                            isExists = true;
                        }

                        #region Podio Item
                        rcm.PodioItemId = (int)podioItem.Result.ItemId;
                        rcm.PodioUniqueId = podioItem.Result.AppItemIdFormatted.ToString();
                        rcm.PodioRevision = podioItem.Result.CurrentRevision.Revision;
                        rcm.PodioLink = podioItem.Result.Link.ToString();
                        rcm.CreatedBy = podioItem.Result.CreatedBy.Name.ToString();
                        rcm.CreatedOn = DateTime.Parse(podioItem.Result.CreatedOn.ToString());

                        #region Rcm Fields

                        //Client Reference
                        AppItemField clientApp = podioItem.Result.Field<AppItemField>("client-reference");
                        IEnumerable<Item> clientAppRef = clientApp.Items;
                        rcm.ClientName = clientAppRef.Select(x => x.Title).FirstOrDefault();
                        rcm.ClientItemId = clientAppRef.Select(x => x.ItemId).FirstOrDefault();

                        var clientCheck = _soxContext.ClientSs.FirstOrDefault(id => id.PodioItemId == rcm.ClientItemId);
                        if (clientCheck != null)
                        {
                            rcm.ClientCode = clientCheck.ClientCode;
                        }

                        //Client Name
                        TextItemField txtClientName = podioItem.Result.Field<TextItemField>("title");
                        rcm.ClientNameText = txtClientName.Value;

                        //For what year is the RCM for?
                        TextItemField txtFy = podioItem.Result.Field<TextItemField>("text-3");
                        rcm.FY = txtFy.Value;

                        //Process
                        TextItemField txtProcess = podioItem.Result.Field<TextItemField>("text-2");
                        rcm.Process = txtProcess.Value;

                        //Sub-process
                        TextItemField txtSubProcess = podioItem.Result.Field<TextItemField>("text");
                        rcm.Subprocess = txtSubProcess.Value;

                        //Control Objective
                        TextItemField txtControlObj = podioItem.Result.Field<TextItemField>("control-objective");
                        rcm.ControlObjective = txtControlObj.Value;

                        //Specific Risk
                        TextItemField txtSpecRisk = podioItem.Result.Field<TextItemField>("specific-risk");
                        rcm.SpecificRisk = txtSpecRisk.Value;

                        //Financial Statement Elements
                        TextItemField txtFinStatement = podioItem.Result.Field<TextItemField>("financial-statement-elements");
                        rcm.FinStatemenElement = txtFinStatement.Value;

                        //Completeness & Accuracy
                        TextItemField txtCompleteness = podioItem.Result.Field<TextItemField>("completeness-accuracy");
                        rcm.CompletenessAccuracy = txtCompleteness.Value;
                        //rcm.CompletenessAccuracy = "Yes";
                        //Existence & Occurrence
                        TextItemField txtExistence = podioItem.Result.Field<TextItemField>("existence-occurrence");
                        rcm.ExistenceDisclosure = txtExistence.Value;
                        //rcm.ExistenceDisclosure = "Yes";
                        //Presentation & Disclosure
                        TextItemField txtPresentation = podioItem.Result.Field<TextItemField>("presentation-disclosure");
                        rcm.PresentationDisclosure = txtPresentation.Value;

                        //Rights & Obligations
                        TextItemField txtRightsObligation = podioItem.Result.Field<TextItemField>("rights-obligations");
                        rcm.RigthsObligation = txtRightsObligation.Value;

                        //Valuation & Allocation
                        TextItemField txtValuation = podioItem.Result.Field<TextItemField>("valuation-allocation");
                        rcm.ValuationAlloc = txtValuation.Value;

                        //Control ID
                        TextItemField txtControlId = podioItem.Result.Field<TextItemField>("control-id");
                        rcm.ControlId = txtControlId.Value;

                        //Control Activity FY19
                        TextItemField txtControlActFy19 = podioItem.Result.Field<TextItemField>("control-activity-fy19");
                        rcm.ControlActivityFy19 = txtControlActFy19.Value;

                        //Control In Place Date
                        TextItemField txtControlDate = podioItem.Result.Field<TextItemField>("control-in-place-date");
                        rcm.ControlPlaceDate = txtControlDate.Value;

                        //Control Owner
                        TextItemField txtControlOwner = podioItem.Result.Field<TextItemField>("control-owner");
                        rcm.ControlOwner = txtControlOwner.Value;

                        //Control Frequency
                        TextItemField txtControlFreq = podioItem.Result.Field<TextItemField>("control-frequency");
                        rcm.ControlFrequency = txtControlFreq.Value;

                        //Key/ Non-Key
                        TextItemField txtKey = podioItem.Result.Field<TextItemField>("key-non-key");
                        rcm.Key = txtKey.Value;

                        //Control Type
                        TextItemField txtControlType = podioItem.Result.Field<TextItemField>("control-type");
                        rcm.ControlType = txtControlType.Value;

                        //Nature of Procedure
                        TextItemField txtNature = podioItem.Result.Field<TextItemField>("nature-of-procedure");
                        rcm.NatureProc = txtNature.Value;

                        //Fraud Control
                        TextItemField txtFraudControl = podioItem.Result.Field<TextItemField>("fraud-control");
                        rcm.FraudControl = txtFraudControl.Value;

                        //Risk Level
                        TextItemField txtRisk = podioItem.Result.Field<TextItemField>("risk-level");
                        rcm.RiskLvl = txtRisk.Value;

                        //Management Review Control
                        TextItemField txtMgmtControl = podioItem.Result.Field<TextItemField>("management-review-control");
                        rcm.ManagementRevControl = txtMgmtControl.Value;

                        //PBC List
                        TextItemField txtPbcList = podioItem.Result.Field<TextItemField>("pbc-list");
                        rcm.PbcList = txtPbcList.Value;

                        //Test procedures
                        TextItemField txtTestProc = podioItem.Result.Field<TextItemField>("test-procedures");
                        rcm.TestProc = txtTestProc.Value;

                        #endregion

                        #endregion

                        if (isExists)
                        {
                            //if exists we update database row
                            _soxContext.Update(rcm);
                        }
                        else
                        {
                            //create new item
                            _soxContext.Add(rcm);
                        }

                        await _soxContext.SaveChangesAsync();
                        context.Commit();
                    }
                }

                catch (Exception ex)
                {
                    await context.RollbackAsync();
                    FileLog.Write($"Error SaveWebhookRcm {ex}", "ErrorSaveWebhookRcm");
                }

            }

            result = true;

            return result;
        }

        public async Task<bool> DeleteRcmAsync(int ItemId)
        {
            bool result = false;

            using (var context = _soxContext.Database.BeginTransaction())
            {
                try
                {
                    var rcmCheck = _soxContext.Rcm.FirstOrDefault(id => id.PodioItemId == (int)ItemId);

                    if (rcmCheck != null)
                    {

                        _soxContext.Remove(rcmCheck);

                        await _soxContext.SaveChangesAsync();
                        context.Commit();
                    }
                }

                catch (Exception ex)
                {
                    await context.RollbackAsync();
                    FileLog.Write($"Error DeleteRcmCtaAsync {ex}", "ErrorDeleteRcmAsync");
                }

            }

            result = true;

            return result;
        }
    
        public async Task<Rcm> SavePodioRcm(Rcm rcm)
        {
            //bool status = false;
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
                    rcm.PodioItemId = 0; //initially set to 0
                    Item rcmItem = new Item();


                    #region Podio Fields

                    if (rcm.FY != string.Empty && rcm.FY != null)
                    {
                        int q1Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q1").Value);
                        var textQ1 = rcmItem.Field<TextItemField>(q1Field);
                        textQ1.Value = rcm.FY;
                    }

                    if (rcm.ClientName != string.Empty && rcm.ClientName != null)
                    {
                        var checkClientItemId = _soxContext.ClientSs.Where(x => x.Name.Equals(rcm.ClientName)).Select(x => x.ItemId).FirstOrDefault();
                        if (checkClientItemId != null)
                        {
                            int q2Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q2").Value);
                            var appReference = rcmItem.Field<AppItemField>(q2Field);
                            List<int> listRoundItem = new List<int>();
                            listRoundItem.Add(checkClientItemId.Value);
                            appReference.ItemIds = listRoundItem;
                        }
                    }

                    if (rcm.Process != string.Empty && rcm.Process != null)
                    {
                        int q3Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q3").Value);
                        var textQ3 = rcmItem.Field<TextItemField>(q3Field);
                        textQ3.Value = rcm.Process;
                    }



                    if (rcm.Subprocess != string.Empty && rcm.Subprocess != null)
                    {
                        int q4Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q4").Value);
                        var textQ4 = rcmItem.Field<TextItemField>(q4Field);
                        textQ4.Value = rcm.Subprocess;
                    }



                    if (rcm.ControlObjective != string.Empty && rcm.ControlObjective != null)
                    {
                        int q5Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q5").Value);
                        var textQ5 = rcmItem.Field<TextItemField>(q5Field);
                        textQ5.Value = rcm.ControlObjective;
                    }



                    if (rcm.SpecificRisk != string.Empty && rcm.SpecificRisk != null)
                    {
                        int q6Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q6").Value);
                        var textQ6 = rcmItem.Field<TextItemField>(q6Field);
                        textQ6.Value = rcm.SpecificRisk;
                    }



                    if (rcm.FinStatemenElement != string.Empty && rcm.FinStatemenElement != null)
                    {
                        int q7Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q7").Value);
                        var textQ7 = rcmItem.Field<TextItemField>(q7Field);
                        textQ7.Value = rcm.FinStatemenElement;
                    }


                    if (rcm.FinancialStatementAssert != string.Empty && rcm.FinancialStatementAssert != null)
                    {
                        int q8Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q8").Value);
                        var textQ8 = rcmItem.Field<TextItemField>(q8Field);
                        if(rcm.FinancialStatementAssert.ToLower().Contains("accuracy"))
                        {
                            textQ8.Value = "Yes";
                            rcm.CompletenessAccuracy= "Yes";
                        }
                        else
                            textQ8.Value = "No";
                            rcm.CompletenessAccuracy = "No";
                    }

                    if (rcm.FinancialStatementAssert != string.Empty && rcm.FinancialStatementAssert != null)
                    {
                        int q8ExistenceField = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q8Existence").Value);
                        var textQ8Existence=rcmItem.Field<TextItemField>(q8ExistenceField);
                        if (rcm.FinancialStatementAssert.ToLower().Contains("existence"))
                        {
                            textQ8Existence.Value = "Yes";
                            rcm.ExistenceDisclosure = "Yes";
                        }
                        else
                            textQ8Existence.Value = "No";
                            rcm.ExistenceDisclosure = "No";
                    }

                    if (rcm.FinancialStatementAssert != string.Empty && rcm.FinancialStatementAssert != null)
                    {
                        int q8PresentationField = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q8Presentation").Value);
                        var textQ8Presentation = rcmItem.Field<TextItemField>(q8PresentationField);
                        if (rcm.FinancialStatementAssert.ToLower().Contains("presentation"))
                        {
                            textQ8Presentation.Value = "Yes";
                            rcm.PresentationDisclosure = "Yes";
                        }
                        else
                            textQ8Presentation.Value = "No";
                            rcm.PresentationDisclosure = "No";
                    }

                    if (rcm.FinancialStatementAssert != string.Empty && rcm.FinancialStatementAssert != null)
                    {
                        int q8RightsField = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q8Rights").Value);
                        var textQ8Rights = rcmItem.Field<TextItemField>(q8RightsField);
                        if (rcm.FinancialStatementAssert.ToLower().Contains("rights"))
                        {
                            textQ8Rights.Value = "Yes";
                            rcm.RigthsObligation = "Yes";
                        }
                        else
                            textQ8Rights.Value = "No";
                        rcm.RigthsObligation = "No";
                    }

                    if (rcm.FinancialStatementAssert != string.Empty && rcm.FinancialStatementAssert != null)
                    {
                        int q8ValuationField = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q8Valuation").Value);
                        var textQ8Valuation = rcmItem.Field<TextItemField>(q8ValuationField);
                        if (rcm.FinancialStatementAssert.ToLower().Contains("valuation"))
                        {
                            textQ8Valuation.Value = "Yes";
                            rcm.ValuationAlloc = "Yes";
                        }
                        else
                            textQ8Valuation.Value = "No";
                        rcm.ValuationAlloc = "No";
                    }



                    if (rcm.ControlId != string.Empty && rcm.ControlId != null)
                    {
                        int q9Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q9").Value);
                        var textQ9 = rcmItem.Field<TextItemField>(q9Field);
                        textQ9.Value = rcm.ControlId;
                    }



                    if (rcm.ControlActivity != string.Empty && rcm.ControlActivity != null)
                    {
                        int q10Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q10").Value);
                        var textQ10 = rcmItem.Field<TextItemField>(q10Field);
                        textQ10.Value = rcm.ControlActivity;
                    }

                    if (rcm.ShortDescription != string.Empty && rcm.ShortDescription != null)
                    {
                        int shortDescField = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("ShortDesc").Value);
                        var txtShortDesc = rcmItem.Field<TextItemField>(shortDescField);
                        txtShortDesc.Value = rcm.ShortDescription;
                    }


                    if (rcm.ControlPlaceDate != string.Empty && rcm.ControlPlaceDate != null)
                    {
                        int q11Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q11").Value);
                        var textQ11 = rcmItem.Field<TextItemField>(q11Field);
                        textQ11.Value = rcm.ControlPlaceDate;
                    }



                    if (rcm.ControlOwner != string.Empty && rcm.ControlOwner != null)
                    {
                        int q12Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q12").Value);
                        var textQ12 = rcmItem.Field<TextItemField>(q12Field);
                        textQ12.Value = rcm.ControlOwner;
                    }



                    if (rcm.ControlFrequency != string.Empty && rcm.ControlFrequency != null)
                    {
                        int q13Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q13").Value);
                        var textQ13 = rcmItem.Field<TextItemField>(q13Field);
                        textQ13.Value = rcm.ControlFrequency;
                    }



                    if (rcm.Key != string.Empty && rcm.Key != null)
                    {
                        int q14Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q14").Value);
                        var textQ14 = rcmItem.Field<TextItemField>(q14Field);
                        textQ14.Value = rcm.Key;
                    }



                    if (rcm.ControlType != string.Empty && rcm.ControlType != null)
                    {
                        int q15Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q15").Value);
                        var textQ15 = rcmItem.Field<TextItemField>(q15Field);
                        textQ15.Value = rcm.ControlType;
                    }



                    if (rcm.NatureProc != string.Empty && rcm.NatureProc != null)
                    {
                        int q16Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q16").Value);
                        var textQ16 = rcmItem.Field<TextItemField>(q16Field);
                        textQ16.Value = rcm.NatureProc;
                    }



                    if (rcm.FraudControl != string.Empty && rcm.FraudControl != null)
                    {
                        int q17Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q17").Value);
                        var textQ17 = rcmItem.Field<TextItemField>(q17Field);
                        textQ17.Value = rcm.FraudControl;
                    }



                    if (rcm.RiskLvl != string.Empty && rcm.RiskLvl != null)
                    {
                        int q18Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q18").Value);
                        var textQ18 = rcmItem.Field<TextItemField>(q18Field);
                        textQ18.Value = rcm.RiskLvl;
                    }



                    if (rcm.ManagementRevControl != string.Empty && rcm.ManagementRevControl != null)
                    {
                        int q19Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q19").Value);
                        var textQ19 = rcmItem.Field<TextItemField>(q19Field);
                        textQ19.Value = rcm.ManagementRevControl;
                    }



                    if (rcm.PbcList != string.Empty && rcm.PbcList != null)
                    {
                        int q20Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q20").Value);
                        var text20 = rcmItem.Field<TextItemField>(q20Field);
                        text20.Value = rcm.PbcList;
                    }



                    if (rcm.TestProc != string.Empty && rcm.TestProc != null)
                    {
                        int q21Field = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Q21").Value);
                        var text21 = rcmItem.Field<TextItemField>(q21Field);
                        text21.Value = rcm.TestProc;
                    }



                    int statusField = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Status").Value);
                    var categoryStatus = rcmItem.Field<CategoryItemField>(statusField);
                    categoryStatus.OptionText = "Active";

                    if (rcm.Duration != null && rcm.Duration.HasValue)
                    {
                        int durationField = int.Parse(_config.GetSection("RcmPodioApp").GetSection("RcmAppField").GetSection("Duration").Value);
                        DurationItemField durationItem = rcmItem.Field<DurationItemField>(durationField);
                        durationItem.Value = rcm.Duration;
                    }


                    #endregion

                    var roundId = await podio.ItemService.AddNewItem(Int32.Parse(PodioAppKey.AppId), rcmItem);

                    rcm.PodioItemId = int.Parse(roundId.ToString());
                    //status = true;

                   


                }
            }

            return rcm;
        }

        public async Task<bool> SaveRcmToDatabase(Rcm rcm)
        {
            //SoxTracker _soxtracker = new SoxTracker();
            try
            {
                using (var context = _soxContext.Database.BeginTransaction())
                {
                    //QuestionnaireRoundSet _roundSet = new QuestionnaireRoundSet();
                    //_roundSet = roundSet;

                    //bool isExists = false;

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

                    if (podio.IsAuthenticated() && rcm != null)
                    {
                        var podioItem = podio.ItemService.GetItem(rcm.PodioItemId);
                        var rawResponse = Newtonsoft.Json.JsonConvert.SerializeObject(podioItem.Result);
                        FileLog.Write($"{rawResponse}", "SavePodioRcmAsync");
                        rcm.PodioItemId = (int)podioItem.Result.ItemId;
                        rcm.PodioUniqueId = podioItem.Result.AppItemIdFormatted.ToString();
                        rcm.PodioRevision = podioItem.Result.CurrentRevision.Revision;
                        rcm.PodioLink = podioItem.Result.Link.ToString();
                        rcm.CreatedBy = podioItem.Result.CreatedBy.Name.ToString();
                        rcm.CreatedOn = DateTime.Parse(podioItem.Result.CreatedOn.ToString());
                    }
                    rcm.Status = "Active";

                    //validate created date
                    if (DateTime.TryParse(rcm.CreatedOn.Value.DateTime.ToString(), out DateTime dtCreated) && rcm.CreatedOn.Value.DateTime.ToString() != "01/01/0001 00:00:00")
                    {
                        rcm.CreatedOn = dtCreated;
                    }
                    else
                    {
                        rcm.CreatedOn = DateTime.Now;
                    }


                    //Get client code and client item id
                    var clientCheck = _soxContext.ClientSs.Where(x => x.ClientName.Equals(rcm.ClientName)).FirstOrDefault();
                    if (clientCheck != null)
                    {
                        rcm.ClientCode = clientCheck.ClientCode;
                        rcm.ClientItemId = clientCheck.ClientItemId;
                    }

                    //check for previous entry, we do upsert
                    var rcmCheck = _soxContext.Rcm.AsNoTracking().FirstOrDefault(id => id.PodioItemId.Equals(rcm.PodioItemId));
                    if (rcmCheck != null)
                    {
                        //rcm already exists and needs to update
                        rcm.Id = rcmCheck.Id;
                        _soxContext.Entry(rcmCheck).CurrentValues.SetValues(rcm);
                        //_soxtracker = _soxContext.SoxTracker
                        //           .Where(x => x.ClientName.Equals(rcm.ClientName)
                        //               && x.FY.Equals(rcm.FY)
                        //               && x.ControlId.Equals(rcm.ControlId))
                        //           .FirstOrDefault();
                        //if(_soxtracker != null)
                        //{
                        //    _soxtracker.ControlOwner = rcmCheck.ControlOwner;
                        //    _soxtracker.ControlActivity = rcmCheck.ControlActivity;
                        //    _soxtracker.TestProc = rcmCheck.TestProc;

                        //    _soxContext.Update(_soxtracker);
                        //}
                        await _soxContext.SaveChangesAsync();
                        context.Commit();
                    }
                    else
                    {
                        //rcm is new and needs to be added
                        _soxContext.Add(rcm);
                        await _soxContext.SaveChangesAsync();
                        context.Commit();
                    }
                }
                //Console.WriteLine(rcm.ControlId);
                return true;
            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorSaveRcmToDatabase");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SaveRcmToDatabase");
                return false;
            }
            
        }
        public async Task<bool> CheckDuplicateRCMDB(Rcm rcm)
        {
            //SoxTracker _soxtracker = new SoxTracker();
            try
            {
                using (var context = _soxContext.Database.BeginTransaction())
                {
                    //check for previous entry, we do upsert
                    var rcmCheck = _soxContext.Rcm.Where(x => x.ControlId.Equals(rcm.ControlId)
                                                   && x.Process.Equals(rcm.Process)
                                                   && x.Subprocess.Equals(rcm.Subprocess)
                                                   && x.FY.Equals(rcm.FY)
                                                   && x.ClientName.Equals(rcm.ClientName)).FirstOrDefault();
                    if (rcmCheck != null)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
               
            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorCheckingDuplicatesRCM");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "ErrorCheckingDuplicatesRCM");
                return false;
            }

        }
    }
}
