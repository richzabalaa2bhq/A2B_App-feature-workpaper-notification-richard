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
    public class SampleSelectionService
    {
        private SoxContext _soxContext;
        //private PodioField _podioField;
        private readonly IConfiguration _config;
        public SampleSelectionService(SoxContext soxContext, IConfiguration config)
        {
            _soxContext = soxContext;
            _config = config;
        }
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
                FileLog.Write($"Error SaveWebhookSampleSelection {ex}", "ErrorSaveWebhookSampleSelection");
            }

            return result;
        }

        public async Task<bool> SavePodioSampleSelectionAsync(int ItemId)
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
                    PodioAppKey.AppId = _config.GetSection("SampleSelectionPodioApp").GetSection("AppId").Value;
                    PodioAppKey.AppToken = _config.GetSection("SampleSelectionPodioApp").GetSection("AppToken").Value;

                    var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                    ServicePointManager.Expect100Continue = true;
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                    await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                    if (podio.IsAuthenticated())
                    {

                        var podioItem = podio.ItemService.GetItem(ItemId);
                        var rawResponse = Newtonsoft.Json.JsonConvert.SerializeObject(podioItem.Result);
                        FileLog.Write($"{rawResponse}", "SaveSampleSelectionAsync");

                        SampleSelection sampleSelection = new SampleSelection();

                        //check if item_id already exists in the database
                        bool isExists = false;
                        var sampleSelectionCheck = _soxContext.SampleSelection.FirstOrDefault(id => id.PodioItemId == (int)podioItem.Result.ItemId);
                        if (sampleSelectionCheck != null)
                        {
                            sampleSelection = sampleSelectionCheck;
                            isExists = true;
                        }

                        #region Podio Item
                        sampleSelection.PodioItemId = (int)podioItem.Result.ItemId;
                        sampleSelection.PodioUniqueId = podioItem.Result.AppItemIdFormatted.ToString();
                        sampleSelection.PodioRevision = podioItem.Result.CurrentRevision.Revision;
                        sampleSelection.PodioLink = podioItem.Result.Link.ToString();
                        sampleSelection.CreatedBy = podioItem.Result.CreatedBy.Name.ToString();
                        sampleSelection.CreatedOn = DateTime.Parse(podioItem.Result.CreatedOn.ToString());

                        #region sample selection fields

                        //client reference field
                        AppItemField clientApp = podioItem.Result.Field<AppItemField>("client");
                        IEnumerable<Item> clientAppRef = clientApp.Items;
                        sampleSelection.ClientName = clientAppRef.Select(x => x.Title).FirstOrDefault();
                        sampleSelection.ClientId = clientAppRef.Select(x => x.ItemId).FirstOrDefault();

                        //external auditor field
                        TextItemField calcExtAuditor = podioItem.Result.Field<TextItemField>("external-auditor");
                        sampleSelection.ExternalAuditor = calcExtAuditor.Value;

                        //Q4 (R3) Sample Required? field
                        CategoryItemField catSampleRequired = podioItem.Result.Field<CategoryItemField>("q4-r3-sample-required");
                        IEnumerable<CategoryItemField.Option> listCatSampleRequired = catSampleRequired.Options;
                        sampleSelection.Q4R3SampleRequired = listCatSampleRequired.Select(x => x.Text).FirstOrDefault();

                        //How many samples to be tested in Q4? field
                        TextItemField txtQ4Required = podioItem.Result.Field<TextItemField>("how-many-samples-to-be-tested-in-q4");
                        sampleSelection.CountSampleQ4R3 = (txtQ4Required.Value != null ? int.Parse(txtQ4Required.Value) : 0);

                        //Risk field
                        CategoryItemField catRisk = podioItem.Result.Field<CategoryItemField>("risk");
                        IEnumerable<CategoryItemField.Option> listcatRisk = catRisk.Options;
                        sampleSelection.Risk = listcatRisk.Select(x => x.Text).FirstOrDefault();

                        //Annual Population field
                        TextItemField txtAnnualPop = podioItem.Result.Field<TextItemField>("annual-population");
                        sampleSelection.AnnualPopulation = (txtAnnualPop.Value != null ? int.Parse(txtAnnualPop.Value) : 0);

                        //Annual Sample Size field
                        TextItemField txtAnnualSampleSize = podioItem.Result.Field<TextItemField>("annual-sample-size");
                        sampleSelection.AnnualSampleSize = (txtAnnualSampleSize.Value != null ? int.Parse(txtAnnualSampleSize.Value) : 0);

                        //Frequency field
                        CategoryItemField catFrequency = podioItem.Result.Field<CategoryItemField>("frequency");
                        IEnumerable<CategoryItemField.Option> listCatFrequency = catFrequency.Options;
                        sampleSelection.Frequency = listCatFrequency.Select(x => x.Text).FirstOrDefault();

                        //is Materiality Required field
                        CategoryItemField catMatReq = podioItem.Result.Field<CategoryItemField>("is-materiality-required");
                        IEnumerable<CategoryItemField.Option> listcatMatReq = catMatReq.Options;
                        sampleSelection.IsMateriality = listcatMatReq.Select(x => x.Text).FirstOrDefault();

                        //Materiality to Consider Round 1 field
                        TextItemField txtConsiderMat1 = podioItem.Result.Field<TextItemField>("materiality-to-consider-round-1");
                        sampleSelection.ConsiderMateriality1 = (txtConsiderMat1.Value != null ? txtConsiderMat1.Value.ToString() : string.Empty);

                        //Materiality to Consider Round 2 field
                        TextItemField txtConsiderMat2 = podioItem.Result.Field<TextItemField>("materiality-to-consider-round-2");
                        sampleSelection.ConsiderMateriality2 = (txtConsiderMat2.Value != null ? txtConsiderMat2.Value.ToString() : string.Empty);

                        //Materiality to Consider Round 3 field
                        TextItemField txtConsiderMat3 = podioItem.Result.Field<TextItemField>("materiality-to-consider-round-3");
                        sampleSelection.ConsiderMateriality3 = (txtConsiderMat3.Value != null ? txtConsiderMat3.Value.ToString() : string.Empty);


                        #region Round 1

                        //Population by Start Date Round 1 field
                        DateItemField dtRound1Start = podioItem.Result.Field<DateItemField>("round-1");
                        sampleSelection.Round1Start = dtRound1Start.Start; //Nullable DateTime

                        //Population by Start Date Round 1 field
                        DateItemField dtRound1End = podioItem.Result.Field<DateItemField>("round-2");
                        sampleSelection.Round1End = dtRound1End.Start; //Nullable DateTime

                        //Population by Round 1
                        NumericItemField numPopRound1 = podioItem.Result.Field<NumericItemField>("population-by-round-1");
                        sampleSelection.PopulationByRound1 = numPopRound1.Value != null ? (int)numPopRound1.Value : 0;

                        //Samples by Round 1
                        NumericItemField numSamplesbyRound1 = podioItem.Result.Field<NumericItemField>("samples-by-round-1");
                        sampleSelection.SamplesByRound1 = numSamplesbyRound1.Value != null ? (int)numSamplesbyRound1.Value : 0;

                        //Samples Closed Round 1
                        NumericItemField numSamplesClosedRound1 = podioItem.Result.Field<NumericItemField>("samples-closed-round-1");
                        sampleSelection.SamplesCloseRound1 = numSamplesClosedRound1.Value != null ? (int)numSamplesClosedRound1.Value : 0;

                        //Samples Remaining Round 1
                        NumericItemField numSamplesRemainRound1 = podioItem.Result.Field<NumericItemField>("samples-remaining-round-1");
                        sampleSelection.SamplesRemainingRound1 = numSamplesRemainRound1.Value != null ? (int)numSamplesRemainRound1.Value : 0;

                        #endregion

                        #region Round 2
                        //Population by Start Date Round 2 field
                        DateItemField dtRound2Start = podioItem.Result.Field<DateItemField>("round-1-2");
                        sampleSelection.Round2Start = dtRound2Start.Start; //Nullable DateTime

                        //Population by Start Date Round 2 field
                        DateItemField dtRound2End = podioItem.Result.Field<DateItemField>("round-2-2");
                        sampleSelection.Round2End = dtRound2End.Start; //Nullable DateTime

                        //Population by Round 2
                        NumericItemField numPopRound2 = podioItem.Result.Field<NumericItemField>("population-by-round-2");
                        sampleSelection.PopulationByRound2 = numPopRound2.Value != null ? (int)numPopRound2.Value : 0;

                        //Samples by Round 2
                        NumericItemField numSamplesbyRound2 = podioItem.Result.Field<NumericItemField>("samples-by-round-2");
                        sampleSelection.SamplesByRound2 = numSamplesbyRound2.Value != null ? (int)numSamplesbyRound2.Value : 0;

                        //Samples Closed Round 2
                        NumericItemField numSamplesClosedRound2 = podioItem.Result.Field<NumericItemField>("samples-closed-round-2");
                        sampleSelection.SamplesCloseRound2 = numSamplesClosedRound2.Value != null ? (int)numSamplesClosedRound2.Value : 0;

                        //Samples Remaining Round 2
                        NumericItemField numSamplesRemainRound2 = podioItem.Result.Field<NumericItemField>("samples-remaining-round-2");
                        sampleSelection.SamplesRemainingRound2 = numSamplesRemainRound2.Value != null ? (int)numSamplesRemainRound2.Value : 0;

                        #endregion

                        #region Round 3
                        //Population by Start Date Round 3 field
                        DateItemField dtRound3Start = podioItem.Result.Field<DateItemField>("population-by-start-date-round-3");
                        sampleSelection.Round3Start = dtRound3Start.Start; //Nullable DateTime

                        //Population by Start Date Round 3 field
                        DateItemField dtRound3End = podioItem.Result.Field<DateItemField>("population-by-end-date-round-3");
                        sampleSelection.Round3End = dtRound3End.Start; //Nullable DateTime

                        //Population by Round 3
                        NumericItemField numPopRound3 = podioItem.Result.Field<NumericItemField>("population-by-round-3");
                        sampleSelection.PopulationByRound3 = numPopRound3.Value != null ? (int)numPopRound3.Value : 0;

                        //Samples by Round 3
                        NumericItemField numSamplesbyRound3 = podioItem.Result.Field<NumericItemField>("samples-by-round-3");
                        sampleSelection.SamplesByRound3 = numSamplesbyRound3.Value != null ? (int)numSamplesbyRound3.Value : 0;

                        //Samples Closed Round 3
                        NumericItemField numSamplesClosedRound3 = podioItem.Result.Field<NumericItemField>("samples-closed-round-3");
                        sampleSelection.SamplesCloseRound3 = numSamplesClosedRound3.Value != null ? (int)numSamplesClosedRound3.Value : 0;

                        //Samples Remaining Round 3
                        NumericItemField numSamplesRemainRound3 = podioItem.Result.Field<NumericItemField>("samples-remaining-round-3");
                        sampleSelection.SamplesRemainingRound3 = numSamplesRemainRound3.Value != null ? (int)numSamplesRemainRound3.Value : 0;

                        #endregion


                        //Population by Round
                        CalculationItemField calcPopRoundTotal = podioItem.Result.Field<CalculationItemField>("population-by-round");
                        sampleSelection.PopulationByRoundTot = calcPopRoundTotal.Value != null ? (int)calcPopRoundTotal.Value : 0;

                        //Samples by Round
                        CalculationItemField calcSamplesRoundTotal = podioItem.Result.Field<CalculationItemField>("samples-by-round");
                        sampleSelection.SamplesByRoundTot = calcSamplesRoundTotal.Value != null ? (int)calcSamplesRoundTotal.Value : 0;

                        //Samples Closed
                        CalculationItemField calcClosedRoundTotal = podioItem.Result.Field<CalculationItemField>("samples-closed");
                        sampleSelection.SamplesCloseRoundTot = calcClosedRoundTotal.Value != null ? (int)calcClosedRoundTotal.Value : 0;

                        //Samples Remaining
                        CalculationItemField calcRemainingRoundTotal = podioItem.Result.Field<CalculationItemField>("samples-remaining");
                        sampleSelection.SamplesRemainingRoundTot = calcRemainingRoundTotal.Value != null ? (int)calcRemainingRoundTotal.Value : 0;

                        //Sharefile Link
                        EmbedItemField embedField = podioItem.Result.Field<EmbedItemField>("sharefile-link");
                        IEnumerable<Embed> embeds = embedField.Embeds;
                        if (embeds.Count() > 0)
                        {
                            //Get url of the first link
                            sampleSelection.SharefileLink = embeds.FirstOrDefault().OriginalUrl;
                        }


                        //Population Filename field
                        TextItemField txtPopulationFilename = podioItem.Result.Field<TextItemField>("population-guid-name");
                        sampleSelection.PopulationFile = (txtPopulationFilename.Value != null ? txtPopulationFilename.Value : string.Empty);

                        //Json Data field
                        TextItemField txtJsonData = podioItem.Result.Field<TextItemField>("json-data");
                        sampleSelection.JsonData = (txtJsonData.Value != null ? txtJsonData.Value : string.Empty);

                        //Version field
                        TextItemField txtVersion = podioItem.Result.Field<TextItemField>("version");
                        sampleSelection.Version = (txtVersion.Value != null ? txtVersion.Value : string.Empty);

                        //RCM Item Id field
                        TextItemField txtRcmItemId = podioItem.Result.Field<TextItemField>("rcm-item-id");
                        sampleSelection.RcmPodioItemId = (txtRcmItemId.Value != null ? int.Parse(txtRcmItemId.Value) : 0);

                        //client reference field
                        AppItemField testRoundApp = podioItem.Result.Field<AppItemField>("round-test");
                        IEnumerable<Item> testRoundAppRef = testRoundApp.Items;

                        //List<TestRoundCreatedItemId> listCreatedItem = new List<TestRoundCreatedItemId>();
                        //if (testRoundAppRef.Count() > 0)
                        //{
                        //    foreach (var item in testRoundAppRef)
                        //    {

                        //        TestRoundCreatedItemId createdItemId = new TestRoundCreatedItemId();
                        //        createdItemId.PodioItemID = item.ItemId;
                        //        createdItemId.SampleSelectionData = sampleSelection;
                        //        listCreatedItem.Add(createdItemId);
                        //        //sampleSelection.ListRefId.Add(createdItemId);

                        //    }
                        //}



                        #endregion

                        #endregion

                        if (isExists)
                        {
                            //if exists we update database row
                            _soxContext.Update(sampleSelection);
                        }
                        else
                        {
                            //create new item
                            _soxContext.Add(sampleSelection);
                        }

                        List<TestRoundSampleSelectionReference> listCreatedItem = new List<TestRoundSampleSelectionReference>();
                        if (testRoundAppRef.Count() > 0)
                        {
                            foreach (var item in testRoundAppRef)
                            {
                                var testRoundCheck = _soxContext.TestRoundSampleSelectionReference.FirstOrDefault(c => c.PodioItemID == item.ItemId);
                                if (testRoundCheck == null)
                                {
                                    TestRoundSampleSelectionReference createdItemId = new TestRoundSampleSelectionReference();
                                    createdItemId.PodioItemID = item.ItemId;
                                    createdItemId.SampleSelectionData = sampleSelection;
                                    listCreatedItem.Add(createdItemId);
                                }
                            }
                        }

                        if (listCreatedItem != null && listCreatedItem.Count > 0)
                        {
                            _soxContext.AddRange(listCreatedItem);
                        }

                        await _soxContext.SaveChangesAsync();
                        context.Commit();
                    }
                }

                catch (Exception ex)
                {
                    await context.RollbackAsync();
                    FileLog.Write($"Error SaveSampleSelectionAsync {ex}", "ErrorSaveSampleSelectionAsync");
                }

            }

            result = true;

            return result;
        }

        public async Task<bool> DeleteSampleSelectionAsync(int ItemId)
        {
            bool result = false;

            using (var context = _soxContext.Database.BeginTransaction())
            {
                try
                {
                    var sampleSelectionCheck = _soxContext.SampleSelection.FirstOrDefault(id => id.PodioItemId == (int)ItemId);

                    if (sampleSelectionCheck != null)
                    {

                        _soxContext.Remove(sampleSelectionCheck);


                        var testRoundSampleSelectionRefCheck = _soxContext.TestRoundSampleSelectionReference.Where((System.Linq.Expressions.Expression<Func<TestRoundSampleSelectionReference, bool>>)(id => id.SampleSelectionData == sampleSelectionCheck));
                        if (testRoundSampleSelectionRefCheck != null)
                        {
                            _soxContext.RemoveRange(testRoundSampleSelectionRefCheck);
                        }

                        await _soxContext.SaveChangesAsync();
                        context.Commit();
                    }



                }

                catch (Exception ex)
                {
                    await context.RollbackAsync();
                    FileLog.Write($"Error DeleteSampleSelectionAsync {ex}", "ErrorDeleteSampleSelectionAsync");
                }

            }

            result = true;

            return result;
        }

        public async Task<bool> SavePodioTestRoundAsync(int ItemId)
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
                    PodioAppKey.AppId = _config.GetSection("SampleSelectionPodioApp").GetSection("RoundAppId").Value;
                    PodioAppKey.AppToken = _config.GetSection("SampleSelectionPodioApp").GetSection("RoundAppToken").Value;

                    var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                    ServicePointManager.Expect100Continue = true;
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                    await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                    if (podio.IsAuthenticated())
                    {

                        var podioItem = podio.ItemService.GetItem(ItemId);
                        var rawResponse = Newtonsoft.Json.JsonConvert.SerializeObject(podioItem.Result);
                        FileLog.Write($"{rawResponse}", "SaveTestRoundAsync");

                        TestRound testRound = new TestRound();

                        //check if item_id already exists in the database
                        bool isExists = false;
                        var testRoundCheck = _soxContext.TestRounds.FirstOrDefault(id => id.PodioItemId == (int)podioItem.Result.ItemId);
                        if (testRoundCheck != null)
                        {
                            testRound = testRoundCheck;
                            isExists = true;
                        }

                        #region Podio Item
                        testRound.PodioItemId = (int)podioItem.Result.ItemId;
                        testRound.PodioUniqueId = podioItem.Result.AppItemIdFormatted.ToString();
                        testRound.PodioRevision = podioItem.Result.CurrentRevision.Revision;
                        testRound.PodioLink = podioItem.Result.Link.ToString();
                        testRound.CreatedBy = podioItem.Result.CreatedBy.Name.ToString();
                        testRound.CreatedOn = DateTime.Parse(podioItem.Result.CreatedOn.ToString());

                        #region test round fields

                        //Testing Round field
                        TextItemField txtTitle = podioItem.Result.Field<TextItemField>("title");
                        testRound.TestingRound = txtTitle.Value;

                        //A2Q2 Samples field
                        TextItemField txtA2Q2Samples = podioItem.Result.Field<TextItemField>("a2q2-samples");
                        testRound.A2Q2Samples = txtA2Q2Samples.Value;

                        //Date field
                        DateItemField dtRoundStart = podioItem.Result.Field<DateItemField>("date");
                        testRound.Date = dtRoundStart.Start; //Nullable DateTime

                        //Status field
                        CategoryItemField catStatus = podioItem.Result.Field<CategoryItemField>("status");
                        IEnumerable<CategoryItemField.Option> listcatStatus = catStatus.Options;
                        testRound.Status = listcatStatus.Select(x => x.Text).FirstOrDefault();

                        //A2Q2 Comments field
                        TextItemField txtComments = podioItem.Result.Field<TextItemField>("a2q2-comments");
                        testRound.Comment = txtComments.Value;

                        //Header Display 1 field
                        TextItemField txtHeader1 = podioItem.Result.Field<TextItemField>("materiality-value-1");
                        testRound.HeaderRoundDisplay1 = txtHeader1.Value;

                        //Header Display 2 field
                        TextItemField txtHeader2 = podioItem.Result.Field<TextItemField>("materiality-value-2");
                        testRound.HeaderRoundDisplay2 = txtHeader2.Value;

                        //Header Display 3 field
                        TextItemField txtHeader3 = podioItem.Result.Field<TextItemField>("materiality-value-3");
                        testRound.HeaderRoundDisplay3 = txtHeader3.Value;

                        //Header Display 4 field
                        TextItemField txtHeader4 = podioItem.Result.Field<TextItemField>("header-display-4");
                        testRound.HeaderRoundDisplay4 = txtHeader4.Value;

                        //Header Display 5 field
                        TextItemField txtHeader5 = podioItem.Result.Field<TextItemField>("header-display-5");
                        testRound.HeaderRoundDisplay5 = txtHeader5.Value;

                        //Content Display 1 field
                        TextItemField txtContent1 = podioItem.Result.Field<TextItemField>("content-display-1");
                        testRound.ContentDisplay1 = txtContent1.Value;

                        //Content Display 2 field
                        TextItemField txtContent2 = podioItem.Result.Field<TextItemField>("content-display-2");
                        testRound.ContentDisplay2 = txtContent2.Value;

                        //Content Display 3 field
                        TextItemField txtContent3 = podioItem.Result.Field<TextItemField>("content-display-3");
                        testRound.ContentDisplay3 = txtContent3.Value;

                        //Content Display 4 field
                        TextItemField txtContent4 = podioItem.Result.Field<TextItemField>("content-display-4");
                        testRound.ContentDisplay4 = txtContent4.Value;

                        //Content Display 5 field
                        TextItemField txtContent5 = podioItem.Result.Field<TextItemField>("content-display-5");
                        testRound.ContentDisplay5 = txtContent5.Value;

                        #endregion

                        #endregion

                        if (isExists)
                        {
                            //if exists we update database row
                            _soxContext.Update(testRound);
                        }
                        else
                        {
                            //create new item
                            _soxContext.Add(testRound);
                        }

                        await _soxContext.SaveChangesAsync();
                        context.Commit();
                    }
                }

                catch (Exception ex)
                {
                    await context.RollbackAsync();
                    FileLog.Write($"Error SaveTestRoundAsync {ex}", "ErrorSaveTestRoundAsync");
                }

            }

            result = true;

            return result;
        }

        public async Task<bool> DeleteTestRoundAsync(int ItemId)
        {
            bool result = false;

            using (var context = _soxContext.Database.BeginTransaction())
            {
                try
                {
                    var sampleSelectionCheck = _soxContext.TestRounds.FirstOrDefault(id => id.PodioItemId == (int)ItemId);

                    if (sampleSelectionCheck != null)
                    {

                        _soxContext.Remove(sampleSelectionCheck);

                        await _soxContext.SaveChangesAsync();
                        context.Commit();
                    }

                }

                catch (Exception ex)
                {
                    await context.RollbackAsync();
                    FileLog.Write($"Error DeleteTestRoundAsync {ex}", "ErrorDeleteTestRoundAsync");
                }

            }

            result = true;

            return result;
        }

        public async Task<int> SaveSampleSelectionAsync(SampleSelection SampleSelection)
        {
            bool status = false;
            SampleSelection _sampleSelection = new SampleSelection();
            using (var context = _soxContext.Database.BeginTransaction())
            {

                try
                {

                    _sampleSelection = SampleSelection;

                    bool isExists = false;
                    var sampleSelectionCheck = _soxContext.SampleSelection.FirstOrDefault(id => id.PodioItemId == (int)SampleSelection.PodioItemId);
                    if (sampleSelectionCheck != null)
                    {
                        isExists = true;
                    }

                    if (isExists)
                    {
                        //_soxContext.Update(SampleSelection);
                        _sampleSelection.Id = sampleSelectionCheck.Id;
                        _soxContext.Entry(sampleSelectionCheck).CurrentValues.SetValues(_sampleSelection);
                        //_soxContext.Update(_sampleSelection);
                    }
                    else
                    {
                        _soxContext.Add(_sampleSelection);
                    }

                    List<TestRound> listTestRoundAdd = new List<TestRound>();
                    List<TestRound> listTestRoundUpdate = new List<TestRound>();
                    if (_sampleSelection.ListTestRound.Count > 0)
                    {
                        foreach (var item in _sampleSelection.ListTestRound)
                        {

                            var testRoundCheck = _soxContext.TestRounds.FirstOrDefault(c => c.PodioItemId == item.PodioItemId);
                            if (testRoundCheck == null)
                            {
                                TestRound createdItemId = new TestRound();
                                createdItemId = item;
                                createdItemId.SampleSelectionData = _sampleSelection;
                                listTestRoundAdd.Add(createdItemId);
                            }

                        }
                    }
                    if (listTestRoundAdd != null && listTestRoundAdd.Count > 0)
                    {
                        _soxContext.AddRange(listTestRoundAdd);
                    }

                    await _soxContext.SaveChangesAsync();
                    context.Commit();
                    status = true;
                }

                catch (Exception ex)
                {
                    await context.RollbackAsync();
                    status = false;
                    FileLog.Write($"Error SaveSampleSelectionAsync {ex}", "ErrorSaveSampleSelectionAsync");

                }

            }

            if (status)
            {
                return _sampleSelection.Id;
            }
            else
            {
                return 0;
            }
        }

        public async Task<bool> SavePodioClientSSAsync(int ItemId)
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
                    PodioAppKey.AppId = _config.GetSection("SampleSelectionPodioApp").GetSection("ClientAppId").Value;
                    PodioAppKey.AppToken = _config.GetSection("SampleSelectionPodioApp").GetSection("ClientAppToken").Value;

                    var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                    ServicePointManager.Expect100Continue = true;
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                    await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                    if (podio.IsAuthenticated())
                    {

                        var podioItem = podio.ItemService.GetItem(ItemId);
                        var rawResponse = Newtonsoft.Json.JsonConvert.SerializeObject(podioItem.Result);
                        FileLog.Write($"{rawResponse}", "SavePodioClientSSAsync");

                        ClientSs clientSs = new ClientSs();

                        //check if item_id already exists in the database
                        bool isExists = false;
                        var clientCheck = _soxContext.ClientSs.FirstOrDefault(id => id.PodioItemId == (int)podioItem.Result.ItemId);
                        if (clientCheck != null)
                        {
                            clientSs = clientCheck;
                            isExists = true;
                        }

                        #region Podio Item
                        clientSs.PodioItemId = (int)podioItem.Result.ItemId;
                        clientSs.PodioUniqueId = podioItem.Result.AppItemIdFormatted.ToString();
                        clientSs.PodioRevision = podioItem.Result.CurrentRevision.Revision;
                        clientSs.PodioLink = podioItem.Result.Link.ToString();
                        clientSs.CreatedBy = podioItem.Result.CreatedBy.Name.ToString();
                        clientSs.CreatedOn = DateTime.Parse(podioItem.Result.CreatedOn.ToString());
                        clientSs.ItemId = (int)podioItem.Result.ItemId;
                        #region Podio fields

                        //client reference field
                        AppItemField clientRef = podioItem.Result.Field<AppItemField>("client-relationship");
                        IEnumerable<Item> clientAppRef = clientRef.Items;
                        clientSs.ClientName = clientAppRef.Select(x => x.Title).FirstOrDefault();
                        //clientSs.ItemId = clientAppRef.Select(x => x.ItemId).FirstOrDefault();
                        clientSs.ClientItemId = clientAppRef.Select(x => x.ItemId).FirstOrDefault();

                        //Client Id
                        CalculationItemField calClientId = podioItem.Result.Field<CalculationItemField>("client-id");
                        clientSs.ClientCode = calClientId.ValueAsString;

                        //external auditor field
                        TextItemField calcExtAuditor = podioItem.Result.Field<TextItemField>("external-auditor");
                        clientSs.ExternalAuditor = calcExtAuditor.Value;

                        //Round 1 Percent
                        NumericItemField numRound1Percent = podioItem.Result.Field<NumericItemField>("round-1-percent");
                        clientSs.Percent = numRound1Percent.Value != null ? (int)numRound1Percent.Value : 0;

                        //Round 2 Percent
                        NumericItemField numRound2Percent = podioItem.Result.Field<NumericItemField>("round-2-percent");
                        clientSs.PercentRound2 = numRound2Percent.Value != null ? (int)numRound2Percent.Value : 0;

                        //external auditor field
                        TextItemField calcSharefileId = podioItem.Result.Field<TextItemField>("sharefile-id-save-file");
                        clientSs.SharefileId = calcSharefileId.Value;

                        #endregion

                        #endregion

                        if (isExists)
                        {
                            //if exists we update database row
                            _soxContext.Update(clientSs);
                        }
                        else
                        {
                            //create new item
                            _soxContext.Add(clientSs);
                        }

                        await _soxContext.SaveChangesAsync();
                        context.Commit();
                    }
                }

                catch (Exception ex)
                {
                    await context.RollbackAsync();
                    FileLog.Write($"Error SavePodioClientSSAsync {ex}", "ErrorSavePodioClientSSAsync");
                }

            }

            result = true;

            return result;
        }

        public async Task<bool> DeleteClientSSAsync(int ItemId)
        {
            bool result = false;

            using (var context = _soxContext.Database.BeginTransaction())
            {
                try
                {
                    var clientCheck = _soxContext.ClientSs.FirstOrDefault(id => id.PodioItemId == (int)ItemId);

                    if (clientCheck != null)
                    {

                        _soxContext.Remove(clientCheck);

                        await _soxContext.SaveChangesAsync();
                        context.Commit();
                    }

                }

                catch (Exception ex)
                {
                    await context.RollbackAsync();
                    FileLog.Write($"Error DeleteClientSSAsync {ex}", "ErrorDeleteClientSSAsync");
                }

            }

            result = true;

            return result;
        }

        public async Task<bool> SavePodioMatrixAsync(int ItemId)
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
                    PodioAppKey.AppId = _config.GetSection("SampleSelectionPodioApp").GetSection("MatrixAppId").Value;
                    PodioAppKey.AppToken = _config.GetSection("SampleSelectionPodioApp").GetSection("MatrixAppToken").Value;

                    var podio = new Podio(PodioKey.ClientId, PodioKey.ClientSecret);
                    ServicePointManager.Expect100Continue = true;
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                    await podio.AuthenticateWithApp(Int32.Parse(PodioAppKey.AppId), PodioAppKey.AppToken);

                    if (podio.IsAuthenticated())
                    {

                        var podioItem = podio.ItemService.GetItem(ItemId);
                        var rawResponse = Newtonsoft.Json.JsonConvert.SerializeObject(podioItem.Result);
                        FileLog.Write($"{rawResponse}", "SavePodioMatrixAsync");

                        Matrix matrix = new Matrix();

                        //check if item_id already exists in the database
                        bool isExists = false;
                        var matrixCheck = _soxContext.Matrix.FirstOrDefault(id => id.PodioItemId == (int)podioItem.Result.ItemId);
                        if (matrixCheck != null)
                        {
                            matrix = matrixCheck;
                            isExists = true;
                        }

                        #region Podio Item
                        matrix.PodioItemId = (int)podioItem.Result.ItemId;
                        matrix.PodioUniqueId = podioItem.Result.AppItemIdFormatted.ToString();
                        matrix.PodioRevision = podioItem.Result.CurrentRevision.Revision;
                        matrix.PodioLink = podioItem.Result.Link.ToString();
                        matrix.CreatedBy = podioItem.Result.CreatedBy.Name.ToString();
                        matrix.CreatedOn = DateTime.Parse(podioItem.Result.CreatedOn.ToString());

                        #region Podio fields

                        //client reference field
                        AppItemField clientRef = podioItem.Result.Field<AppItemField>("client");
                        IEnumerable<Item> clientAppRef = clientRef.Items;
                        matrix.ClientName = clientAppRef.Select(x => x.Title).FirstOrDefault();
                        matrix.ClientItemId = clientAppRef.Select(x => x.ItemId).FirstOrDefault();

                        var clientCheck = _soxContext.ClientSs.FirstOrDefault(id => id.PodioItemId == matrix.ClientItemId);
                        if (clientCheck != null)
                        {
                            matrix.ClientCode = clientCheck.ClientCode;
                        }

                        //External Auditor
                        CalculationItemField calcExternalAuditor = podioItem.Result.Field<CalculationItemField>("external-auditor");
                        matrix.ExternalAuditor = calcExternalAuditor.ValueAsString;

                        //Frequency
                        CategoryItemField catFrequency = podioItem.Result.Field<CategoryItemField>("frequency");
                        IEnumerable<CategoryItemField.Option> listCatFrequency = catFrequency.Options;
                        matrix.Frequency = listCatFrequency.Select(x => x.Text).FirstOrDefault();

                        //Risk
                        CategoryItemField catRisk = podioItem.Result.Field<CategoryItemField>("risk");
                        IEnumerable<CategoryItemField.Option> listCatRisk = catRisk.Options;
                        matrix.Risk = listCatRisk.Select(x => x.Text).FirstOrDefault();

                        //Value
                        NumericItemField numValue = podioItem.Result.Field<NumericItemField>("value");
                        matrix.SizeValue = numValue.Value != null ? (int)numValue.Value : 0;

                        //Population Start
                        NumericItemField numPopStart = podioItem.Result.Field<NumericItemField>("population-start");
                        matrix.StartPopulation = numPopStart.Value != null ? (int)numPopStart.Value : 0;

                        //Population End
                        NumericItemField numPopEnd = podioItem.Result.Field<NumericItemField>("population-end");
                        matrix.EndPopulation = numPopEnd.Value != null ? (int)numPopEnd.Value : 0;

                        #endregion

                        #endregion

                        if (isExists)
                        {
                            //if exists we update database row
                            _soxContext.Update(matrix);
                        }
                        else
                        {
                            //create new item
                            _soxContext.Add(matrix);
                        }

                        await _soxContext.SaveChangesAsync();
                        context.Commit();
                    }
                }

                catch (Exception ex)
                {
                    await context.RollbackAsync();
                    FileLog.Write($"Error SavePodioMatrixAsync {ex}", "ErrorSavePodioMatrixAsync");
                }

            }

            result = true;

            return result;
        }

        public async Task<bool> DeleteMatrixAsync(int ItemId)
        {
            bool result = false;

            using (var context = _soxContext.Database.BeginTransaction())
            {
                try
                {
                    var matrixCheck = _soxContext.Matrix.FirstOrDefault(id => id.PodioItemId == (int)ItemId);

                    if (matrixCheck != null)
                    {

                        _soxContext.Remove(matrixCheck);

                        await _soxContext.SaveChangesAsync();
                        context.Commit();
                    }

                }

                catch (Exception ex)
                {
                    await context.RollbackAsync();
                    FileLog.Write($"Error DeleteMatrixAsync {ex}", "ErrorDeleteMatrixAsync");
                }

            }

            result = true;

            return result;
        }

    }
}
