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
    public class SoxTrackerService
    {
        private SoxContext _soxContext;
        private readonly IConfiguration _config;

        public SoxTrackerService(SoxContext soxContext, IConfiguration config)
        {
            _soxContext = soxContext;
            _config = config;
        }

        public async Task<bool> SaveSoxTrackerToDatabase(Rcm rcm)
        {
            SoxTracker soxtracker = new SoxTracker();
            try
            {
                //Console.WriteLine("testetasdjh");
                using (var context = _soxContext.Database.BeginTransaction())
                {
                    //QuestionnaireRoundSet _roundSet = new QuestionnaireRoundSet();
                    //_roundSet = roundSet;

                    //bool isExists = false;

           

                    //check for previous entry, we do upsert
                    var soxTrackerCheck = _soxContext.SoxTracker.AsNoTracking().FirstOrDefault(id => id.ControlId.Equals(rcm.ControlId));
                    if (soxTrackerCheck != null)
                    {
                        //Console.WriteLine("true"+rcm.ControlId);
                        //rcm already exists and needs to update
                        soxtracker.PBC = rcm.PbcList;
                        _soxContext.Entry(soxTrackerCheck).CurrentValues.SetValues(soxtracker);
                     
                        await _soxContext.SaveChangesAsync();
                        context.Commit();
                    }
                    else
                    {
                        //Console.WriteLine("false");
                        return false;
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorSaveSoxTrackerToDatabase");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SaveSoxTrackerToDatabase");
                return false;
            }

        }
    }
}
