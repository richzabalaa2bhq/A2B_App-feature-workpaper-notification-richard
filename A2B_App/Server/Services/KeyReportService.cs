using A2B_App.Server.Data;
using AutoMapper.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace A2B_App.Server.Services
{
    public class KeyReportService
    {
        private SoxContext _soxContext;
        private readonly IConfiguration _config;

        public KeyReportService(SoxContext soxContext, IConfiguration config)
        {
            _soxContext = soxContext;
            _config = config;
        }
    }

}
