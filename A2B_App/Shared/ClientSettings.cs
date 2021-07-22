using System;
using System.Collections.Generic;
using System.Text;

namespace A2B_App.Shared
{
    public class ClientSettings
    {
        //private readonly static string ApiServer = "http://localhost:8081/";
        ////public readonly static string ApiServer = "https://sjo5-api.ngrok.io/";

        //private readonly static string SodUri = "https://le.ngrok.io/sod";
        //private readonly static string RcmUri = "https://le.ngrok.io/rcm";
        //private readonly static string SoxTrackerUri = "https://le.ngrok.io/sox-tracker";

        public string ApiServer { get; set; }
        public string SodUri { get; set; }
        public string RcmUri { get; set; }
        public string SoxTrackerUri { get; set; }
        public string IUCSystemAppId { get; set; }
        public string IUCNonSystemAppId { get; set; }
        public SoxQuestionnaire SoxQuestionnaire { get; set; }
        public string ClientUI { get; set; }
        public string GetApiServer()
        {
            return ApiServer;
        }
        public string GetClientUi ()
        {
            return ClientUI;
        }
        public string GetSodUri()
        {
            return SodUri;
        }

        public string GetRcmUri()
        {
            return RcmUri;
        }

        public string GetSoxTrackerUri()
        {
            return SoxTrackerUri;
        }

        public string GetIUCSystemAppId()
        {
            return IUCSystemAppId;
        }

        public string GetIUCNonSystemAppId()
        {
            return IUCNonSystemAppId;
        }

    }


    public partial class SoxQuestionnaire
    {
        public Eri Eri { get; set; }
    }

    public partial class Eri
    {
        public string ELC10 { get; set; }
    }

}
