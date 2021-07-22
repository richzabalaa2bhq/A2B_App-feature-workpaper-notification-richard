using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace A2B_App.Server.Models
{
    public class QuestionnaireApp
    {
        List<ClientApp> ClientApp { get; set; }
    }

    public class ClientApp
    {
        public string AppId { get; set; }
        public string AppToken { get; set; }
    }

}
