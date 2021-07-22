using System;
using System.Collections.Generic;
using System.Text;

namespace A2B_App.Shared.Email
{
    public class EmailParam
    {
        public string Subject { get; set; }
        public string Message { get; set; }
        public List<string> ListEmailTo { get; set; }
        public List<string> ListEmailCc { get; set; }
    }

    public class EmailNotificationBody
    {
        public string UserName { get; set; }
        public string Message { get; set; }
        public string Url { get; set; }
        public string UrlName { get; set; }
    }


}
