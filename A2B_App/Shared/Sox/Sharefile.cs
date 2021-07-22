using System;
using System.Collections.Generic;
using System.Text;

namespace A2B_App.Shared.Sox
{
    public class SharefileUser
    {
        public string Username { get; set; }
        public string Password { get; set; }
        public string Subdomain { get; set; }
        public string ControlPlane { get; set; }
    }

    public class SharefileItem
    {
        public string FileName { get; set; }
        public string FilePath { get; set; }
        public string Directory { get; set; }
    }
}
