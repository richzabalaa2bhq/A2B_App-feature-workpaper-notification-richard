using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text;
using System.Text.Json.Serialization;

namespace A2B_App.Shared.Time
{
    public class t_Employee
    {

    }

    public class A2BPodioUser
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int f_empid { get; set; }
        public string f_uniqueid { get; set; }
        public string f_name { get; set; }
        public string f_org { get; set; }
        public string f_podio_userid { get; set; }
        public string f_podio_profileid { get; set; }
        public string f_email { get; set; }
        public string f_itemid { get; set; }
        public string f_emplink { get; set; }
    }
}
