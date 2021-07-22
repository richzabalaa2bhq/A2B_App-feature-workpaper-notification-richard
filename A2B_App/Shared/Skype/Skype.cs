using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text.Json.Serialization;

namespace A2B_App.Shared.Skype
{
    public class Skype
    {
        public string Address { get; set; }
        public string Message { get; set; }
        public Mention Mention { get; set; }
    }

    public class Mention
    {
        public string Id { get; set; }
        public string Name { get; set; }
    }

    public class Conversation
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Sys_id { get; set; }
        public string id { get; set; }
        public bool isgroup { get; set; }
    }


    public class SkypeObj
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Sys_id { get; set; }
        public string id { get; set; }
        public string channelId { get; set; }
        public string serviceUrl { get; set; }
        public Conversation conversation { get; set; }

    }

    public class SkypeAddress
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Sys_id { get; set; }
        public string Name { get; set; }
        public SkypeObj SkypeObj { get; set; }
        public List<KeyWordGC> ListKeyword { get; set; }
        public bool IsAllGC { get; set; } //true if notify all, false if keyword base
        public bool IsEnabled { get; set; } //true if enable, false to disable
        public bool IsBizDev { get; set; } //true if bizdev, false if internal
        public bool Is3PMNotification { get; set; } //true if it notifies all meeting for tomorrow at 3PM 
    }

    public class KeyWordGC
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Sys_id { get; set; }
        public string Keyword { get; set; }
    }

}
