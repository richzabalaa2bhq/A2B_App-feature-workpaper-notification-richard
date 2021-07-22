using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text.Json.Serialization;

namespace A2B_App.Shared.Sms
{
    public class GlobeSms
    {
        public string subscriber_number { get; set; }
        public string access_token { get; set; }
        public virtual unsubscribed unsubscribed { get; set; }
    }

    public class unsubscribed
    {
        public string subscriber_number { get; set; }
        public string time_stamp { get; set; }
    }

    public class Subscribe
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string SubScriberNum { get; set; }
        public string Token { get; set; }
        public string Status { get; set; }
        public DateTimeOffset? DateCreated { get; set; }
        public DateTimeOffset? DateUpdated { get; set; }
    }

    public class EmployeeSmsReference
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string EmployeeId { get; set; }
        public int SubscribeId { get; set; }
    }

    public class SmsSend
    {
        public string Address { get; set; }
        public string Message { get; set; }
    }
}
