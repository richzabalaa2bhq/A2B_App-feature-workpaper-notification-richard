using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text.Json.Serialization;

namespace A2B_App.Shared.Verification
{
    public class Verification
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string RequestedFromUserId { get; set; }
        public string AppUse { get; set; }
        public string RefId { get; set; }
        public string VerificationNum { get; set; }
        public string Status { get; set; }
        public string VerifiedVia { get; set; }
        public int Attempt { get; set; }
        public DateTimeOffset? ExpiryDate { get; set; }
        public DateTimeOffset? DateCreated { get; set; }
        public DateTimeOffset? DateUpdated { get; set; }
    }

    public class RequestVerification
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string RequestedFromUserId { get; set; }
        public string AppUse { get; set; }
        public string RefId { get; set; }
        public string VerificationNum { get; set; }
    }

}
