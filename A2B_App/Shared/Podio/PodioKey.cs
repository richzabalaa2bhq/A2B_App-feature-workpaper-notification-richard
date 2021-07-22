using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text.Json.Serialization;

namespace A2B_App.Shared.Podio
{
    public class PodioApiKey
    {
        public string ClientId { get; set; }
        public string ClientSecret { get; set; }
    }

    public class PodioAppKey
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string AppId { get; set; }
        public string AppToken { get; set; }
    }

    public class SyncDateRange
    {
        public DateTime? startDate { get; set; }
        public DateTime? endDate { get; set; }
        public int limit { get; set; }
        public int offset { get; set; }
    }

    public class PodioRef
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public int ItemId { get; set; }
        public string UniqueId { get; set; }
        public int Revision { get; set; }
        public string Link { get; set; }
        public string CreatedBy { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
    }


}
