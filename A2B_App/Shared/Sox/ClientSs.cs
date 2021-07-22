
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text.Json.Serialization;

namespace A2B_App.Shared.Sox
{
    public class ClientSs
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string Name { get; set; }
        public string ClientName { get; set; }
        public string ClientCode { get; set; }
        public int? ClientItemId { get; set; }
        public string ExternalAuditor { get; set; }
        public int? ItemId { get; set; }
        public string SharefileId { get; set; }
        public string SharefileScreenshotId { get; set; } //Sharefile key report screenshot id
        public string SharefileReportId { get; set; } //Sharefile key report folder id
        public int Percent { get; set; }
        public int PercentRound2 { get; set; }
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }

        public DateTimeOffset? CreatedOn { get; set; }

     
    }
}
