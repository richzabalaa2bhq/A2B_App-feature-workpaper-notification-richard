
using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text.Json.Serialization;

namespace A2B_App.Shared.Sox
{
    public class Matrix
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string ClientName { get; set; }
        public int? ClientItemId { get; set; }
        public string ClientCode { get; set; }
        public string ExternalAuditor { get; set; }
        public string Frequency { get; set; }
        public string Risk { get; set; }
        public int SizeValue { get; set; }
        public int StartPopulation { get; set; }
        public int EndPopulation { get; set; }

        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
    }
}
