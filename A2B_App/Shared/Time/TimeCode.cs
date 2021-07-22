using A2B_App.Shared.Podio;
using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text.Json.Serialization;
namespace A2B_App.Shared.Time
{
    public class TimeCode
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string ClientRef { get; set; }
        public string ClientCode { get; set; }
        public int ClientRefId { get; set; }
        public string ProjectRef { get; set; }
        public int ProjectRefId { get; set; }
        public string TaskRef { get; set; }
        public int TaskRefId { get; set; }
        public string Status { get; set; }
        public PodioRef PodioRef { get; set; }
        //public int PodioItemId { get; set; }
        //public string PodioUniqueId { get; set; }
        //public int PodioRevision { get; set; }
        //public string PodioLink { get; set; }
        //public string CreatedBy { get; set; }
        //public DateTimeOffset? CreatedOn { get; set; }

    }

    public class ClientReference
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string ClientRef { get; set; }
        public string ClientCode { get; set; }
        public string ClientHostedDomain { get; set; }
        public PodioRef PodioRef { get; set; }
        //public int PodioItemId { get; set; }
        //public string PodioUniqueId { get; set; }
        //public int PodioRevision { get; set; }
        //public string PodioLink { get; set; }
        //public string CreatedBy { get; set; }
        //public DateTimeOffset? CreatedOn { get; set; }
    }

    public class ProjectReference
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string ProjectRef { get; set; }
        public PodioRef PodioRef { get; set; }
        //public int PodioItemId { get; set; }
        //public string PodioUniqueId { get; set; }
        //public int PodioRevision { get; set; }
        //public string PodioLink { get; set; }
        //public string CreatedBy { get; set; }
        //public DateTimeOffset? CreatedOn { get; set; }
    }

    public class TaskReference
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string TaskRef { get; set; }
        public PodioRef PodioRef { get; set; }
        //public int PodioItemId { get; set; }
        //public string PodioUniqueId { get; set; }
        //public int PodioRevision { get; set; }
        //public string PodioLink { get; set; }
        //public string CreatedBy { get; set; }
        //public DateTimeOffset? CreatedOn { get; set; }
    }
}
