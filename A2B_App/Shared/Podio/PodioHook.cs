using System;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text.Json.Serialization;
namespace A2B_App.Shared.Podio
{
    public class PodioHook
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public long item_id { get; set; }
        public string item_revision_id { get; set; }
        public string type { get; set; }
        public int hook_id { get; set; }
        public string code { get; set; }
        public string status { get; set; }
        public string url { get; set; }
        public string AppId { get; set; }
        public DateTimeOffset? DateCreated { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset DateUpdated { get; set; } = DateTime.Now;
        
    }
}
