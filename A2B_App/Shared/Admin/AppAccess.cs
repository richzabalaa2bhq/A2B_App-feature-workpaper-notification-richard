
using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text.Json.Serialization;

namespace A2B_App.Shared.Admin
{
    public class AppAccess
    {

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string User { get; set; }
        public string AppId { get; set; }
        public string AppToken { get; set; }
        public string Status { get; set; }
        public string CreatedBy { get; set; }
        public DateTime Created{ get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTime LastUpdate { get; set; } = DateTime.Now;
    }

    public class ModalYesWithId
    {
        public string Answer { get; set; }
        public string Id { get; set; }
    }
}
