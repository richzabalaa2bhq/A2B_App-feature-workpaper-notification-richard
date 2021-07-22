using A2B_App.Shared.User;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text.Json.Serialization;


namespace A2B_App.Shared.Skype
{
    public class Notification
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string Title { get; set; }
        public bool Group  { get; set; }
        public int UserId { get; set; }
        public virtual ICollection<Category> ListCategory { get; set; }

    }

    /// <summary>
    /// PTO
    /// Meeting
    /// BizDev
    /// 3PMNotification
    /// Workpaper
    /// RCM
    /// SoxTracker
    /// KeyReport
    /// </summary>
    public class Category
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string CategoryName { get; set; }
        public virtual ICollection<SkypeObj> ListSkypeObj { get; set; }
        public virtual ICollection<KeyWordGC> ListKeyWord { get; set; }
        public bool IsAllMeeting { get; set; } //Applicable in meeting's notification only

    }




}
