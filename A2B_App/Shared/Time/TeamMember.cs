using A2B_App.Shared.Podio;
using A2B_App.Shared.Skype;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace A2B_App.Shared.Time
{
    public class TeamMember
    {

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Newtonsoft.Json.JsonIgnore]
        public int Id { get; set; }
        public string Name { get; set; }
        //public virtual ICollection<UserEmail> ListEmail { get; set; }
        public string Email { get; set; }
        public string Organization { get; set; }
        public string Status { get; set; }
        public int UserId { get; set; }
        public int ProfileId { get; set; }
        public string SkypeName { get; set; }
        public string SkypeObjRaw { get; set; }
        public SkypeObj SkypeObj { get; set; }
        public ProfileImage ProfileImage { get; set; }
        public PodioRef PodioRef { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        public TeamMemberDetail TeamMemberDetail { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
    }

    public class TeamMemberDetail
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Newtonsoft.Json.JsonIgnore]
        public int Id { get; set; }
        public DateTime? BirthDate { get; set; }
        public DateTime? HiredDate { get; set; }
        public string LinkedInURL { get; set; }
        public string About { get; set; }
        public string Address { get; set; }
        public string Zip { get; set; }
        public string City { get; set; }
        public string Country { get; set; }
        public string State { get; set; }
        public string Mobile { get; set; }
        public string JobTitle { get; set; }
        public string Url { get; set; }
        public virtual ICollection<Skills> ListSkill { get; set; }

    }

    public class Skills
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Newtonsoft.Json.JsonIgnore]
        public int Id { get; set; }
        public string Skill { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
        [ForeignKey("TeamMemberDetail")]
        public int TeamMemberDetailId { get; set; }
    }

    public class Organization
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Newtonsoft.Json.JsonIgnore]
        public int Id { get; set; }
        public string OrgName { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
    }

    public class UserEmail
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Newtonsoft.Json.JsonIgnore]
        public int Id { get; set; }
        public string Email { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
    }

    public class UserStatus
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Newtonsoft.Json.JsonIgnore]
        public int Id { get; set; }
        public string Status { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
    }

    public class ProfileImage
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Newtonsoft.Json.JsonIgnore]
        public int Id { get; set; }
        public string Link { get; set; }
        public string Filename { get; set; }

    }

}
