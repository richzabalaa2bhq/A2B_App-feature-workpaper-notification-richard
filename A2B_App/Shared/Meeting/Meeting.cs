using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text.Json.Serialization;

namespace A2B_App.Shared.Meeting
{

    public class meeting
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int sys_id { get; set; }
        public string title { get; set; }
        public DateTime startdate { get; set; }
        public DateTime enddate { get; set; }
        public int duration { get; set; }
        public string platform { get; set; }
        public string pa { get; set; }
        public string poc { get; set; }
        public string team_member { get; set; }
        public string other_participant { get; set; }
        public string status { get; set; }
        public string help_needed { get; set; }
        public string agenda { get; set; }
        public string account { get; set; }
        public string day { get; set; }
        public string type { get; set; }
        public string meeting_id { get; set; }
        public string call_in { get; set; }
        public string pass_code { get; set; }
        public string meeting_link { get; set; }
        public string location { get; set; }
        public string notes { get; set; }
        public string response { get; set; }
        public string start_url { get; set; }
        public DateTime sys_created { get; set; }
        public DateTime sys_timestamp { get; set; }
        public string sys_status { get; set; }
        public string unique_id { get; set; }
        public DateTime created_on { get; set; }
        public string created_by { get; set; }
        public string item_id { get; set; }
        public int revision { get; set; }
        public string podio_link { get; set; }
        public string appointmentID { get; set; }
        public string occurrence_id { get; set; }
        public DateTime last_schedule_date { get; set; }
        public string auto_record { get; set; }
        public string client_name { get; set; }
        public string client_code { get; set; }
        public string client_emaildomain { get; set; }
    }

    public class recording_gtm
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int sys_id { get; set; }
        public string lastname { get; set; }
        public int num_attendees { get; set; }
        public string subject { get; set; }
        public string meeting_id { get; set; }
        public string organizer_key { get; set; }
        public string session_id { get; set; }
        public string locale { get; set; }
        public string meeting_type { get; set; }
        public string account_key { get; set; }
        public int duration { get; set; }
        public string firstname { get; set; }
        public string conference_call_info { get; set; }
        public string starttime { get; set; }
        public string endtime { get; set; }
        public string email { get; set; }
        public int status { get; set; }

    }
    public class recording_gtm_details
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int sys_id { get; set; }
        public string file_size { get; set; }
        public string recording_name { get; set; }
        public string download_url { get; set; }
        public string share_url { get; set; }
        public string recording_id { get; set; }
        public string meeting_id { get; set; }
        public string session_id { get; set; }
        public string organizer_key { get; set; }
        public int status { get; set; }
        public int status_transciption { get; set; }
    }
    public class dailymeeting
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int sys_id { get; set; }
        public string title { get; set; }
        public string item_id { get; set; }
        public string podio_link { get; set; }
        public string meeting_id { get; set; }
        public string startdate { get; set; }
        public string startdate_pst { get; set; }
        public string timeonly { get; set; }
        public int duration { get; set; }
        public string platform { get; set; }
        public string start_url { get; set; }
        public string account { get; set; }
        public string password { get; set; }
        public string pass_code { get; set; }
        public string agenda { get; set; }
        public string meeting_link { get; set; }

    }

    public class dailymeetingBizDev
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int sys_id { get; set; }
        public string Regarding { get; set; }
        public string Link { get; set; }
        public string MeetingWith { get; set; }
        public string DateTimeofMeeting { get; set; }
        public string startdate_PST { get; set; }
        public string TIMEONLY { get; set; }
        public string A2Q2Attendees { get; set; }
        public string Status { get; set; }
        public string MeetingLocationAddress { get; set; }
        public string ItemId { get; set; }
        

    }

    public class weeklymeeting
    {
        public int sys_id { get; set; }
        public string title { get; set; }
        public string item_id { get; set; }
        public string podio_link { get; set; }
        public string meeting_id { get; set; }
        public string startdate { get; set; }
        public string startdate_pst { get; set; }
        public string timeonly { get; set; }
        public int duration { get; set; }
        public string platform { get; set; }
        public string start_url { get; set; }
        public string account { get; set; }
        public string password { get; set; }
        public string pass_code { get; set; }
        public string agenda { get; set; }
        public string meeting_link { get; set; }
    }

    public class daterangemeeting
    {
        public int sys_id { get; set; }
        public string title { get; set; }
        public string item_id { get; set; }
        public string podio_link { get; set; }
        public string meeting_id { get; set; }
        public string startdate { get; set; }
        public string startdate_pst { get; set; }
        public string timeonly { get; set; }
        public int duration { get; set; }
        public string platform { get; set; }
        public string start_url { get; set; }
        public string account { get; set; }
        public string password { get; set; }
        public string pass_code { get; set; }
        public string agenda { get; set; }
        public string meeting_link { get; set; }
    }


    public class monthlymeeting
    {
        public int sys_id { get; set; }
        public string title { get; set; }
        public string item_id { get; set; }
        public string podio_link { get; set; }
        public string meeting_id { get; set; }
        public string startdate { get; set; }
        public string startdate_pst { get; set; }
        public string timeonly { get; set; }
        public int duration { get; set; }
        public string platform { get; set; }
        public string start_url { get; set; }
        public string account { get; set; }
        public string password { get; set; }
        public string pass_code { get; set; }
        public string agenda { get; set; }
        public string meeting_link { get; set; }


    }

    public class DateRangeCustom
    {
        public DateTimeOffset Start { get; set; }
        public DateTimeOffset End { get; set; }
    }

    public class dailymeetingAttendee
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int sys_id { get; set; }
        public string title { get; set; }
        public string item_id { get; set; }
        public string podio_link { get; set; }
        public string meeting_id { get; set; }
        public string startdate { get; set; }
        public string team_member { get; set; }
        public string pass_code { get; set; }
        public string other_participant { get; set; }
        public string startdate_pst { get; set; }
        public string timeonly { get; set; }
        public int duration { get; set; }
        public string platform { get; set; }
        public string start_url { get; set; }
        public string account { get; set; }
        //public string password { get; set; }
        public string meeting_link { get; set; }
    }

    public class meeting_skype
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int sys_id { get; set; }

    }

    public class MeetingAttendees
    {
        [JsonIgnore]
        public int meetingId { get; set; }
        public string meetingTitle { get; set; }
        public string attendees { get; set; }
        public string otherAttendees { get; set; }
    }

    public class t_employee
    {
        [JsonIgnore]
        public int f_empid { get; set; }
        public string f_name { get; set; }
        public string f_email { get; set; }
        public string f_status { get; set; }
    }

}
