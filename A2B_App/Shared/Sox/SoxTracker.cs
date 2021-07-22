using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text.Json.Serialization;

namespace A2B_App.Shared.Sox
{
    public class SoxTracker
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string FY { get; set; }
        public string ClientName { get; set; }
        public string ClientCode { get; set; }
        // public string KeyReport { get; set; }
        // public string KeyReportName { get; set; }
        //public string KeyReportC { get; set; }
        //public string KeyReportName { get; set; }
        //public string ControlOwner { get; set; }
        //public string ControlActivity { get; set; } 
        //public string TestProc { get; set; }
        public int? ClientItemId { get; set; }
        public string Process { get; set; }
        public string Subprocess { get; set; }
        public string ControlId { get; set; }
        public string PBC { get; set; }
        public string PBCOwner { get; set; }
        public string PBCOwnerOther { get; set; }
        public string PopulationFileRequest { get; set; }
        public string SampleSelection { get; set; }
        public string R3Sample { get; set; }
        public string WTPBC { get; set; }
        public string R1PBC { get; set; }
        public string R2PBC { get; set; }
        public string R3PBC { get; set; }
        public string WTTester { get; set; }
        public string WT1LReviewer { get; set; }
        public string WT2LReviewer { get; set; }
        public string R1Tester { get; set; }
        public string R11LReviewer { get; set; }
        public string R12LReviewer { get; set; }
        public string R2Tester { get; set; }
        public string R21LReviewer { get; set; }
        public string R22LReviewer { get; set; }
        public string R3Tester { get; set; }
        public string R31LReviewer { get; set; }
        public string R32LReviewer { get; set; }
        public string WTTestingStatus { get; set; }
        public string R1TestingStatus { get; set; }
        public string R2TestingStatus { get; set; }
        public string R3TestingStatus { get; set; }
        public int RCRWT { get; set; }
        public int RCRR1 { get; set; }
        public int RCRR2 { get; set; }
        public int RCRR3 { get; set; }
        public string ExternalAuditorSample { get; set; }

        public string KeyReport { get; set; }
        public string KeyReportName { get; set; }

        public string Status { get; set; }
        public TimeSpan? Duration { get; set; }

        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;


    }

    public class SoxTrackerQuestionnaire
    {

        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }


        public string Q6Label { get; set; }
        public int Q6FieldId { get; set; }

        public string Q7Label { get; set; }
        public int Q7FieldId { get; set; }

        public string Q8Label { get; set; }
        public int Q8FieldId { get; set; }     

        public string Q9Label { get; set; }
        public int Q9FieldId { get; set; }

        public string Q10Label { get; set; }
        public int Q10FieldId { get; set; }

        public string Q11Label { get; set; }
        public int Q11FieldId { get; set; }

        public string Q12ALabel { get; set; }
        public int Q12AFieldId { get; set; }

        public string Q12BLabel { get; set; }
        public int Q12BFieldId { get; set; }

        public string Q12CLabel { get; set; }
        public int Q12CFieldId { get; set; }

        public string Q12DLabel { get; set; }
        public int Q12DFieldId { get; set; }

        public string Q13ALabel { get; set; }
        public int Q13AFieldId { get; set; }

        public string Q13BLabel { get; set; }
        public int Q13BFieldId { get; set; }

        public string Q13CLabel { get; set; }
        public int Q13CFieldId { get; set; }

        public string Q13DLabel { get; set; }
        public int Q13DFieldId { get; set; }

        public string Q13ELabel { get; set; }
        public int Q13EFieldId { get; set; }

        public string Q13FLabel { get; set; }
        public int Q13FFieldId { get; set; }

        public string Q13GLabel { get; set; }
        public int Q13GFieldId { get; set; }

        public string Q13HLabel { get; set; }
        public int Q13HFieldId { get; set; }

        public string Q13ILabel { get; set; }
        public int Q13IFieldId { get; set; }

        public string Q13JLabel { get; set; }
        public int Q13JFieldId { get; set; }
 
        public string Q13KLabel { get; set; }
        public int Q13KFieldId { get; set; }

        public string Q13LLabel { get; set; }
        public int Q13LFieldId { get; set; }

        public string Q14ALabel { get; set; }
        public int Q14AFieldId { get; set; }

        public string Q14BLabel { get; set; }
        public int Q14BFieldId { get; set; }

        public string Q14CLabel { get; set; }
        public int Q14CFieldId { get; set; }

        public string Q14DLabel { get; set; }
        public int Q14DFieldId { get; set; }

        public string Q14ELabel { get; set; }
        public int Q14EFieldId { get; set; }

        /*
        public string Q15Label { get; set; }
        public int Q15EFieldId { get; set; }


        public string Q16Label { get; set; }
        public int Q16EFieldId { get; set; }

        */

        public virtual ICollection<SoxTrackerAppCategory> ListSoxTrackerAppCategory { get; set; }
        public virtual ICollection<SoxTrackerAppRelationship> ListSoxTrackerAppRelationship { get; set; }

        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;

    }

    public class SoxTrackerAppRelationship
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string Title { get; set; }
        public SoxTrackerQuestionnaire SoxTrackerQuestionnaire { get; set; }
        public int FieldId { get; set; }
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
    }

    public class SoxTrackerAppCategory
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string Option { get; set; }
        public int FieldId { get; set; }
        public SoxTrackerQuestionnaire SoxTrackerQuestionnaire { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
    }

    public class SoxTrackerViewAccess : SoxTracker
    {
        //public string FY { get; set; }
        //public string ClientName { get; set; }
        //public string Process { get; set; }
        //public string Subprocess { get; set; }
        //public string ControlId { get; set; }
        //public string PBC { get; set; }
        //public string PBCOwner { get; set; }
        //public string PopulationFileRequest { get; set; }
        //public string SampleSelection { get; set; }
        //public string R3Sample { get; set; }
        //public string WTPBC { get; set; }
        //public string R1PBC { get; set; }
        //public string R2PBC { get; set; }
        //public string R3PBC { get; set; }
        //public string WTTester { get; set; }


        //public string WT1LReviewer { get; set; }
        //public string WT2LReviewer { get; set; }
        //public string R1Tester { get; set; }
        //public string R11LReviewer { get; set; }
        //public string R12LReviewer { get; set; }
        //public string R2Tester { get; set; }
        //public string R1TestingStatus { get; set; }
        //public string R2TestingStatus { get; set; }
        //public string R21LReviewer { get; set; }
        //public string R22LReviewer { get; set; }
        //public string R3Tester { get; set; }
        //public string R31LReviewer { get; set; }
        //public string R32LReviewer { get; set; }
        //public string R3TestingStatus { get; set; }
        //public string WTTestingStatus { get; set; }
        

    }



}
