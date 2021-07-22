 using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text.Json.Serialization;

namespace A2B_App.Shared.Sox
{
    public class Rcm
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string ClientName { get; set; }
        public string ClientCode { get; set; }
        public int? ClientItemId { get; set; }

        public string ClientNameText { get; set; }
        public string FY { get; set; }
        public string Process { get; set; }
        public string Subprocess { get; set; }
        public string ControlObjective { get; set; }
        public string SpecificRisk { get; set; }
        public string FinStatemenElement { get; set; }

        public string FinancialStatementAssert { get; set; } // store checkbox answer #8

        public string CompletenessAccuracy { get; set; } // Yes or No
        public string ExistenceDisclosure { get; set; } // Yes or No
        public string PresentationDisclosure { get; set; } // Yes or No
        public string RigthsObligation { get; set; } // Yes or No
        public string ValuationAlloc { get; set; } // Yes or No
        public string ControlId { get; set; }
        public string ControlActivityFy19 { get; set; }
        public string ShortDescription { get; set; }
        public string ControlPlaceDate { get; set; }
        public string ControlOwner { get; set; }
        public string ControlFrequency { get; set; }
        public string Key { get; set; }
        public string ControlType { get; set; }
        public string NatureProc { get; set; }
        public string FraudControl { get; set; }
        public string RiskLvl { get; set; }
        public string ManagementRevControl { get; set; }
        public string PbcList { get; set; }
        public string TestProc { get; set; }
        public string TestingPeriod { get; set; }
        public string PopulationFileRequired { get; set; }
        public string MethodUsed { get; set; }
        public string TestValidation { get; set; }
        public string ControlActivity { get; set; }

        public DateTimeOffset? StartDate { get; set; }
        public DateTimeOffset? EndDate { get; set; }
        public TimeSpan? Duration { get; set; }
        public string AssignTo { get; set; }
        public string Reviewer { get; set; }
        public string Status { get; set; }
        public string WorkpaperStatus { get; set; }
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public string SharefileLink { get; set; }
        public string JsonData { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
    }

    public class RcmQuestionnaireField
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string AppId { get; set; }
        public string ControlName { get; set; }
        public string QuestionString { get; set; }
        public string Type { get; set; }
        public int Position { get; set; }
        public string Description { get; set; }
        public string Tag { get; set; }
        public int? FieldId { get; set; }
        public virtual ICollection<RcmQuestionnaireFieldOptions> Options { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        public DateTimeOffset? UpdatedOn { get; set; }
    }

    public class RcmQuestionnaireFieldOptions
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string OptionName { get; set; }
        public string OptionId { get; set; }
        public int AppId { get; set; }
        public RcmQuestionnaireField RcmQuestionnaire { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        public DateTimeOffset? UpdatedOn { get; set; }
    }

    //RCM Questionnaire Model
    public class RcmQuestionnaire
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }

        public string Q1Label { get; set; }
        public int Q1FieldId { get; set; }
        public virtual ICollection<RcmFY> ListQ1Year { get; set; }
        
        public string Q2Label { get; set; }
        public int Q2FieldId { get; set; }
        public string Q2Client { get; set; }
        public int Q2ClientItemId { get; set; }

        public string Q3Label { get; set; }
        public int Q3FieldId { get; set; }
        public string Q3Process { get; set; }
        public int Q3ProcessItemId { get; set; }

        public string Q4Label { get; set; }
        public int Q4FieldId { get; set; }
        public string Q4SubProcess { get; set; }
        public int Q4SubProcessItemId { get; set; }

        public string Q5Label { get; set; }
        public int Q5FieldId { get; set; }
        public string Q5ControlObjective { get; set; }
        public string Q6Label { get; set; }
        public int Q6FieldId { get; set; }
        public string Q6SpecificRisk { get; set; }

        public string Q7Label { get; set; }
        public int Q7FieldId { get; set; }
        public virtual ICollection<RcmFinancialStatement> ListQ7FinStatementElement { get; set; }

        public string Q8Label { get; set; }
        public int Q8FieldId { get; set; }
        public virtual ICollection<RcmFinancialStatementAssert> ListQ8FinStatementAssert { get; set; }

        public string Q9Label { get; set; }
        public int Q9FieldId { get; set; }
        public virtual ICollection<RcmControlId> ListQ9ControlId { get; set; }


        public string Q10Label { get; set; }
        public int Q10FieldId { get; set; }
        public string Q10ControlActivity { get; set; }


        public string Q11Label { get; set; }
        public int Q11FieldId { get; set; }
        public string Q11ControlInPlace { get; set; }

        public string Q12Label { get; set; }
        public int Q12FieldId { get; set; }
        public virtual ICollection<RcmControlOwner> ListQ12ControlOwner { get; set; }

        public string Q13Label { get; set; }
        public int Q13FieldId { get; set; }
        public virtual ICollection<RcmFrequency> ListQ13Frequency { get; set; }

        public string Q14Label { get; set; }
        public int Q14FieldId { get; set; }
        public virtual ICollection<RcmControlKey> ListQ14ControlKey { get; set; }

        public string Q15Label { get; set; }
        public int Q15FieldId { get; set; }
        public virtual ICollection<RcmControlType> ListQ15ControlType { get; set; }

        public string Q16Label { get; set; }
        public int Q16FieldId { get; set; }
        public virtual ICollection<RcmNatureProcedure> ListQ16NatureProcedure { get; set; }

        public string Q17Label { get; set; }
        public int Q17FieldId { get; set; }
        public virtual ICollection<RcmFraudControl> ListQ17FraudControl { get; set; }

        public string Q18Label { get; set; }
        public int Q18FieldId { get; set; }
        public virtual ICollection<RcmRiskLevel> ListQ18RiskLevel { get; set; }

        public string Q19Label { get; set; }
        public int Q19FieldId { get; set; }
        public virtual ICollection<RcmManagementReviewControl> ListQ19MgmtReviewControl { get; set; }

        public string Q20Label { get; set; }
        public int Q20FieldId { get; set; }
        public string Q20PBCNeeded { get; set; }


        public string Q21Label { get; set; }
        public int Q21FieldId { get; set; }
        public string Q21TestProcedure { get; set; }

        public string Status { get; set; }

        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;

    }

    public class RcmFY
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string FY { get; set; }
        public RcmQuestionnaire RcmQuestionnaire { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
    }


    //RCM Process Model
    public class RcmProcess
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string Process { get; set; }
        public RcmQuestionnaire RcmQuestionnaire { get; set; }
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;

    }

    //RCM Sub Process Model
    public class RcmSubProcess
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string SubProcess { get; set; }
        public RcmQuestionnaire RcmQuestionnaire { get; set; }
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
    }

    //RCM Financial Statement
    public class RcmFinancialStatement
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string FinStatement { get; set; }
        public RcmQuestionnaire RcmQuestionnaire { get; set; }
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
    }

    public class RcmFinancialStatementAssert
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string FinStatementAssert { get; set; }
        public RcmQuestionnaire RcmQuestionnaire { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
    }

    //RCM Control Id
    public class RcmControlId
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string ControlId { get; set; }
        public RcmQuestionnaire RcmQuestionnaire { get; set; }
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
    }

    //RCM Control Owner
    public class RcmControlOwner
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string ControlOwner { get; set; }
        public RcmQuestionnaire RcmQuestionnaire { get; set; }
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
    }

    //RCM Frequency
    public class RcmFrequency
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string Frequency { get; set; }
        public RcmQuestionnaire RcmQuestionnaire { get; set; }
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
    }

    public class RcmControlKey
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string Option { get; set; }
        public RcmQuestionnaire RcmQuestionnaire { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
    }

    public class RcmFraudControl
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string Option { get; set; }
        public RcmQuestionnaire RcmQuestionnaire { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
    }

    public class RcmRiskLevel
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string Option { get; set; }
        public RcmQuestionnaire RcmQuestionnaire { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
    }

    //RCM Control Type
    public class RcmControlType
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string ControlType { get; set; }
        public RcmQuestionnaire RcmQuestionnaire { get; set; }
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
    }

    //RCM Nature Procedure
    public class RcmNatureProcedure
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string NatureProcedure { get; set; }
        public RcmQuestionnaire RcmQuestionnaire { get; set; }
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
    }

    //RCM Nature Procedure
    public class RcmManagementReviewControl
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string MgmtReviewControl { get; set; }
        public RcmQuestionnaire RcmQuestionnaire { get; set; }
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
    }

    public class RcmQuestionnaireFilter
    {
        public string FY { get; set; }
        public string Client { get; set; }
        public string Process { get; set; }
        public string SubProcess { get; set; }
        public string ControlId { get; set; }
    }

    public class RcmQ13toQ19
    {
        public List<CheckBoxItem> ListQ7FinStatementElement { get; set; }
        public List<CheckBoxItem> ListQ8FinStatementAssert { get; set; }
        public List<CheckBoxItem> ListQ12ControlOwner { get; set; }
        public List<string> ListQ13Frequency { get; set; }
        public List<string> ListQ14ControlKey { get; set; }
        public List<string> ListQ15ControlType { get; set; }
        public List<string> ListQ16NatureProcedure { get; set; }
        public List<string> ListQ17FraudControl { get; set; }
        public List<string> ListQ18RiskLevel { get; set; }
        public List<string> ListQ19MgmtReviewControl { get; set; }
    }

    public class CheckBoxItem
    {
        public bool selected { get; set; }
        public string item { get; set; }
    }

    public class RcmOutputFile
    {
        public string ClientName { get; set; }
        public string LoadingStatus { get; set; } 
        public string SharefileLink { get; set; }
        public string FileName { get; set; }
    }

    public class RcmViewAccess
    {
        public string ClientName { get; set; }
        public string FY { get; set; }
        public string Process { get; set; }
        public string Subprocess { get; set; }
        public string ControlObjective { get; set; }
        public string SpecificRisk { get; set; }
        public string FinStatementElement { get; set; }
        public string CompletenessAccuracy { get; set; }
        public string ExistenceDisclosure { get; set; }
        public string PresentationDisclosure { get; set; }
        public string RightsObligation { get; set; }
        public string ValuationAlloc { get; set; }
        public string ControlActivityFy19 { get; set; }
        public string ControlId { get; set; }
        public string ControlPlaceDate { get; set; }
        public string ControlOwner { get; set; }
        public string ControlFrequency { get; set; }
        public string Key { get; set; }
        public string ControlType { get; set; }
        public string NatureProc { get; set; }
        public string FraudControl { get; set; }
        public string RiskLvl { get; set; }
        public string ManagementRevControl { get; set; }
        public string PbcList { get; set; }
        public string TestProc { get; set; }
    }


}
