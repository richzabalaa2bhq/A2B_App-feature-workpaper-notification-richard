using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text.Json.Serialization;

namespace A2B_App.Shared.Sox
{
    public class RcmCta
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string ClientName { get; set; }
        public string ClientCode { get; set; }
        public string SubProcess { get; set; }
        public string FinancialStatementElement { get; set; }
        public string SpecificRisk { get; set; }
        public string Q5ACompletenessAccuracy { get; set; }
        public string Q5BExistenceOccur { get; set; }
        public string Q5CPresentationDisclose { get; set; }
        public string Q5DRightObligation { get; set; }
        public string Q5EValuationAlloc { get; set; }
        public string ControlOwner { get; set; }
        public string ControlPurpose { get; set; }
        public string ControlId { get; set; }
        public string ControlActivity { get; set; }
        public string ControlObjective { get; set; }
        public string TestingPeriod { get; set; }
        public string ControlType { get; set; }
        public string TestValidation { get; set; }
        public string MethodUsed { get; set; }
        public string Frequency { get; set; }
        public string Entity { get; set; }
        public string Notes { get; set; }
        public string KeyControl { get; set; }
        public string ControlPlaceDate { get; set; }
        public string NatureProcedure { get; set; }
        public string FraudControl { get; set; }
        public string ReviewControl { get; set; }
        public string Risk { get; set; }
        public string PerformTesting { get; set; }
        public string TestingProcedure { get; set; }
        public string Reviewer { get; set; }
        public string PopulationFileRequired { get; set; }

        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public string SharefileLink { get; set; }
        public string JsonData { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
    }

    public class RcmItemFilter
    {
        public string ClientName { get; set; }
        public string ControlName { get; set; }
        public DateTime? StartDate { get; set; }
        public DateTime? EndDate { get; set; }
        public int Limit { get; set; }
        public int Offset { get; set; }
        public string WorkpaperVersion { get; set; }
    }


}
