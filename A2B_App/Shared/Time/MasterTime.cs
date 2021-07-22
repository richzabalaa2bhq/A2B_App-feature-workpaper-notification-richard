using A2B_App.Shared.Podio;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text.Json.Serialization;

namespace A2B_App.Shared.Time
{
    public class MasterTime
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public DateTime Date { get; set; }
        public decimal Hours { get; set; }
        public string BillStat { get; set; }
        public string ClientCode { get; set; }
        public string ClientName { get; set; } //webhook client in mastertime
        public string Project { get; set; } //webhook project in mastertime
        public string Task { get; set; } //webhook task in mastertime
        public string Employee { get; set; } //webhook team member in mastertime
        public string Comment { get; set; }
        public PodioRef PodioRef { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
    }

    public class MasterTimeDetail
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string ClientRef { get; set; }
        public string ProjectRef { get; set; }
        public string TaskRef { get; set; }
        public string EmployeeRef { get; set; }
        public string Invoice { get; set; }
        public string InvoiceNum { get; set; }
        public string InvoicePeriod { get; set; }
        public string Todo { get; set; }
        public string Organization { get; set; }
        public string Manager { get; set; }
        public string TeamLead { get; set; }

        public DateTimeOffset? CreatedOn { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
    }



}
