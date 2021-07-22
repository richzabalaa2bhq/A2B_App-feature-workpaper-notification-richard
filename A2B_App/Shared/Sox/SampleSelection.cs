using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text.Json.Serialization;

namespace A2B_App.Shared.Sox
{
    public class SampleSelection
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }

        public int? ClientId { get; set; }
        public string ClientName { get; set; }
        public string ExternalAuditor { get; set; }
        public string Q4R3SampleRequired { get; set; }
        public int? CountSampleQ4R3 { get; set; }
        public string Risk { get; set; }
        public int? AnnualPopulation { get; set; }
        public int? AnnualSampleSize { get; set; }
        public string Frequency { get; set; }
        public DateTimeOffset? Round1Start { get; set; }
        public DateTimeOffset? Round1End { get; set; }
        public DateTimeOffset? Round2Start { get; set; }
        public DateTimeOffset? Round2End { get; set; }
        public DateTimeOffset? Round3Start { get; set; }
        public DateTimeOffset? Round3End { get; set; }
        public int PopulationByRound1 { get; set; }
        public int DaysPeriodRound1 { get; set; }
        public int DaysPeriodRound2 { get; set; }
        public int DaysPeriodRound3 { get; set; }
        public int DaysPeriodRoundTot { get; set; }
        public int PopulationByRound2 { get; set; }
        public int PopulationByRound3 { get; set; }
        public int PopulationByRoundTot { get; set; }
        public int SamplesByRound1 { get; set; }
        public int SamplesByRound2 { get; set; }
        public int SamplesByRound3 { get; set; }
        public int SamplesByRoundTot { get; set; }
        public int SamplesCloseRound1 { get; set; }
        public int SamplesCloseRound2 { get; set; }
        public int SamplesCloseRound3 { get; set; }
        public int SamplesCloseRoundTot { get; set; }
        public int SamplesRemainingRound1 { get; set; }
        public int SamplesRemainingRound2 { get; set; }
        public int SamplesRemainingRound3 { get; set; }
        public int SamplesRemainingRoundTot { get; set; }

        public int WTSampleTested { get; set; }

        public string IsWTSampleTested { get; set; }

        #region Materiality Model
        public string Version { get; set; }
        public string IsMateriality { get; set; }
        public string ConsiderMateriality1 { get; set; }
        public string ConsiderMateriality2 { get; set; }
        public string ConsiderMateriality3 { get; set; }
        public string DisplayHeaderR11 { get; set; }
        public string DisplayHeaderR12 { get; set; }
        public string DisplayHeaderR13 { get; set; }
        public string DisplayHeaderR14 { get; set; }
        public string DisplayHeaderR15 { get; set; }
        public string DisplayHeaderR16 { get; set; }
        public string DisplayHeaderR17 { get; set; }
        public string DisplayHeaderR18 { get; set; }
        public string DisplayHeaderR19 { get; set; }
        public string DisplayHeaderR110 { get; set; }
        public string DisplayHeaderR111 { get; set; }
        public string DisplayHeaderR112 { get; set; }
        public string DisplayHeaderR113 { get; set; }
        public string DisplayHeaderR114 { get; set; }
        public string DisplayHeaderR115 { get; set; }
        public string DisplayHeaderR116 { get; set; }
        public string DisplayHeaderR117 { get; set; }
        public string DisplayHeaderR118 { get; set; }
        public string DisplayHeaderR119 { get; set; }
        public string DisplayHeaderR120 { get; set; }
        public string DisplayHeaderR21 { get; set; }
        public string DisplayHeaderR22 { get; set; }
        public string DisplayHeaderR23 { get; set; }
        public string DisplayHeaderR24 { get; set; }
        public string DisplayHeaderR25 { get; set; }
        public string DisplayHeaderR26 { get; set; }
        public string DisplayHeaderR27 { get; set; }
        public string DisplayHeaderR28 { get; set; }
        public string DisplayHeaderR29 { get; set; }
        public string DisplayHeaderR210 { get; set; }
        public string DisplayHeaderR211 { get; set; }
        public string DisplayHeaderR212 { get; set; }
        public string DisplayHeaderR213 { get; set; }
        public string DisplayHeaderR214 { get; set; }
        public string DisplayHeaderR215 { get; set; }
        public string DisplayHeaderR216 { get; set; }
        public string DisplayHeaderR217 { get; set; }
        public string DisplayHeaderR218 { get; set; }
        public string DisplayHeaderR219 { get; set; }
        public string DisplayHeaderR220 { get; set; }
        public string DisplayHeaderR31 { get; set; }
        public string DisplayHeaderR32 { get; set; }
        public string DisplayHeaderR33 { get; set; }
        public string DisplayHeaderR34 { get; set; }
        public string DisplayHeaderR35 { get; set; }
        public string DisplayHeaderR36 { get; set; }
        public string DisplayHeaderR37 { get; set; }
        public string DisplayHeaderR38 { get; set; }
        public string DisplayHeaderR39 { get; set; }
        public string DisplayHeaderR310 { get; set; }
        public string DisplayHeaderR311 { get; set; }
        public string DisplayHeaderR312 { get; set; }
        public string DisplayHeaderR313 { get; set; }
        public string DisplayHeaderR314 { get; set; }
        public string DisplayHeaderR315 { get; set; }
        public string DisplayHeaderR316 { get; set; }
        public string DisplayHeaderR317 { get; set; }
        public string DisplayHeaderR318 { get; set; }
        public string DisplayHeaderR319 { get; set; }
        public string DisplayHeaderR320 { get; set; }

        public string DisplayValue1 { get; set; }
        public string DisplayValue2 { get; set; }
        public string DisplayValue3 { get; set; }
        public string PopulationFile { get; set; }
        #endregion

        public virtual ICollection<TestRoundSampleSelectionReference> ListRefId { get; set; }
        public virtual ICollection<TestRound> ListTestRound { get; set; }
        [NotMapped]
        public virtual ICollection<TestRound> ListTestRound2 { get; set; } //this list is for ELC HRP 2.1
        [NotMapped]
        public virtual ICollection<TestRound> ListTestRound3 { get; set; } //this list is for ELC HRP 2.1
        public int RcmPodioItemId { get; set; }

        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public string SharefileLink { get; set; }
        public string JsonData { get; set; }

        public DateTimeOffset? CreatedOn { get; set; }

    }

    public class TestRound
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string TestingRound { get; set; }
        public string A2Q2Samples { get; set; }
        public string UniqueId { get; set; }
        public DateTimeOffset? Date { get; set; }
        public string HeaderRoundDisplay1 { get; set; }
        public string HeaderRoundDisplay2 { get; set; }
        public string HeaderRoundDisplay3 { get; set; }
        public string HeaderRoundDisplay4 { get; set; }
        public string HeaderRoundDisplay5 { get; set; }
        public string HeaderRoundDisplay6 { get; set; }
        public string HeaderRoundDisplay7 { get; set; }
        public string HeaderRoundDisplay8 { get; set; }
        public string HeaderRoundDisplay9 { get; set; }
        public string HeaderRoundDisplay10 { get; set; }
        public string HeaderRoundDisplay11 { get; set; }
        public string HeaderRoundDisplay12 { get; set; }
        public string HeaderRoundDisplay13 { get; set; }
        public string HeaderRoundDisplay14 { get; set; }
        public string HeaderRoundDisplay15 { get; set; }
        public string HeaderRoundDisplay16 { get; set; }
        public string HeaderRoundDisplay17 { get; set; }
        public string HeaderRoundDisplay18 { get; set; }
        public string HeaderRoundDisplay19 { get; set; }
        public string HeaderRoundDisplay20 { get; set; }
        public string MonthOnly { get; set; }
        public string WeeklyOnly { get; set; }
        public string ContentDisplay1 { get; set; }
        public string ContentDisplay2 { get; set; }
        public string ContentDisplay3 { get; set; }
        public string ContentDisplay4 { get; set; }
        public string ContentDisplay5 { get; set; }
        public string ContentDisplay6 { get; set; }
        public string ContentDisplay7 { get; set; }
        public string ContentDisplay8 { get; set; }
        public string ContentDisplay9 { get; set; }
        public string ContentDisplay10 { get; set; }
        public string ContentDisplay11 { get; set; }
        public string ContentDisplay12 { get; set; }
        public string ContentDisplay13 { get; set; }
        public string ContentDisplay14 { get; set; }
        public string ContentDisplay15 { get; set; }
        public string ContentDisplay16 { get; set; }
        public string ContentDisplay17 { get; set; }
        public string ContentDisplay18 { get; set; }
        public string ContentDisplay19 { get; set; }
        public string ContentDisplay20 { get; set; }
        public string Status { get; set; }
        public string Comment { get; set; }
        public SampleSelection SampleSelectionData { get; set; }

        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
    }

    public class TestRoundSampleSelectionReference
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public int PodioItemID { get; set; }
        public SampleSelection SampleSelectionData { get; set; }
    }

    public class PodioField
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string PodioUniqueId { get; set; }
        public string PodioItemId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }

        public SampleSelection SampleSelectionData { get; set; }
    }

    public class SampleSize
    {
        public string ExternalAuditor { get; set; }
        public string Frequency { get; set; }
        public string Risk { get; set; }
        public int SizeValue { get; set; }
        public int StartPopulation { get; set; }
        public int EndPopulation { get; set; }
    }

    public class DropDown
    {
        public string ExternalAuditor { get; set; }
        public int Percent { get; set; }
        public int PercentRound2 { get; set; }
    }

    public class ResponseStatus
    {
        public string Status { get; set; }
        public string Response { get; set; }
        public string PodioItemId { get; set; }
        public string PodioLink { get; set; }
        public string ExcelFilename { get; set; }
        public string ExcelLink { get; set; }
    }

    public class ClientSS
    {

        public string ClientName { get; set; }
        public string ExternalAuditor { get; set; }
        public int? ItemId { get; set; }
        public string SharefileId { get; set; }
        public int Percent { get; set; }
    }

    public class Frequency
    {
        public string Freq { get; set; }
        public int? IntValue { get; set; }
    }

    public class Population
    {
        public string UniqueId { get; set; }
        public string PurchaseOrder { get; set; }
        public string SupplierSched { get; set; }
        public string PoRev { get; set; }
        public string PoLine { get; set; }
        public string Requisition { get; set; }
        public string RequisitionLine { get; set; }
        public string EnteredBy { get; set; }
        public string Status { get; set; }
        public string Buyer { get; set; }
        public string Contact { get; set; }
        public string OrderDate { get; set; }
        public string Supplier { get; set; }
        public string ShipTo { get; set; }
        public string SortName { get; set; }
        public string Telephone { get; set; }
        public string ItemNumber { get; set; }
        public string ProdLine { get; set; }
        public string ProdDescription { get; set; }
        public string Site { get; set; }
        public string Location { get; set; }
        public string ItemRevision { get; set; }
        public string SupplierItem { get; set; }
        public string QuantityOrdered { get; set; }
        public string UnitOfMeasure { get; set; }
        public string UMConversion { get; set; }
        public string QtyOrderedXPOCost { get; set; }
        public string QuantityReceived { get; set; }
        public string QtyOpen { get; set; }
        public string QtyReturned { get; set; }
        public string DueDate { get; set; }
        public string OverDue { get; set; }
        public string PerformanceDate { get; set; }
        public string Currency { get; set; }
        public string StandardCost { get; set; }
        public string PurchasedCost { get; set; }
        public string PurCostBC { get; set; }
        public string OpenPoCost { get; set; }
        public string PpvPerUnit { get; set; }
        public string Type { get; set; }
        public string StdMtlCostNow { get; set; }
        public string WorkOrderId { get; set; }
        public string Operation { get; set; }
        public string PurchAcct { get; set; }
        public string GlAccountDesc { get; set; }
        public string CostCenter { get; set; }
        public string GlDescription { get; set; }
        public string Project { get; set; }
        public string Description { get; set; }
        public string Taxable { get; set; }
        public string Comments { get; set; }
        public int PopType { get; set; } // Type - 1, 2, 3 for population round

    }

    public class StringIndex
    {
        public string Text { get; set; }
        public int Index { get; set; }
    }

    public class SampleSelectionByRound
    {
        public string RoundName { get; set; }
        public SampleSelection SampleSelection { get; set; }
    }

    public class RegenerateData
    {
        public bool IsAll { get; set; }
        public int Index { get; set; }
    }

    public class SampleSelectionDraft
    {
        public SampleSelection SampleSelection { get; set; }
        public SampleSelectionProperties SampleSelProp { get; set; }
    }


}
