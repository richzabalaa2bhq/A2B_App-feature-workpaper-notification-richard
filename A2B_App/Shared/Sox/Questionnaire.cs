using A2B_App.Shared.Podio;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text.Json.Serialization;

namespace A2B_App.Shared.Sox
{
    public class Questionnaire
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string ClientName { get; set; }
        public string Process { get; set; }
        public string ControlOwner { get; set; }
        public string ControlId { get; set; }
        public string ControlActivity { get; set; }
        public string TestValidation { get; set; }
        public string MethodUsed { get; set; }
        public string ControlFrequency { get; set; }
        public string ControlPlaceDate { get; set; }
        public string RiskAssessment { get; set; }
        public string TestingPhase { get; set; }
        public string SamplePeriod { get; set; }
        public string PopulationSize { get; set; }
        public string SampleSizeDerivation { get; set; }
        public int SampleSizeRound1 { get; set; }
        public int SampleSizeRound2 { get; set; }
        public int SampleSizeRound3 { get; set; }
        public string IpeInformation { get; set; }
        public string SourceFileType { get; set; }
        public string DocumentProvided { get; set; }
        public string DocumentReceived { get; set; }

        public virtual ICollection<SampleRound> ListSampleRound { get; set; }
        public virtual ICollection<NotesUnique> ListNotes { get; set; }

        public string TestOfDesign { get; set; }
        public string TestOperatingEffectiveness { get; set; }
        public string TestingStatus { get; set; }
        public string TestPerfomedBy { get; set; }
        public DateTimeOffset? DateOfTesting { get; set; }
        public string ReviewedBy { get; set; }
        public DateTimeOffset? DateOfReviewed { get; set; }
        public string ReviewersNote { get; set; }
        public int SampleSelectionPodioItemId { get; set; }
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

    public class NotesUnique
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string Key { get; set; }
        public string Description { get; set; }
        public int Index { get; set; }
    }

    public class QuestionaireAddedInputs  //edited {search keyword "mark" to find it easily}
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string ClientName { get; set; }
        public string ClientCode { get; set; }
        public int? ClientItemId { get; set; }
        public  string roundset { get; set; }
        public string Answer1 { get; set; } = "";
        public string Answer2 { get; set; } = "";
        public string Answer3 { get; set; } = "";
        public string Answer4 { get; set; } = "";
        public string Answer5 { get; set; } = "";
        public string Answer6 { get; set; } = "";
        public string Answer7 { get; set; } = "";
        public string Answer8 { get; set; } = "";
        public string Answer9 { get; set; } = "";
        public string Answer10 { get; set; } = "";
        public string Answer11 { get; set; } = "";
        public string Answer12 { get; set; } = "";
        public string Answer13 { get; set; } = "";
        public string Answer14 { get; set; } = "";
        public string Answer15 { get; set; } = "";

    }
    public class SampleRound
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string SampleNum { get; set; }
        public string RoundName { get; set; }
        public string MedrioActivity16A { get; set; }
        public string InvestigationPerformed16B { get; set; }
        public string MedrioCostAgreements16C { get; set; }
        public string MedrioReportReview16D { get; set; }
        public string MedrioReviewers16E { get; set; }
        public DateTimeOffset? DateMedrioReviewed16F { get; set; }
        public string Content16G { get; set; }
        public string Notes { get; set; }

        public string TickBox1Value { get; set; }
        public string TickBox2Value { get; set; }
        public string TickBox3Value { get; set; }
        public string TickBox4Value { get; set; }
        public string TickBox5Value { get; set; }
        public string TickBox6Value { get; set; }
        public string TickBox7Value { get; set; }
    }

    public class soxtracker
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string FY { get; set; }
        public string ClientName { get; set; }
        public string Process { get; set; }
        public string Subprocess { get; set; }
        public string ControlId { get; set; }
        public string PBC { get; set; }
        public string PBCOwner { get; set; }
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
        public short RCRWT { get; set; }
        public short RCRR1 { get; set; }
        public short RCRR2 { get; set; }
        public short RCRR3 { get; set; }
        public string ExternalAuditorSample { get; set; }
    }

    public class QuestionnaireQuestion
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string ClientName { get; set; }
        public string AppId { get; set; }
        public string ControlName { get; set; }
        public string QuestionString { get; set; }
        public string Type { get; set; }
        public int Position { get; set; }
        public string Description { get; set; }
        public string DtEndRequire { get; set; }
        public int? FieldId { get; set; }
        public virtual ICollection<QuestionnaireOption> Options { get; set; }

        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset UpdatedOn { get; set; } = DateTime.Now;
    }

    public class QuestionnaireOption
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string OptionName { get; set; }
        public string OptionId { get; set; }
        public int AppId { get; set; }
        public QuestionnaireQuestion QuestionnaireQuestion { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset UpdatedOn { get; set; } = DateTime.Now;
    }

    public class QuestionnaireUserInput
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string StrAnswer { get; set; }
        public string StrAnswer2 { get; set; } //Added answer 2 for the second answer, usually use in date range - end date.
        public string StrQuestion { get; set; }
        public string Description { get; set; }
        public int Position { get; set; }
        public string AppId { get; set; }
        public int? FieldId { get; set; }
        public int ItemId { get; set; }
        public string Type { get; set; }
        public string DtEndRequire { get; set; }
        public string RoundName { get; set; }

        public virtual ICollection<RoundItem> ListRoundItem { get; set; } // for the 1st round table or testing attributes
        public virtual ICollection<RoundItem> ListRoundItem2 { get; set; } //for the 2nd round table or testing attributes
        public virtual ICollection<RoundItem> ListRoundItem3 { get; set; } //for the 3rd round table or testing attributes
        public virtual ICollection<NotesItem> ListNoteItem { get; set; }
        public virtual ICollection<IUCSystemGenAnswer> ListIUCSystemGen { get; set; }
        public virtual ICollection<IUCNonSystemGenAnswer> ListIUCNonSystemGen { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset UpdatedOn { get; set; } = DateTime.Now;

    }

    public class ClientControl
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string ClientName { get; set; }
        public string ControlName { get; set; }
    }

    public class QuestionnaireFieldParam
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string ClientName { get; set; }
        public string ControlName { get; set; }
        public PodioAppKey AppKey { get; set; }
        //public string AppId { get; set; }
        //public string AppToken { get; set; }
    }

    public class RoundItem
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public int ItemID { get; set; }
        public string AppId { get; set; }
        public string RoundName { get; set; }
        public string Position { get; set; }
        public string A2Q2Samples { get; set; }

        public string Question1 { get; set; }
        public string Answer1 { get; set; }
        public string OtherAnswer1 { get; set; }
        public string Note1 { get; set; }
        public string Question2 { get; set; }
        public string Answer2 { get; set; }
        public string OtherAnswer2 { get; set; }
        public string Note2 { get; set; }
        public string Question3 { get; set; }
        public string Answer3 { get; set; }
        public string OtherAnswer3 { get; set; }
        public string Note3 { get; set; }
        public string Question4 { get; set; }
        public string Answer4 { get; set; }
        public string OtherAnswer4 { get; set; }
        public string Note4 { get; set; }
        public string Question5 { get; set; }
        public string Answer5 { get; set; }
        public string OtherAnswer5 { get; set; }
        public string Note5 { get; set; }
        public string Question6 { get; set; }
        public string Answer6 { get; set; }
        public string OtherAnswer6 { get; set; }
        public string Note6 { get; set; }
        public string Question7 { get; set; }
        public string Answer7 { get; set; }
        public string OtherAnswer7 { get; set; }
        public string Note7 { get; set; }
        public string Question8 { get; set; }
        public string Answer8 { get; set; }
        public string OtherAnswer8 { get; set; }
        public string Note8 { get; set; }
        public string Question9 { get; set; }
        public string Answer9 { get; set; }
        public string OtherAnswer9 { get; set; }
        public string Note9 { get; set; }
        public string Question10 { get; set; }
        public string Answer10 { get; set; }
        public string OtherAnswer10 { get; set; }
        public string Note10 { get; set; }
        public string Question11 { get; set; }
        public string Answer11 { get; set; }
        public string OtherAnswer11 { get; set; }
        public string Note11 { get; set; }
        public string Question12 { get; set; }
        public string Answer12 { get; set; }
        public string OtherAnswer12 { get; set; }
        public string Note12 { get; set; }
        public string Question13 { get; set; }
        public string Answer13 { get; set; }
        public string OtherAnswer13 { get; set; }
        public string Note13 { get; set; }
        public string Question14 { get; set; }
        public string Answer14 { get; set; }
        public string OtherAnswer14 { get; set; }
        public string Note14 { get; set; }
        public string Question15 { get; set; }
        public string Answer15 { get; set; }
        public string OtherAnswer15 { get; set; }
        public string Note15 { get; set; }


        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public string SharefileLink { get; set; }
        public string JsonData { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }

        [ForeignKey("QuestionnaireRoundSet")]
        public int QuestionnaireRoundSetId { get; set; }
    }

    public class RoundItem2
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public int ItemID { get; set; }
        public string AppId { get; set; }
        public string RoundName { get; set; } //Header, Round 1, Round 2, Round 3
        public int Position { get; set; }
        public string A2Q2Samples { get; set; }

        public virtual ICollection<RoundQA2> ListRoundQA { get; set; } // for the 1st round table or testing attributes

        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public string SharefileLink { get; set; }
        public string JsonData { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }

    }

    public class ListRoundItem2
    {
        public List<RoundItem2> ListRoundItem { get; set; }
        public List<NotesItem2> ListUniqueNotes { get; set; }
    }

    public class RoundQA
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string Position { get; set; }
        public string Question { get; set; }
        public string Answer { get; set; }
        public string Type { get; set; }
        public string DtEndRequire { get; set; }
        public virtual ICollection<QuestionnaireOption> Options { get; set; }
    }

    public class RoundQA2
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public int Position { get; set; }
        public string Question { get; set; }
        public string Answer { get; set; }
        public string Answer2 { get; set; }
        public string Note { get; set; }
        public string Type { get; set; }
        public string DtEndRequire { get; set; }
        public virtual ICollection<QuestionnaireOption> Options { get; set; }
    }

    public class NotesItem
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string Position { get; set; }
        public string Notes { get; set; }
        public string Description { get; set; }
        public int PodioItemId { get; set; }

        [ForeignKey("QuestionnaireRoundSet")]
        public int QuestionnaireRoundSetId { get; set; }
    }

    public class NotesItem2
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public int Position { get; set; }
        public string Notes { get; set; }
        public string Description { get; set; }
        public int PodioItemId { get; set; }
    }

    public class QuestionnaireExcelData
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public virtual ICollection<QuestionnaireUserInput> ListUserInputItem { get; set; }
        public SampleSelection SampleSelection { get; set; }
        public Rcm Rcm { get; set; }
        public GeneralNote GeneralNote { get; set; }
        public List<IPENote> ListIPENote { get; set; }
        public List<HeaderNote> ListHeaderNote { get; set; }
        public List<HeaderNote> ListHeaderNote2 { get; set; }
        public List<HeaderNote> ListHeaderNote3 { get; set; }
        public List<PolicyChanges> ListPolicyNote { get; set; }
        public List<IUCSystemGenAnswer> ListIUCSystemGen { get; set; }
        public List<IUCNonSystemGenAnswer> ListIUCNonSystemGen { get; set; }
    }

    public class QuestionnaireData
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public virtual ICollection<QuestionnaireUserInput> ListUserInputRound1 { get; set; }
        public virtual ICollection<QuestionnaireUserInput> ListUserInputRound2 { get; set; }
        public virtual ICollection<QuestionnaireUserInput> ListUserInputRound3 { get; set; }
        public SampleSelection SampleSelection { get; set; }
        public Rcm Rcm { get; set; }
        public GeneralNote GeneralNote { get; set; }
        public List<IPENote> ListIPENote { get; set; }
        public List<HeaderNote> ListHeaderNote { get; set; }
        public List<HeaderNote> ListHeaderNote2 { get; set; }
        public List<HeaderNote> ListHeaderNote3 { get; set; }
        public List<PolicyChanges> ListPolicyNote { get; set; }
        public List<IUCSystemGen> ListIUCSystemGen { get; set; }
        public List<IUCNonSystemGen> ListIUCNonSystemGen { get; set; }
    }

    
    public class DtQuestionnaire
    {
        public DateTimeOffset? startDate { get; set; }
        public DateTimeOffset? endDate { get; set; }
        public int position { get; set; }
    }

    public class IPENote
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string Name { get; set; }
        public string Note { get; set; }
        public string Description { get; set; }
        public bool Display { get; set; }
    }

    public class ListIPENote
    {
        public List<IPENote> ListNotes { get; set; }
    }

    public class GeneralNote
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string GeneralNoteText { get; set; }
        public string Description { get; set; }
        public bool Display { get; set; }
    }

    public class HeaderNote
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string HeaderNoteText { get; set; }
        public string Description { get; set; }
        public int Position { get; set; }

        [ForeignKey("QuestionnaireRoundSet")]
        public int QuestionnaireRoundSetId { get; set; }
    }

    public class HeaderNote2
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string StrQuestion { get; set; }
        public string HeaderNoteText { get; set; }
        public string Description { get; set; }
        public int Position { get; set; }

        public virtual ICollection<HeaderSeventeen> ListSevenTeen { get; set; }
    }

    public class HeaderSeventeen
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string Name { get; set; }
    }

    public class PolicyChanges
    {
        public string NoteText { get; set; }
        public string Description { get; set; }
        public int Position { get; set; }
    }

    public class UniqueNoteCallback
    {
        public int Position { get; set; }
        public List<NotesItem> Item { get; set; }
    }

    public class IUCSystemGen
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string AppId { get; set; }
        public int Position { get; set; }
        public virtual ICollection<QandA> ListQuestionAnswer { get; set; }
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public string SharefileLink { get; set; }
        public string JsonData { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
    }

    public class IUCNonSystemGen
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string AppId { get; set; }
        public int Position { get; set; }
        public virtual ICollection<QandA> ListQuestionAnswer { get; set; }
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public string SharefileLink { get; set; }
        public string JsonData { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
    }

    public class QandA
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string UniqueId { get; set; }
        public int? FieldId { get; set; }
        public string AppId { get; set; }
        public int Position { get; set; }
        public string Type { get; set; }
        public string Question { get; set; }
        public string Answer { get; set; }
        public string Options { get; set; }
    }

    public class IUCSystemGenIndexOf
    {
        public IUCSystemGen IUCSystemGen { get; set; }
        public int IndexOf { get; set; }
    }

    public class IUCNonSystemGenIndexOf
    {
        public IUCNonSystemGen IUCNonSystemGen { get; set; }
        public int IndexOf { get; set; }
    }

    public class FinalIUCSystemGen
    {
        public List<IUCSystemGenAnswer> Item { get; set; }
        public string roundName { get; set; }
    }

    public class FinalIUCNonSystemGen
    {
        public List<IUCNonSystemGenAnswer> Item { get; set; }
        public string roundName { get; set; }
    }

    public class FinalListUniqueNotes
    {
        public List<NotesItem> Item { get; set; }
    }

    //-----------------------------------------------------
    //Model for new UI
    //-----------------------------------------------------
    public class QuestionnaireUserAnswer
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string StrAnswer { get; set; }
        public string StrAnswer2 { get; set; } //Added answer 2 for the second answer, usually use in date range - end date.
        public string StrQuestion { get; set; }
        public string StrDefaultAnswer { get; set; }
        public bool IsDisabled { get; set; }
        public string Description { get; set; }
        public int Position { get; set; }
        public string AppId { get; set; }
        public int? FieldId { get; set; }
        public int ItemId { get; set; }
        public string Type { get; set; }
        public string DtEndRequire { get; set; }
        public string RoundName { get; set; }
        public virtual ICollection<QuestionnaireOption> Options { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset UpdatedOn { get; set; } = DateTime.Now;

    }

    public class QuestionnaireUserQuestion
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string ClientName { get; set; }
        public string AppId { get; set; }
        public string ControlName { get; set; }
        public string QuestionString { get; set; }
        public string Type { get; set; }
        public int Position { get; set; }
        public string Description { get; set; }
        public string DtEndRequire { get; set; }
        public int? FieldId { get; set; }
        public virtual ICollection<QuestionnaireOption> Options { get; set; }

        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset UpdatedOn { get; set; } = DateTime.Now;
    }

    public class QuestionnaireRoundSet
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public virtual ICollection<QuestionnaireUserAnswer> ListUserInputRound1 { get; set; }
        public virtual ICollection<QuestionnaireUserAnswer> ListUserInputRound2 { get; set; }
        public virtual ICollection<QuestionnaireUserAnswer> ListUserInputRound3 { get; set; }
        public virtual ICollection<NotesItem> ListUniqueNotes { get; set; }
        public virtual ICollection<RoundItem> ListRoundItem { get; set; }
        public virtual ICollection<IUCSystemGenAnswer> ListIUCSystemGen1 { get; set; }
        public virtual ICollection<IUCSystemGenAnswer> ListIUCSystemGen2 { get; set; }
        public virtual ICollection<IUCSystemGenAnswer> ListIUCSystemGen3 { get; set; }
        public virtual ICollection<IUCNonSystemGenAnswer> ListIUCNonSystemGen1 { get; set; }
        public virtual ICollection<IUCNonSystemGenAnswer> ListIUCNonSystemGen2 { get; set; }
        public virtual ICollection<IUCNonSystemGenAnswer> ListIUCNonSystemGen3 { get; set; }
        public virtual ICollection<HeaderNote> ListHeaderNote { get; set; }
        [NotMapped]
        public virtual ICollection<QuestionnaireQuestion> ListUserQuestion { get; set; }
        public SampleSelection sampleSel1 { get; set; }
        public SampleSelection sampleSel2 { get; set; }
        public SampleSelection sampleSel3 { get; set; }
        public bool isRound1 { get; set; }
        public bool isRound2 { get; set; }
        public bool isRound3 { get; set; }
        public string UniqueId { get; set; }
        //public WorkpaperStatus WorkpaperStatus { get; set; }
        //public string AddedBy { get; set; }
        //public int Version { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset UpdatedOn { get; set; } = DateTime.Now;
        public Rcm Rcm { get; set; }
        [NotMapped]
        public QuestionaireAddedInputs QuestionaireAddedInputss { get; set; }
        [NotMapped]
        public Seventeen Seventeenn { get; set; }

    }

    public class QuestionnaireTesterSet
    {
        //Tester will create workpaper
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string UniqueId { get; set; }
        public virtual ICollection<QuestionnaireUserAnswer> ListUserInputRound { get; set; }
        public virtual ICollection<IUCSystemGenAnswer> ListIUCSystemGen { get; set; }
        public virtual ICollection<IUCNonSystemGenAnswer> ListIUCNonSystemGen { get; set; }
        public virtual ICollection<NotesItem2> ListUniqueNotes { get; set; }
        public virtual ICollection<RoundItem2> ListRoundItem2 { get; set; }
        public virtual ICollection<HeaderNote2> ListHeaderNote2 { get; set; }
        public virtual ICollection<IPENote> ListIPENote { get; set; }
        public SampleSelection SampleSel { get; set; }
        public WorkpaperStatus WorkpaperStatus { get; set; }
        public GeneralNote GeneralNote { get; set; }
        [NotMapped]
        public SampleSelectionProperties SampleSelProp { get; set; }
        public string RoundName { get; set; } //Round 1, Round 2, Round 3
        public string AddedBy { get; set; }
        public int DraftNum { get; set; } //odd - tester, even -reviewer
        public int RcmItemId { get; set; }
        public Rcm Rcm { get; set; }
        public int Position { get; set; }
        public string UserAction { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset UpdatedOn { get; set; } = DateTime.Now;
        [NotMapped]
        public QuestionaireAddedInputs QuestionaireAddedInputss { get; set; }
        [NotMapped]
        public Seventeen Seventeenn { get; set; }
    }

    public class SampleSelectionProperties
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string ClientName { get; set;}
        public string Risk { get; set; }
        public string Frequency { get; set; }
        public string Version { get; set; }
        public bool display { get; set; }
    }

    public class QuestionnaireReviewerSet 
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string UniqueId { get; set; }
        //public virtual ICollection<QuestionnaireUserAnswer> ListUserInputRound { get; set; }
        //public virtual ICollection<IUCSystemGenAnswer> ListIUCSystemGen { get; set; }
        //public virtual ICollection<IUCNonSystemGenAnswer> ListIUCNonSystemGen { get; set; }
        //public virtual ICollection<NotesItem> ListUniqueNotes { get; set; }
        //public virtual ICollection<RoundItem> ListRoundItem { get; set; }
        public SampleSelection SampleSel { get; set; }
        public WorkpaperStatus WorkpaperStatus { get; set; }
        public string RoundName { get; set; } //Round 1, Round 2, Round 3
        public string AddedBy { get; set; }
        public int DraftNum { get; set; }
        public int RcmItemId { get; set; }
        public Rcm Rcm { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset UpdatedOn { get; set; } = DateTime.Now;

        [NotMapped]
        public QuestionaireAddedInputs QuestionaireAddedInputss { get; set; }
        [NotMapped]
        public Seventeen Seventeenn { get; set; }
    }

    public class QuestionnaireDraftId
    {
        public string UniqueId { get; set; }
        public int RcmItemId { get; set; }
        public string RoundName { get; set; }
    }

    /// <summary>
    /// Workpaper Status
    //0   Approved
    //1   New
    //3   For Review
    //4   Reviewed With Comments
    //5   Updated
    /// </summary>
    public class WorkpaperStatus
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public int Index { get; set; }
        public string StatusName { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset UpdatedOn { get; set; } = DateTime.Now;
    }

    public class IUCSystemGenCount
    {
        public string roundName { get; set; }
        public int count { get; set; }
    }

    public class IUCNonSystemGenCount
    {
        public string roundName { get; set; }
        public int count { get; set; }
    }

    public class IUCQuestionUserAnswer
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string UniqueId { get; set; }
        public int? FieldId { get; set; }
        public string AppId { get; set; }
        public int Position { get; set; }
        public string Type { get; set; }
        public string Question { get; set; }
        public string Answer { get; set; }
        public string Options { get; set; }
    }

    public class IUCSystemGenAnswer
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string AppId { get; set; }
        public int Position { get; set; }
        public virtual ICollection<IUCQuestionUserAnswer> ListQuestionAnswer { get; set; }
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public string SharefileLink { get; set; }
        public string JsonData { get; set; }

        public DateTimeOffset? CreatedOn { get; set; }

        //[ForeignKey("QuestionnaireRoundSet")]
        //public int QuestionnaireRoundSetId { get; set; }
    }

    public class IUCNonSystemGenAnswer
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string AppId { get; set; }
        public int Position { get; set; }
        public virtual ICollection<IUCQuestionUserAnswer> ListQuestionAnswer { get; set; }
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public string SharefileLink { get; set; }
        public string JsonData { get; set; }

        public DateTimeOffset? CreatedOn { get; set; }


        //[ForeignKey("QuestionnaireRoundSet")]
        //public int QuestionnaireRoundSetId { get; set; }



    }

    public class UpdateElement
    {
        public string elementId { get; set; }
        public string elementValue { get; set; }
    }

    public class Seventeen
    {
        public int Id { get; set; }
        public string SeventeenA { get; set; }
        public string SeventeenB { get; set; }
        public string SeventeenC { get; set; }
        public string SeventeenD { get; set; }
        public string SeventeenE { get; set; }
        public string SeventeenF { get; set; }
    }
}
