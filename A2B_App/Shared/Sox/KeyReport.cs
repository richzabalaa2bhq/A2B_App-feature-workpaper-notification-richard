using BlazorInputFile;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text;
using System.Text.Json.Serialization;


namespace A2B_App.Shared.Sox
{

    #region Key Report

    public class KeyReport
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        [Column(TypeName = "VARCHAR(250)")]
        [MaxLength(250)]
        public string ClientName { get; set; }
        [Column(TypeName = "VARCHAR(60)")]
        [MaxLength(60)]
        public string ClientCode { get; set; }
        public int? ClientItemId { get; set; }

        #region Podio Fields
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
        #endregion

    }
    
    public class ParametersLibrary
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string? ClientName { get; set; }
        public string KeyReportName { get; set; }
        public string Method { get; set; }
        public string Parameter { get; set; }
        public string? A1 { get; set; }
        public string? A2 { get; set; }
        public string? A3 { get; set; }
        public string? A4 { get; set; }
        public string? A5 { get; set; }
        public string? A6 { get; set; }
        public string? A7 { get; set; }
        public string? A8 { get; set; }
        public string? A9 { get; set; }
        public string? A10 { get; set; }
        public string? status { get; set; }
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public DateTimeOffset CreatedOn { get; set; }

    }
    
    public class ReportsLibrary
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string? ClientName { get; set; }
        public string KeyReportName { get; set; }
        public string Method { get; set; }
        public string Report { get; set; }
        public string? B1 { get; set; }
        public string? B2 { get; set; }
        public string? B3 { get; set; }
        public string? B4 { get; set; }
        public string? B5 { get; set; }
        public string? B6 { get; set; }
        public string? B7 { get; set; }
        public string? B8 { get; set; }
        public string? B9 { get; set; }
        public string? B10 { get; set; }
        public string? status { get; set; }
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public DateTimeOffset CreatedOn { get; set; }

    }
    
    public class CompletenessLibrary
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string? ClientName { get; set; }
        public string KeyReportName { get; set; }
        public string Method { get; set; }
        public string Completeness { get; set; }
        public string? C1 { get; set; }
        public string? C2 { get; set; }
        public string? C3 { get; set; }
        public string? C4 { get; set; }
        public string? C5 { get; set; }
        public string? C6 { get; set; }
        public string? C7 { get; set; }
        public string? C8 { get; set; }
        public string? C9 { get; set; }
        public string? C10 { get; set; }
        public string? status { get; set; }
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public DateTimeOffset CreatedOn { get; set; }

    }
    
    public class AccuracyLibrary
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string? ClientName { get; set; }
        public string KeyReportName { get; set; }
        public string Method { get; set; }
        public string Accuracy { get; set; }
        public string? D1 { get; set; }
        public string? D2 { get; set; }
        public string? D3 { get; set; }
        public string? D4 { get; set; }
        public string? D5 { get; set; }
        public string? D6 { get; set; }
        public string? D7 { get; set; }
        public string? D8 { get; set; }
        public string? D9 { get; set; }
        public string? D10 { get; set; }
        public string? status { get; set; }
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public DateTimeOffset CreatedOn { get; set; }

    }
    
    public class CAMethodLibrary
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string MethodType { get; set; }
        public string MethodName { get; set; }
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public DateTimeOffset CreatedOn { get; set; }

    }

    public class KeyReportIds {

        public List<KeyReportUserInput> ListReport { get; set; }
        public int ConsolidatedId { get; set; }
        public int UicId { get; set; }
        public int TestId { get; set; }
        public int ExceptionId { get; set; }

    }

    public class KeyReportOrigFormat
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }

        [Column(TypeName = "VARCHAR(250)")]
        [MaxLength(250)]
        public string ClientName { get; set; }
        [Column(TypeName = "VARCHAR(60)")]
        [MaxLength(60)]
        public string ClientCode { get; set; }
        public int? ClientItemId { get; set; }

        public int Num { get; set; }
        [Column(TypeName = "VARCHAR(250)")]
        [MaxLength(250)]
        public string KeyControlId { get; set; }
        [Column(TypeName = "VARCHAR(250)")]
        [MaxLength(250)]
        public string ControlActivity { get; set; }
        [Column(TypeName = "VARCHAR(60)")]
        [MaxLength(60)]
        public string SKey { get; set; }

        public string NameOfKeyReport { get; set; }
        public string KeyReport { get; set; }
        public string IUCType { get; set; }
        public string Source { get; set; }
        public string ReportCustomized { get; set; }
        public string ControlRelyingIUC { get; set; }
        public string Preparer { get; set; }
        public string Reviewer { get; set; }
        public string Notes { get; set; }

        #region Podio Fields
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }

        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
        #endregion

    }

    public class KeyReportAllIUC
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }

        [Column(TypeName = "VARCHAR(250)")]
        [MaxLength(250)]
        public string ClientName { get; set; }
        [Column(TypeName = "VARCHAR(60)")]
        [MaxLength(60)]
        public string ClientCode { get; set; }
        public int? ClientItemId { get; set; }

        public int Num { get; set; }
        [Column(TypeName = "VARCHAR(250)")]
        [MaxLength(250)]
        public string KeyControlId { get; set; }
        [Column(TypeName = "VARCHAR(250)")]
        [MaxLength(250)]
        public string ControlActivity { get; set; }
        [Column(TypeName = "VARCHAR(60)")]
        [MaxLength(60)]
        public string SKey { get; set; }

        public string NameOfIUC { get; set; }
        public string SourceProcess { get; set; }
        public string KeyReport { get; set; }
        public string UniqueKeyReport { get; set; }
        public string IUCType { get; set; }
        public string Source { get; set; }
        public string ReportCustomized { get; set; }
        public string ControlRelyingIUC { get; set; }
        public string Preparer { get; set; }
        public string Reviewer { get; set; }
        public string AddedToKeyReportTracker { get; set; }
        public string ReportNotes { get; set; }
        public string DescriptionOfKeyReport { get; set; }
        public string KeyReportType { get; set; }
        public string HowIsReportGenerated { get; set; }
        public string HowIsReportUsedToSupControl { get; set; }
        public string StepsPerformedValidateAccuracy { get; set; }
        public string StepsPerformedValidateCompleteness { get; set; }
        public string StepsPerformedValidateSourceData { get; set; }
        public string WhoIsAuthorizedChangeReport { get; set; }
        public string EffectiveDate { get; set; }
        public string WhoHasAccessToModifyReport { get; set; }
        public string WhoHasAccessToRunReport { get; set; }

        public string ReportLastModified { get; set; }
        public string TestedLastModified { get; set; }
        public string ITReportOwner { get; set; }
        public string Questions { get; set; }
        public string FastlyNotesQuestions { get; set; }
        public DateTime? MeetingDate { get; set; }
        public string Process { get; set; }

        #region Podio Fields
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
        #endregion

    }

    public class KeyReportTestStatusTracker
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }

        [Column(TypeName = "VARCHAR(250)")]
        [MaxLength(250)]
        public string ClientName { get; set; }
        [Column(TypeName = "VARCHAR(60)")]
        [MaxLength(60)]
        public string ClientCode { get; set; }
        public int? ClientItemId { get; set; }
        public int Num { get; set; }
        public string NameOfIUC { get; set; }
        public string Process { get; set; }
        [Column(TypeName = "VARCHAR(250)")]
        [MaxLength(250)]
        public string KeyControlId { get; set; }
        [Column(TypeName = "VARCHAR(250)")]
        [MaxLength(250)]
        public string ControlActivity { get; set; }
        [Column(TypeName = "VARCHAR(60)")]
        [MaxLength(60)]
        public string SKey { get; set; }
        public string KeyReport { get; set; }
        public string UniqueKeyReport { get; set; }
        public string Preparer { get; set; }
        public string ProcessOwner { get; set; }
        public string KeyReportOwner { get; set; }
        public string KeyReportITOwner { get; set; }
        public string SetupReportLeadSheetTesting { get; set; }
        public DateTime? ScheduleProcessOwnerMeeting { get; set; }
        public string ReportReceived { get; set; }
        public string PBCStatus { get; set; }
        public string Tester { get; set; }
        public string FirstReviewer { get; set; }
        public string SecondReviewer { get; set; }
        public string TestingStatus { get; set; }
        public string A2Q2DueDate { get; set; }
        public string SentToFastly { get; set; }
        public string FastlyReviewStatus { get; set; }
        public string SentToDeloitte { get; set; }
        public string A2Q2Notes { get; set; }

        #region Podio Fields
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        
        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
        #endregion

    }

    public class KeyReportExcepcion
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }

        [Column(TypeName = "VARCHAR(250)")]
        [MaxLength(250)]
        public string ClientName { get; set; }
        [Column(TypeName = "VARCHAR(60)")]
        [MaxLength(60)]
        public string ClientCode { get; set; }
        public int? ClientItemId { get; set; }

        public int Num { get; set; }
        public string NameOfIUC { get; set; }
        [Column(TypeName = "VARCHAR(250)")]
        public string KeyControlId { get; set; }
        [Column(TypeName = "VARCHAR(250)")]
        public string ControlActivity { get; set; }
        [Column(TypeName = "VARCHAR(60)")]
        public string ExceptionNoted { get; set; }
        public string ReasonForException { get; set; }
        public string Remediation { get; set; }

        #region Podio Fields
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        
        public DateTimeOffset? CreatedOn { get; set; } = DateTime.Now;
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
        #endregion

    }

    public class KeyReportFilter
    {
        public string FY { get; set; }
        public string ClientName { get; set; }
        public string KeyReportName { get; set; }
        public string ControlId { get; set; }
        public string UniqueKey { get; set; }

    }

    public class KeyReportFileProperties
    {
        public string FY { get; set; }
        public string ClientName { get; set; }
        public string LoadingStatus { get; set; }
        public string SharefileLink { get; set; }
        public string KeyReportName { get; set; }
        public string FileName { get; set; }
    }

    public class KeyReportScreenshot
    {
        public KeyReportFilter Filter { get; set; }
        public List<string> ListScreenshotName { get; set; }
        public string ReportFilename { get; set; }
    }

    public class KeyReportItemId 
    { 
        public int OrigFormatItemId { get; set; }
        public int AllIUCItemId { get; set; }
        public int TestItemId { get; set; }
        public int ExceptionItemId { get; set; }

    }

    public class KeyReportQuestionsFilter
    {
        public int method { get; set; }
        public string clientName { get; set; }
        public string reportName { get; set; }
    }

    public class ViewAccessOrigFormat
    {
        public int no { get; set; }
        public string keyControl { get; set; }
        public string controlActivity { get; set; }
        public string keyNonKeyControl { get; set; }
        public string nameIUC { get; set; }
        public string sourceProcess { get; set; }
        public string keyReport { get; set; }
        public string iucType { get; set; }
        public string systemSource { get; set; }
        public string reportCustomized { get; set; }
        public string controlRelyingIUC { get; set; }
        public string preparer { get; set; }
        public string reviewer { get; set; }
        public string notes { get; set; }
    }
    
    public class ViewAccessAllIuc
    {
        public int no { get; set; }
        public string keyControl { get; set; }
        public string controlActivity { get; set; }
        public string keyNonKeyControl { get; set; }
        public string nameIUC { get; set; }
        public string sourceProcess { get; set; }
        public string keyReport { get; set; }
        public string uniqueKeyReport { get; set; }
        
        public string iucType { get; set; }
        public string systemSource { get; set; }
        public string reportCustomized { get; set; }
        public string controlRelyingIUC { get; set; }
        public string preparer { get; set; }
        public string reviewer { get; set; }
        public string addedToKeyReportTracker { get; set; }
        public string reportNotes { get; set; }
        public string questions { get; set; }
        public string fastlyNotesAndQuestions { get; set; }
        public string meetingDate { get; set; }
        public string process { get; set; }

        //Business Owners
        public string descriptionKeyReport { get; set; }
        public string keyReportType { get; set; }
        public string howReportGenerated { get; set; }
        public string howReportUsed { get; set; }
        public string stepsPerformedAccuracy { get; set; }
        public string stepsPerformedCompleteness { get; set; }
        public string stepsPerformedValidateSource { get; set; }
        public string areParameters { get; set; }
        public string whoIsAuthorized { get; set; }
        public string effectiveDate { get; set; }
        //Information Technology
        public string whoHasAccessToEdit { get; set; }
        public string whoHasAccessToRun { get; set; }
        public string reportLastModified { get; set; }
        public string howWasItTestedWhenLastModified { get; set; }
        public string itReportOwner { get; set; }



    }
    
    public class ViewAccessTestStatus
    {
        public int no { get; set; }
        public string nameIuc { get; set; }
        public string process { get; set; }
        public string keyControl { get; set; }
        public string controlActvity { get; set; }
        public string keyNonKeyControl { get; set; }
        public string keyReport { get; set; }
        public string uniqueKeyReport { get; set; }
        public string preparer { get; set; }
        public string processOwner { get; set; }
        public string keyReportOwner { get; set; }
        public string keyReportITOwner { get; set; }
        public string setupLeadsheetForTesting { get; set; }
        public string scheduleProcessOwnerMeeting { get; set; }
        public string reportReceived { get; set; }
        public string pbcStatus { get; set; }
        public string tester { get; set; }
        public string firstReviewer { get; set; }
        public string secondReviewer { get; set; }
        public string testingStatus { get; set; }
        public string a2q2DueDateTesting { get; set; }
        public string sentToClient { get; set; }
        public string clientReviewStatus { get; set; }
        public string sentToDeloitte { get; set; }
        public string a2q2Notes { get; set; }
    }
    
    public class ViewAccessListOfExceptions
    {
        public int no { get; set; }
        public string nameIuc { get; set; }
        public string keyControl { get; set; }
        public string controlActvity { get; set; }
        public string exceptionNoted { get; set; }
        public string reasonForException { get; set; }
        public string remediation { get; set; }
    }




    #endregion

    #region Key Report Workpaper

    public class KeyReportQuestion
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        //public string ClientName { get; set; }
        public string AppId { get; set; }
        //public string ControlName { get; set; }
        public string QuestionString { get; set; }
        public string Type { get; set; }
        public int Position { get; set; }
        public string Description { get; set; }
        public string Tag { get; set; }
        public int? FieldId { get; set; }
        public virtual ICollection<KeyReportOption> Options { get; set; }

        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset UpdatedOn { get; set; } = DateTime.Now;
    }

    public class KeyReportOption
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string OptionName { get; set; }
        public string OptionId { get; set; }
        public int AppId { get; set; }
        public KeyReportQuestion KeyReportQuestion { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset UpdatedOn { get; set; } = DateTime.Now;
    }

    public class KeyReportUserInput
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
        [Column(TypeName = "VARCHAR(100)")]
        public string AppId { get; set; }
        public int? FieldId { get; set; }
        public int ItemId { get; set; }
        [Column(TypeName = "VARCHAR(100)")]
        public string Type { get; set; }
        public string Tag { get; set; }


        //Updated Model 3-11-21
        public string Link { get; set; }
        public string TagFY { get; set; }
        public string TagClientName { get; set; }
        public string TagReportName { get; set; }
        public string TagControlId { get; set; }
        public string TagStatus { get; set; }
        public DateTimeOffset? CreatedOn { get; set; } = DateTime.Now;
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset UpdatedOn { get; set; } = DateTime.Now;

    }

    public class KeyReportReturnAnswer
    {
        public List<KeyReportUserInput> ListKeyReportUserInput { get; set; }
        public int Position { get; set; }
    }
    #endregion

    #region Key Report Reference Apps
    public class KeyReportControlActivity
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string ControlActivity { get; set; } 

        #region Podio Fields
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
        #endregion

    }

    public class KeyReportKeyControl
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string Key { get; set; }

        #region Podio Fields
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
        #endregion

    }

    public class KeyReportName
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string Name { get; set; }

        #region Podio Fields
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
        #endregion

    }

    public class KeyReportSystemSource
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string SystemSource { get; set; }
        public string SystemSourceCategory { get; set; }

        #region Podio Fields
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
        #endregion

    }

    public class KeyReportNonKeyReport
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string Report { get; set; }

        #region Podio Fields
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
        #endregion
    }

    public class KeyReportReportCustomized
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string ReportCustomized { get; set; }

        #region Podio Fields
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
        #endregion

    }

    public class KeyReportIUCType
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string IUCType { get; set; }

        #region Podio Fields
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
        #endregion

    }

    public class KeyReportControlsRelyingIUC
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string ControlsRelyingIUC { get; set; }

        #region Podio Fields
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
        #endregion

    }

    public class KeyReportPreparer
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string Preparer { get; set; }

        #region Podio Fields
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
        #endregion
    }

    public class KeyReportUniqueKeyReport
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string UniqueKeyReport { get; set; }

        #region Podio Fields
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
        #endregion
    }

    public class KeyReportNotes
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string ReportNotes { get; set; }

        #region Podio Fields
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
        #endregion
    }

    public class KeyReportNumber
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string ReportNumber { get; set; }

        #region Podio Fields
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
        #endregion
    }

    public class KeyReportTester
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string Tester { get; set; }

        #region Podio Fields
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
        #endregion
    }

    public class KeyReportReviewer
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [JsonIgnore]
        public int Id { get; set; }
        public string Reviewer { get; set; }

        #region Podio Fields
        public int PodioItemId { get; set; }
        public string PodioUniqueId { get; set; }
        public int PodioRevision { get; set; }
        public string PodioLink { get; set; }
        public string CreatedBy { get; set; }
        public DateTimeOffset? CreatedOn { get; set; }
        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        public DateTimeOffset LastUpdate { get; set; } = DateTime.Now;
        #endregion
    }


    #endregion

    #region Key Report File Upload
    public class KeyReportScreenshotUpload
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Newtonsoft.Json.JsonIgnore]
        public int Id { get; set; }
        [NotMapped]
        [Newtonsoft.Json.JsonIgnore]
        public IFileListEntry IFileEntry { get; set; }        
        public int Position { get; set; }
        public string Status { get; set; }
        public string Filename { get; set; }
        public string NewFilename { get; set; }
        public string Percent { get; set; }
        public bool IsLoading { get; set; }
        public string Client { get; set; }
        public string Fy { get; set; }
        public string ReportName { get; set; }
        public string ControlId { get; set; }
    }

    public class KeyReportFile
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Newtonsoft.Json.JsonIgnore]
        public int Id { get; set; }
        [NotMapped]
        [Newtonsoft.Json.JsonIgnore]
        public IFileListEntry IFileEntry { get; set; }
        public string Status { get; set; }
        public string Percent { get; set; }
        public bool IsLoading { get; set; }
        public string Filename { get; set; }
        public string NewFilename { get; set; }
        public string Client { get; set; }
        public string Fy { get; set; }
        public string ReportName { get; set; }
        public string ControlId { get; set; }
    }


    #endregion

}
