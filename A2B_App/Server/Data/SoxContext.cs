using A2B_App.Shared.Podio;
using A2B_App.Shared.Sox;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Design;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.IO;


namespace A2B_App.Server.Data
{
    public class SoxContext : DbContext
    {

        public class SampleSelectionContextFactory : IDesignTimeDbContextFactory<SoxContext>
        {
            private IConfiguration _configuration;

            public SampleSelectionContextFactory()
            {
                var builder = new ConfigurationBuilder()
                        .SetBasePath(Directory.GetCurrentDirectory())
                        .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

                _configuration = builder.Build();
            }

            public SoxContext CreateDbContext(string[] args)
            {
                var optionsBuilder = new DbContextOptionsBuilder<SoxContext>();
                optionsBuilder.UseMySQL(_configuration.GetConnectionString("SampleSelectionCon"), opts => opts.CommandTimeout((int)TimeSpan.FromMinutes(30).TotalSeconds));

                return new SoxContext(optionsBuilder.Options);
            }

        }

        public SoxContext(DbContextOptions<SoxContext> options)
        : base(options)
        {
            Database.SetCommandTimeout(600);
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {

            #region SampleSelection

            modelBuilder.Entity<SampleSelection>(e =>
            {
                e.HasAlternateKey(x => x.PodioItemId);
                e.HasMany(x => x.ListTestRound);
                e.HasMany(x => x.ListRefId);
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.PodioItemId });

            });

            //modelBuilder.Entity<TestRoundSampleSelectionReference>().HasIndex(x => x.PodioItemID).IsUnique();
            modelBuilder.Entity<TestRoundSampleSelectionReference>(e =>
            {
                e.HasIndex(x => x.PodioItemID).IsUnique();
                e.Property(x => x.Id).ValueGeneratedOnAdd();
            });

            modelBuilder.Entity<TestRound>(e =>
            {
                
                e.HasIndex(x => x.PodioItemId).IsUnique();
                e.HasIndex(x => x.Id);
                e.Property(x => x.Id).ValueGeneratedOnAdd();

            });

            modelBuilder.Entity<PodioHook>(e => {
                e.HasIndex(x => x.Id);
                e.Property(x => x.Id).ValueGeneratedOnAdd();
            });

            modelBuilder.Entity<SampleRound>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
            });

            modelBuilder.Entity<ClientSs>(e =>
            {
                e.HasIndex(x => x.PodioItemId).IsUnique();
                e.HasIndex(x => x.Id);
                e.Property(x => x.Id).ValueGeneratedOnAdd();
            });

            modelBuilder.Entity<Matrix>(e =>
            {
                e.HasIndex(x => x.PodioItemId).IsUnique();
                e.HasIndex(x => x.Id);
                e.Property(x => x.Id).ValueGeneratedOnAdd();
            });

            #endregion

            #region RCM

            modelBuilder.Entity<RcmCta>().HasIndex(x => x.PodioItemId).IsUnique();


            modelBuilder.Entity<Rcm>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.PodioItemId, x.ClientItemId, x.ClientName, x.ClientCode, x.ControlId });
                e.Property(x => x.ClientName).HasMaxLength(250);
                e.Property(x => x.ClientCode).HasMaxLength(250);
                e.Property(x => x.ControlId).HasMaxLength(250);
            });

            modelBuilder.Entity<RcmQuestionnaireField>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasMany(x => x.Options);
                e.HasIndex(x => new { x.Id, x.AppId, x.Position, x.FieldId });
                e.Property(x => x.Position).HasMaxLength(250);
                e.Property(x => x.AppId).HasMaxLength(250);
      
            });

            modelBuilder.Entity<RcmQuestionnaireFieldOptions>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.OptionId, x.AppId});
                e.Property(x => x.OptionId).HasMaxLength(250);

            });

            modelBuilder.Entity<RcmQuestionnaire>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasMany(x => x.ListQ1Year);
                e.HasMany(x => x.ListQ7FinStatementElement);
                e.HasMany(x => x.ListQ8FinStatementAssert);
                e.HasMany(x => x.ListQ9ControlId);
                e.HasMany(x => x.ListQ12ControlOwner);
                e.HasMany(x => x.ListQ13Frequency);
                e.HasMany(x => x.ListQ14ControlKey);
                e.HasMany(x => x.ListQ15ControlType);
                e.HasMany(x => x.ListQ16NatureProcedure);
                e.HasMany(x => x.ListQ17FraudControl);
                e.HasMany(x => x.ListQ18RiskLevel);
                e.HasMany(x => x.ListQ19MgmtReviewControl);
                e.HasIndex(x => new { x.Id, x.PodioItemId, x.Q1FieldId, x.Q2FieldId, x.Q3FieldId, x.Q4FieldId, x.Q9FieldId });

            });

            modelBuilder.Entity<RcmFY>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id });
            });

            modelBuilder.Entity<RcmProcess>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.PodioItemId });
            });

            modelBuilder.Entity<RcmSubProcess>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.PodioItemId });
            });

            modelBuilder.Entity<RcmFinancialStatement>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.PodioItemId });
            });

            modelBuilder.Entity<RcmFinancialStatementAssert>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id });
            });

            modelBuilder.Entity<RcmControlId>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.PodioItemId });
            });

            modelBuilder.Entity<RcmControlOwner>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.PodioItemId });
            });

            modelBuilder.Entity<RcmFrequency>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.PodioItemId });
            });

            modelBuilder.Entity<RcmControlKey>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id });
            });

            modelBuilder.Entity<RcmFraudControl>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id });
            });

            modelBuilder.Entity<RcmRiskLevel>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id});
            });

            modelBuilder.Entity<RcmControlType>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.PodioItemId });
            });

            modelBuilder.Entity<RcmNatureProcedure>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.PodioItemId });
            });

            modelBuilder.Entity<RcmManagementReviewControl>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.PodioItemId });
            });

            #endregion

            #region WorkPaper

            modelBuilder.Entity<Questionnaire>(e =>
            {
                e.HasIndex(x => x.PodioItemId).IsUnique();
                e.HasMany(x => x.ListNotes);
                e.HasMany(x => x.ListSampleRound);
                e.Property(x => x.Id).ValueGeneratedOnAdd();
            });

            modelBuilder.Entity<NotesUnique>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => x.Id);
            });

            modelBuilder.Entity<QuestionnaireQuestion>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.Property(x => x.ClientName).HasMaxLength(250);
                e.Property(x => x.AppId).HasMaxLength(250);
                e.HasIndex(x => new { x.Id, x.ClientName, x.AppId, x.FieldId });
            });

            modelBuilder.Entity<QuestionnaireOption>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.AppId });
            });

            modelBuilder.Entity<QuestionnaireUserInput>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasMany(x => x.ListRoundItem);
                e.HasIndex(x => new { x.Id, x.AppId, x.FieldId, x.Type });
                e.Property(x => x.Type).HasMaxLength(250);
                e.Property(x => x.AppId).HasMaxLength(250);
            });

            modelBuilder.Entity<PodioAppKey>(e =>
            {
                e.HasIndex(x => x.AppId).IsUnique();
                e.HasIndex(x => x.Id);
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.Property(x => x.AppId).HasMaxLength(250);
            });

            modelBuilder.Entity<IUCSystemGen>(e =>
            {
                e.HasIndex(x => new { x.Id, x.AppId, x.PodioItemId, x.PodioUniqueId });
                e.HasMany(x => x.ListQuestionAnswer);
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.Property(x => x.AppId).HasMaxLength(250);
                e.Property(x => x.PodioUniqueId).HasMaxLength(250);

            });

            modelBuilder.Entity<IUCNonSystemGen>(e =>
            {
                e.HasIndex(x => new { x.Id, x.AppId, x.PodioItemId, x.PodioUniqueId });
                e.HasMany(x => x.ListQuestionAnswer);
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.Property(x => x.AppId).HasMaxLength(250);
                e.Property(x => x.PodioUniqueId).HasMaxLength(250);

            });

            modelBuilder.Entity<QandA>(e =>
            {
                //e.HasIndex(x => new { x.Id, x.FieldId, x.Type }).IsUnique();
                e.HasIndex(x => new { x.Id, x.FieldId, x.Type });
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.Property(x => x.Type).HasMaxLength(250);

            });

            modelBuilder.Entity<ClientControl>(e =>
            {
                e.HasIndex(x => new { x.ClientName, x.ControlName });
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.Property(x => x.ClientName).HasMaxLength(250);
                e.Property(x => x.ControlName).HasMaxLength(250);
            });

            modelBuilder.Entity<QuestionnaireFieldParam>(e =>
            {
                e.HasIndex(x => new { x.ClientName, x.ControlName });
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.Property(x => x.ClientName).HasMaxLength(250);
                e.Property(x => x.ControlName).HasMaxLength(250);
            });

            modelBuilder.Entity<QuestionnaireUserAnswer>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.Property(x => x.RoundName).HasMaxLength(256);
                e.HasIndex(x => new { x.Id, x.AppId, x.FieldId, x.Type, x.RoundName });
                e.Property(x => x.AppId).HasMaxLength(250);
                e.Property(x => x.Type).HasMaxLength(250);
                e.Property(x => x.RoundName).HasMaxLength(250);
                e.HasMany(x => x.Options);

            });

            modelBuilder.Entity<NotesItem>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.PodioItemId });
            });

            modelBuilder.Entity<NotesItem2>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.PodioItemId });
            });

            modelBuilder.Entity<GeneralNote>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id });
            });

            modelBuilder.Entity<IPENote>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id });
            });

            modelBuilder.Entity<RoundItem>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.Property(x => x.RoundName).HasMaxLength(256);
                e.HasIndex(x => new { x.Id, x.PodioItemId, x.AppId, x.RoundName });
                e.Property(x => x.AppId).HasMaxLength(250);
                e.Property(x => x.RoundName).HasMaxLength(250);
            });

            modelBuilder.Entity<RoundItem2>(e =>
            {
                e.HasKey(x => x.Id);
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.Property(x => x.RoundName).HasMaxLength(256);
                e.HasIndex(x => new { x.Id, x.PodioItemId, x.AppId, x.RoundName });
                e.HasMany(x => x.ListRoundQA);
                e.Property(x => x.AppId).HasMaxLength(250);
                e.Property(x => x.RoundName).HasMaxLength(250);
            });

            modelBuilder.Entity<RoundQA2>(e =>
            {
                e.HasKey(x => x.Id);
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.Position});
                e.HasMany(x => x.Options);
            });
            

            modelBuilder.Entity<QuestionnaireRoundSet>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasMany(x => x.ListUserInputRound1);
                e.HasMany(x => x.ListUserInputRound2);
                e.HasMany(x => x.ListUserInputRound3);
                e.HasMany(x => x.ListUniqueNotes);
                e.HasMany(x => x.ListRoundItem);
                e.HasMany(x => x.ListIUCSystemGen1);
                e.HasMany(x => x.ListIUCSystemGen2);
                e.HasMany(x => x.ListIUCSystemGen3);
                e.HasMany(x => x.ListIUCNonSystemGen1);
                e.HasMany(x => x.ListIUCNonSystemGen2);
                e.HasMany(x => x.ListIUCNonSystemGen3);
                e.HasIndex(x => new { x.Id, x.UniqueId });
                e.Property(x => x.UniqueId).HasMaxLength(250);

            });
          
            modelBuilder.Entity<IUCQuestionUserAnswer>(e =>
            {
                e.HasIndex(x => new { x.Id, x.FieldId, x.Type }).IsUnique();
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.Property(x => x.Type).HasMaxLength(250);
            });

            modelBuilder.Entity<IUCSystemGenAnswer>(e =>
            {
                e.HasIndex(x => new { x.Id, x.AppId, x.PodioItemId, x.PodioUniqueId });
                e.HasMany(x => x.ListQuestionAnswer);
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.Property(x => x.AppId).HasMaxLength(250);
                e.Property(x => x.PodioUniqueId).HasMaxLength(250);
                //e.HasOne(x => x.QuestionnaireRoundSet);
                //e.HasOne(x => x.QuestionnaireRoundSet2);
                //e.HasOne(x => x.QuestionnaireRoundSet3);

            });

            modelBuilder.Entity<IUCNonSystemGenAnswer>(e =>
            {
                e.HasIndex(x => new { x.Id, x.AppId, x.PodioItemId, x.PodioUniqueId });
                e.HasMany(x => x.ListQuestionAnswer);
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.Property(x => x.AppId).HasMaxLength(250);
                e.Property(x => x.PodioUniqueId).HasMaxLength(250);
                //e.HasOne(x => x.QuestionnaireRoundSet);
                //e.HasOne(x => x.QuestionnaireRoundSet2);
                //e.HasOne(x => x.QuestionnaireRoundSet3);

            });

            modelBuilder.Entity<HeaderNote>(e =>
            {
                e.HasIndex(x => new { x.Id});
                e.Property(x => x.Id).ValueGeneratedOnAdd();

            });

            //modelBuilder.Entity<HeaderNote2>(e =>
            //{
            //    e.HasIndex(x => new { x.Id });
            //    e.Property(x => x.Id).ValueGeneratedOnAdd();
            //    e.HasMany(x => x.ListSevenTeen);
            //});

            modelBuilder.Entity<QuestionnaireTesterSet>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasMany(x => x.ListUserInputRound);
                e.HasMany(x => x.ListUniqueNotes);
                e.HasMany(x => x.ListRoundItem2);
                e.HasMany(x => x.ListIUCSystemGen);
                e.HasMany(x => x.ListIUCNonSystemGen);
                e.Property(x => x.UniqueId).HasMaxLength(250);
                e.Property(x => x.RoundName).HasMaxLength(100);
                e.HasIndex(x => new { x.Id, x.UniqueId, x.RoundName, x.DraftNum });
                
            });

            modelBuilder.Entity<QuestionnaireReviewerSet>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                //e.HasMany(x => x.ListUserInputRound);
                //e.HasMany(x => x.ListUniqueNotes);
                //e.HasMany(x => x.ListRoundItem);
                //e.HasMany(x => x.ListIUCSystemGen);
                //e.HasMany(x => x.ListIUCNonSystemGen);
                e.Property(x => x.UniqueId).HasMaxLength(250);
                e.Property(x => x.RoundName).HasMaxLength(100);
                e.HasIndex(x => new { x.Id, x.UniqueId, x.RoundName, x.DraftNum });
                
            });

            modelBuilder.Entity<WorkpaperStatus>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.Property(x => x.StatusName).HasMaxLength(100);
                e.HasIndex(x => new { x.Id, x.Index, x.StatusName });

            });

            #endregion

            #region SoxTracker

            modelBuilder.Entity<SoxTracker>(e =>
            {
                e.HasIndex(x => new { x.Id, x.ClientName, x.ControlId, x.PodioItemId, x.PodioUniqueId });
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.Property(x => x.ClientName).HasMaxLength(250);
                e.Property(x => x.ControlId).HasMaxLength(250);
                e.Property(x => x.PodioUniqueId).HasMaxLength(250);

            });

            modelBuilder.Entity<SoxTrackerQuestionnaire>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.PodioItemId });
            });

            modelBuilder.Entity<SoxTrackerQuestionnaire>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasMany(x => x.ListSoxTrackerAppCategory);
                e.HasMany(x => x.ListSoxTrackerAppRelationship);
                e.HasIndex(x => new { x.Id, x.PodioItemId });
            });

            modelBuilder.Entity<SoxTrackerAppRelationship>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.PodioItemId, x.FieldId });
            });

            modelBuilder.Entity<SoxTrackerAppCategory>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.FieldId });
            });

            #endregion

            #region KeyReport

            modelBuilder.Entity<KeyReport>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.ClientName, x.ClientItemId, x.PodioItemId });
                e.Property(x => x.ClientName).HasMaxLength(250);

            });

            modelBuilder.Entity<KeyReportOrigFormat>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { 
                    x.Id, 
                    x.ClientName, 
                    x.ClientItemId, 
                    x.KeyControlId,
                    x.ControlActivity,
                    x.PodioItemId 
                });
                e.Property(x => x.ClientName).HasMaxLength(250);
                e.Property(x => x.KeyControlId).HasMaxLength(250);
                e.Property(x => x.ControlActivity).HasMaxLength(700);
            });

            modelBuilder.Entity<KeyReportAllIUC>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new {
                    x.Id,
                    x.ClientName,
                    x.ClientItemId,
                    x.KeyControlId,
                    x.ControlActivity,
                    x.PodioItemId
                });
                e.Property(x => x.ClientName).HasMaxLength(250);
                e.Property(x => x.KeyControlId).HasMaxLength(250);
                e.Property(x => x.ControlActivity).HasMaxLength(700);
            });

            modelBuilder.Entity<KeyReportTestStatusTracker>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new {
                    x.Id,
                    x.ClientName,
                    x.ClientItemId,
                    x.KeyControlId,
                    x.ControlActivity,
                    x.PodioItemId
                });
                e.Property(x => x.ClientName).HasMaxLength(250);
                e.Property(x => x.KeyControlId).HasMaxLength(250);
                e.Property(x => x.ControlActivity).HasMaxLength(700);
            });

            modelBuilder.Entity<KeyReportExcepcion>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new {
                    x.Id,
                    x.ClientName,
                    x.ClientItemId,
                    x.KeyControlId,
                    x.ControlActivity,
                    x.PodioItemId
                });
                e.Property(x => x.ClientName).HasMaxLength(250);
                e.Property(x => x.KeyControlId).HasMaxLength(250);
                e.Property(x => x.ControlActivity).HasMaxLength(700);
            });

            modelBuilder.Entity<KeyReportQuestion>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.AppId, x.FieldId });
                e.Property(x => x.AppId).HasMaxLength(250);

            });

            modelBuilder.Entity<KeyReportOption>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.AppId });
            });

            modelBuilder.Entity<KeyReportUserInput>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.AppId, x.FieldId, x.Type, x.ItemId });
                e.Property(x => x.AppId).HasMaxLength(250);
                e.Property(x => x.Type).HasMaxLength(250);
            });

            modelBuilder.Entity<KeyReportControlActivity>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.PodioItemId });
            });

            modelBuilder.Entity<KeyReportKeyControl>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.PodioItemId });
            });

            modelBuilder.Entity<KeyReportName>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.PodioItemId });
            });

            modelBuilder.Entity<KeyReportSystemSource>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.PodioItemId });
            });

            modelBuilder.Entity<KeyReportNonKeyReport>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.PodioItemId });
            });

            modelBuilder.Entity<KeyReportReportCustomized>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.PodioItemId });
            });

            modelBuilder.Entity<KeyReportIUCType>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.PodioItemId });
            });

            modelBuilder.Entity<KeyReportControlsRelyingIUC>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.PodioItemId });
            });

            modelBuilder.Entity<KeyReportPreparer>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.PodioItemId });
            });

            modelBuilder.Entity<KeyReportUniqueKeyReport>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.PodioItemId });
            });

            modelBuilder.Entity<KeyReportNotes>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.PodioItemId });
            });

            modelBuilder.Entity<KeyReportNumber>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.PodioItemId });
            });

            modelBuilder.Entity<KeyReportTester>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.PodioItemId });
            });

            modelBuilder.Entity<KeyReportReviewer>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.PodioItemId });
            });
            modelBuilder.Entity<ParametersLibrary>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.PodioItemId });
            });
            modelBuilder.Entity<ReportsLibrary>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.PodioItemId });
            });
            modelBuilder.Entity<CompletenessLibrary>(e =>
            {
                    e.Property(x => x.Id).ValueGeneratedOnAdd();
                    e.HasIndex(x => new { x.Id, x.PodioItemId });
            });
            modelBuilder.Entity<AccuracyLibrary>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.PodioItemId });
            });
            modelBuilder.Entity<CAMethodLibrary>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.PodioItemId });
            });

            modelBuilder.Entity<KeyReportScreenshotUpload>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.Property(x => x.Client).HasMaxLength(250);
                e.Property(x => x.Fy).HasMaxLength(250);
                e.Property(x => x.ReportName).HasMaxLength(250);
                e.Property(x => x.ControlId).HasMaxLength(250);
                e.HasIndex(x => new { x.Id  });
            });

            modelBuilder.Entity<KeyReportFile>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.Property(x => x.Client).HasMaxLength(250);
                e.Property(x => x.Fy).HasMaxLength(250);
                e.Property(x => x.ReportName).HasMaxLength(250);
                e.Property(x => x.ControlId).HasMaxLength(250);
                e.HasIndex(x => new { x.Id });
            });

            #endregion

        }

        #region SampleSelection
        public DbSet<SampleSelection> SampleSelection { get; set; }
        public DbSet<TestRound> TestRounds { get; set; }
        public DbSet<PodioHook> PodioHook { get; set; }
        public DbSet<TestRoundSampleSelectionReference> TestRoundSampleSelectionReference { get; set; }
        #endregion

        #region RCM
        public DbSet<RcmCta> RcmCta { get; set; }
        public DbSet<Rcm> Rcm { get; set; }
        public DbSet<RcmQuestionnaireField> RcmQuestionnaireField { get; set; }
        public DbSet<RcmQuestionnaireFieldOptions> RcmQuestionnaireFieldOptions { get; set; }
        public DbSet<RcmQuestionnaire> RcmQuestionnaire { get; set; }
        public DbSet<RcmFY> RcmFY { get; set; }
        public DbSet<RcmProcess> RcmProcess { get; set; }
        public DbSet<RcmSubProcess> RcmSubProcess { get; set; }
        public DbSet<RcmFinancialStatement> RcmFinancialStatement { get; set; }
        public DbSet<RcmFinancialStatementAssert> RcmFinancialStatementAssert { get; set; }
        public DbSet<RcmControlId> RcmControlId { get; set; }
        public DbSet<RcmControlOwner> RcmControlOwner { get; set; }
        public DbSet<RcmFrequency> RcmFrequency { get; set; }
        public DbSet<RcmControlKey> RcmControlKey { get; set; }
        public DbSet<RcmFraudControl> RcmFraudControl { get; set; }
        public DbSet<RcmRiskLevel> RcmRiskLevel { get; set; }
        public DbSet<RcmControlType> RcmControlType { get; set; }
        public DbSet<RcmNatureProcedure> RcmNatureProcedure { get; set; }
        public DbSet<RcmManagementReviewControl> RcmManagementReviewControl { get; set; }

        #endregion

        #region WorkPaper
        public DbSet<Questionnaire> Questionnaire { get; set; }
        public DbSet<NotesUnique> NotesUnique { get; set; }
        public DbSet<SampleRound> SampleRound { get; set; }
        public DbSet<ClientSs> ClientSs { get; set; }
        public DbSet<Matrix> Matrix { get; set; }
        public DbSet<QuestionnaireQuestion> QuestionnaireQuestion { get; set; }
        public DbSet<QuestionnaireOption> QuestionnaireOption { get; set; }
        public DbSet<QuestionnaireUserInput> QuestionnaireUserInput { get; set; }
        public DbSet<PodioAppKey> PodioAppKey { get; set; }
        public DbSet<IUCSystemGen> IUCSystemGen { get; set; }
        public DbSet<IUCNonSystemGen> IUCNonSystemGen { get; set; }
        public DbSet<QandA> QandA { get; set; }
        public DbSet<ClientControl> ClientControl { get; set; }
        public DbSet<QuestionnaireFieldParam> QuestionnaireFieldParam { get; set; }
        public DbSet<QuestionnaireRoundSet> QuestionnaireRoundSet { get; set; }
        public DbSet<QuestionnaireUserAnswer> QuestionnaireUserAnswer { get; set; }
        public DbSet<IUCQuestionUserAnswer> IUCQuestionUserAnswer { get; set; }
        public DbSet<IUCSystemGenAnswer> IUCSystemGenAnswer { get; set; }
        public DbSet<IUCNonSystemGenAnswer> IUCNonSystemGenAnswer { get; set; }
        public DbSet<NotesItem> NotesItem { get; set; }
        public DbSet<NotesItem> NotesItem2 { get; set; }
        public DbSet<RoundItem> RoundItem { get; set; }
        public DbSet<RoundItem2> RoundItem2 { get; set; }
        public DbSet<RoundQA2> RoundQA2 { get; set; }
        public DbSet<HeaderNote> HeaderNote { get; set; }
        public DbSet<HeaderNote2> HeaderNote2 { get; set; }
        public DbSet<QuestionnaireTesterSet> QuestionnaireTesterSet { get; set; }
        public DbSet<QuestionnaireReviewerSet> QuestionnaireReviewerSet { get; set; }
        public DbSet<WorkpaperStatus> WorkpaperStatus { get; set; }
        public DbSet<SampleSelectionProperties> SampleSelectionProperties { get; set; }
        #endregion

        #region SoxTracker
        public DbSet<SoxTracker> SoxTracker { get; set; }
        public DbSet<SoxTrackerQuestionnaire> SoxTrackerQuestionnaire { get; set; }
        public DbSet<SoxTrackerAppRelationship> SoxTrackerAppRelationship { get; set; }
        public DbSet<SoxTrackerAppCategory> SoxTrackerAppCategory { get; set; }
        #endregion

        #region KeyReport
        public DbSet<KeyReport> KeyReport { get; set; }
        public DbSet<KeyReportOrigFormat> KeyReportOrigFormat { get; set; }
        public DbSet<KeyReportAllIUC> KeyReportAllIUC { get; set; }
        public DbSet<KeyReportTestStatusTracker> KeyReportTestStatusTracker { get; set; }
        public DbSet<KeyReportExcepcion> KeyReportExcepcion { get; set; }
        public DbSet<KeyReportQuestion> KeyReportQuestion { get; set; }
        public DbSet<KeyReportOption> KeyReportOption { get; set; }
        public DbSet<KeyReportUserInput> KeyReportUserInput { get; set; }


        public DbSet<KeyReportControlActivity> KeyReportControlActivity { get; set; }
        public DbSet<KeyReportKeyControl> KeyReportKeyControl { get; set; }
        public DbSet<KeyReportName> KeyReportName { get; set; }
        public DbSet<KeyReportSystemSource> KeyReportSystemSource { get; set; }
        public DbSet<KeyReportNonKeyReport> KeyReportNonKeyReport { get; set; }
        public DbSet<KeyReportReportCustomized> KeyReportReportCustomized { get; set; }
        public DbSet<KeyReportIUCType> KeyReportIUCType { get; set; }
        public DbSet<KeyReportControlsRelyingIUC> KeyReportControlsRelyingIUC { get; set; }
        public DbSet<KeyReportPreparer> KeyReportPreparer { get; set; }
        public DbSet<KeyReportUniqueKeyReport> KeyReportUniqueKeyReport { get; set; }
        public DbSet<KeyReportNotes> KeyReportNotes { get; set; }
        public DbSet<KeyReportNumber> KeyReportNumber { get; set; }
        public DbSet<KeyReportTester> KeyReportTester { get; set; }
        public DbSet<KeyReportReviewer> KeyReportReviewer { get; set; }
        public DbSet<ParametersLibrary> ParametersLibrary { get; set; }
        public DbSet<ReportsLibrary> ReportsLibrary { get; set; }
        public DbSet<CompletenessLibrary> CompletenessLibrary { get; set; }
        public DbSet<AccuracyLibrary> AccuracyLibrary { get; set; }
        public DbSet<CAMethodLibrary> CAMethodLibrary { get; set; }
        public DbSet<QuestionaireAddedInputs> QuestionaireAddedInputs { get; set; }
        public DbSet<KeyReportScreenshotUpload> KeyReportScreenshotUpload { get; set; }
        public DbSet<KeyReportFile> KeyReportFile { get; set; }

    #endregion


}
}
