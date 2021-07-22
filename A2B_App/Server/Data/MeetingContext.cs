
using A2B_App.Shared.Meeting;
using A2B_App.Shared.Skype;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Design;
using Microsoft.Extensions.Configuration;
using System.IO;

namespace A2B_App.Server.Data
{
    public class MeetingContext : DbContext
    {

        public class MeetingContextFactory : IDesignTimeDbContextFactory<MeetingContext>
        {
            private IConfiguration _configuration;

            public MeetingContextFactory()
            {
                var builder = new ConfigurationBuilder()
                        .SetBasePath(Directory.GetCurrentDirectory())
                        .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

                _configuration = builder.Build();
            }

            public MeetingContext CreateDbContext(string[] args)
            {
                var optionsBuilder = new DbContextOptionsBuilder<MeetingContext>();
                optionsBuilder.UseMySQL(_configuration.GetConnectionString("MeetingCon"));

                return new MeetingContext(optionsBuilder.Options);
            }

        }

        public MeetingContext(DbContextOptions<MeetingContext> options)
        : base(options)
        {

        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {


            modelBuilder.Entity<meeting>(e =>
            {
                e.Property(x => x.sys_id).ValueGeneratedOnAdd();
            });

            modelBuilder.Entity<Conversation>(e =>
            {
                e.Property(x => x.Sys_id).ValueGeneratedOnAdd();
            });

            modelBuilder.Entity<SkypeObj>(e =>
            {
                e.Property(x => x.Sys_id).ValueGeneratedOnAdd();
            });

            modelBuilder.Entity<KeyWordGC>(e =>
            {
                e.Property(x => x.Sys_id).ValueGeneratedOnAdd();
            });

            modelBuilder.Entity<SkypeAddress>(e =>
            {
                e.Property(x => x.Sys_id).ValueGeneratedOnAdd();
            });

            modelBuilder.Entity<t_employee>().HasNoKey();


        }
        public DbSet<meeting> meeting { get; set; }
        public DbSet<dailymeeting> DailyMeeting { get; set; }
        public DbSet<dailymeetingBizDev> DailyMeetingBizDev { get; set; }
        
        public DbSet<dailymeetingAttendee> DailyMeetingParticipant { get; set; }

        public DbSet<recording_gtm_details> recording_gtm_details { get; set; }

        public DbSet<recording_gtm> recording_gtm { get; set; }

        public DbSet<Conversation> Conversation { get; set; }
        public DbSet<SkypeObj> SkypeObj { get; set; }
        public DbSet<KeyWordGC> KeyWordGC { get; set; }
        public DbSet<SkypeAddress> SkypeAddress { get; set; }
        public DbSet<t_employee> ClientParticipant { get; set; }


    }
}
