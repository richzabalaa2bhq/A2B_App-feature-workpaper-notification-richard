using A2B_App.Shared.Podio;
using A2B_App.Shared.Time;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Design;
using Microsoft.Extensions.Configuration;
using System;
using System.IO;

namespace A2B_App.Server.Data
{
    public class TimeContext : DbContext
    {
        public class TimeContextFactory : IDesignTimeDbContextFactory<TimeContext>
        {
            private IConfiguration _configuration;

            public TimeContextFactory()
            {
                var builder = new ConfigurationBuilder()
                        .SetBasePath(Directory.GetCurrentDirectory())
                        .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

                _configuration = builder.Build();
            }

            public TimeContext CreateDbContext(string[] args)
            {
                var optionsBuilder = new DbContextOptionsBuilder<TimeContext>();
                optionsBuilder.UseMySQL(_configuration.GetConnectionString("TimeCon"), opts => opts.CommandTimeout((int)TimeSpan.FromMinutes(30).TotalSeconds));
               
                return new TimeContext(optionsBuilder.Options);
            }

        }

        public TimeContext(DbContextOptions<TimeContext> options)
        : base(options)
        {

        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {

            modelBuilder.Entity<TimeCode>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.ClientRef, x.ClientCode, x.ProjectRef, x.TaskRef, x.Status });
                e.Property(x => x.ClientRef).HasMaxLength(250);
                e.Property(x => x.ClientCode).HasMaxLength(250);
                e.Property(x => x.ProjectRef).HasMaxLength(250);
                e.Property(x => x.TaskRef).HasMaxLength(250);
                e.Property(x => x.Status).HasMaxLength(250);
            });

            modelBuilder.Entity<ClientReference>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.ClientRef, x.ClientCode});
                e.Property(x => x.ClientRef).HasMaxLength(250);
                e.Property(x => x.ClientCode).HasMaxLength(250);
            });

            modelBuilder.Entity<ProjectReference>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.ProjectRef});
                e.Property(x => x.ProjectRef).HasMaxLength(250);
            });

            modelBuilder.Entity<TaskReference>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.TaskRef });
                e.Property(x => x.TaskRef).HasMaxLength(250);
            });

            modelBuilder.Entity<MasterTime>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.BillStat, x.ClientCode, x.ClientName, x.Project, x.Employee });
                e.Property(x => x.BillStat).HasMaxLength(250);
                e.Property(x => x.ClientCode).HasMaxLength(250);
                e.Property(x => x.ClientName).HasMaxLength(250);
                e.Property(x => x.Project).HasMaxLength(250);
                e.Property(x => x.Employee).HasMaxLength(250);

            });

            modelBuilder.Entity<MasterTimeDetail>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id });
            });

            modelBuilder.Entity<PodioRef>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.ItemId, x.UniqueId, x.CreatedOn });
                e.Property(x => x.UniqueId).HasMaxLength(250);
            });

            modelBuilder.Entity<PodioHook>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id });
            });


        }
        public DbSet<TimeCode> TimeCode { get; set; }
        public DbSet<ClientReference> ClientReference { get; set; }
        public DbSet<ProjectReference> ProjectReference { get; set; }
        public DbSet<TaskReference> TaskReference { get; set; }
        public DbSet<PodioHook> PodioHook { get; set; }
        public DbSet<MasterTime> MasterTime { get; set; }
        public DbSet<MasterTimeDetail> MasterTimeDetail { get; set; }
        public DbSet<PodioRef> PodioRef { get; set; }
        public DbSet<A2BPodioUser> A2BPodioUser { get; set; }
    }
}
