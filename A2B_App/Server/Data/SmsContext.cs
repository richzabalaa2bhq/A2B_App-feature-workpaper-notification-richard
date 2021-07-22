using A2B_App.Shared.Sms;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Design;
using Microsoft.Extensions.Configuration;
using System.IO;

namespace A2B_App.Server.Data
{
    public class SmsContext : DbContext
    {

        public class SmsContextFactory : IDesignTimeDbContextFactory<SmsContext>
        {
            private IConfiguration _configuration;

            public SmsContextFactory()
            {
                var builder = new ConfigurationBuilder()
                        .SetBasePath(Directory.GetCurrentDirectory())
                        .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

                _configuration = builder.Build();
            }

            public SmsContext CreateDbContext(string[] args)
            {
                var optionsBuilder = new DbContextOptionsBuilder<SmsContext>();
                optionsBuilder.UseMySQL(_configuration.GetConnectionString("SmsCon"));

                return new SmsContext(optionsBuilder.Options);
            }

        }

        public SmsContext(DbContextOptions<SmsContext> options)
        : base(options)
        {

        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {

            //modelBuilder.Entity<GlobeSms>(e =>
            //{
            //    e.Property(x => x.Id).ValueGeneratedOnAdd();
            //});

            //modelBuilder.Entity<unsubscribed>(e =>
            //{
            //    e.Property(x => x.Id).ValueGeneratedOnAdd();
            //});

            modelBuilder.Entity<Subscribe>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
            });

            modelBuilder.Entity<EmployeeSmsReference>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
            });

        }
        public DbSet<Subscribe> Subscribe { get; set; }
        public DbSet<EmployeeSmsReference> EmployeeSmsReference { get; set; }

    }
}
