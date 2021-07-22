
using A2B_App.Shared.Verification;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Design;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace A2B_App.Server.Data
{
    public class VerificationContext : DbContext
    {

        public class VerificationContextFactory : IDesignTimeDbContextFactory<VerificationContext>
        {
            private IConfiguration _configuration;

            public VerificationContextFactory()
            {
                var builder = new ConfigurationBuilder()
                        .SetBasePath(Directory.GetCurrentDirectory())
                        .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

                _configuration = builder.Build();
            }

            public VerificationContext CreateDbContext(string[] args)
            {
                var optionsBuilder = new DbContextOptionsBuilder<VerificationContext>();
                optionsBuilder.UseMySQL(_configuration.GetConnectionString("VerificationCon"));

                return new VerificationContext(optionsBuilder.Options);
            }

        }

        public VerificationContext(DbContextOptions<VerificationContext> options)
        : base(options)
        {

        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {

            modelBuilder.Entity<Verification>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
            });

            modelBuilder.Entity<RequestVerification>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
            });

            //modelBuilder.Entity<RequestVerification>().HasNoKey();

        }

        public DbSet<Verification> Verification { get; set; }
        public DbSet<RequestVerification> RequestVerification { get; set; }

    }
}
