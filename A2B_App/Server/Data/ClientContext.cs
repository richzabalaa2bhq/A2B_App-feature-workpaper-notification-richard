using A2B_App.Shared.Admin;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Design;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace A2B_App.Server.Data
{
    public class ClientContext : DbContext
    {
        public class ClientContextFactory : IDesignTimeDbContextFactory<ClientContext>
        {
            private IConfiguration _configuration;

            public ClientContextFactory()
            {
                var builder = new ConfigurationBuilder()
                        .SetBasePath(Directory.GetCurrentDirectory())
                        .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

                _configuration = builder.Build();
            }

            public ClientContext CreateDbContext(string[] args)
            {
                var optionsBuilder = new DbContextOptionsBuilder<ClientContext>();
                optionsBuilder.UseMySQL(_configuration.GetConnectionString("SmsCon"));

                return new ClientContext(optionsBuilder.Options);
            }

        }

        public ClientContext(DbContextOptions<ClientContext> options)
        : base(options)
        {

        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<AppAccess>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
            });

        }
        public DbSet<AppAccess> AppAccess { get; set; }

    }
}
