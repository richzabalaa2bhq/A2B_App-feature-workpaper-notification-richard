using A2B_App.Shared.Skype;
using A2B_App.Shared.Sms;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Design;
using Microsoft.Extensions.Configuration;
using System.IO;

namespace A2B_App.Server.Data
{
    public class NotificationContext : DbContext
    {

        public class NotificationContextFactory : IDesignTimeDbContextFactory<NotificationContext>
        {
            private IConfiguration _configuration;

            public NotificationContextFactory()
            {
                var builder = new ConfigurationBuilder()
                        .SetBasePath(Directory.GetCurrentDirectory())
                        .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

                _configuration = builder.Build();
            }

            public NotificationContext CreateDbContext(string[] args)
            {
                var optionsBuilder = new DbContextOptionsBuilder<NotificationContext>();
                optionsBuilder.UseMySQL(_configuration.GetConnectionString("NotificationCon"));

                return new NotificationContext(optionsBuilder.Options);
            }

        }

        public NotificationContext(DbContextOptions<NotificationContext> options)
        : base(options)
        {
            Database.SetCommandTimeout(600);
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {

            modelBuilder.Entity<Notification>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasMany(x => x.ListCategory);
                e.Property(x => x.Title).HasMaxLength(250);
                e.HasIndex(x => new { x.Id });
            });

            modelBuilder.Entity<Category>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasMany(x => x.ListSkypeObj);
                e.HasMany(x => x.ListKeyWord);
                e.HasIndex(x => new { x.Id });
            });

            modelBuilder.Entity<SkypeObj>(e =>
            {
                e.Property(x => x.Sys_id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Sys_id });
            });

            modelBuilder.Entity<KeyWordGC>(e =>
            {
                e.Property(x => x.Sys_id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Sys_id, x.Keyword });
                e.Property(x => x.Keyword).HasMaxLength(250);
            });

            modelBuilder.Entity<Conversation>(e =>
            {
                e.Property(x => x.Sys_id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Sys_id });
            });

        }
        public DbSet<Category> Category { get; set; }
        public DbSet<Notification> Notification { get; set; }
        public DbSet<SkypeObj> SkypeObj { get; set; }
        public DbSet<KeyWordGC> KeyWordGC { get; set; }
        public DbSet<Conversation> Conversation { get; set; }
    }
}
