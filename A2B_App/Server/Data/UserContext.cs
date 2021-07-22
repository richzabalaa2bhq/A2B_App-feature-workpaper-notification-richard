using A2B_App.Shared.Podio;
using A2B_App.Shared.Skype;
using A2B_App.Shared.Sms;
using A2B_App.Shared.Time;
using A2B_App.Shared.User;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Design;
using Microsoft.Extensions.Configuration;
using System.IO;

namespace A2B_App.Server.Data
{
    public class UserContext : DbContext
    {

        public class UserContextFactory : IDesignTimeDbContextFactory<UserContext>
        {
            private IConfiguration _configuration;

            public UserContextFactory()
            {
                var builder = new ConfigurationBuilder()
                        .SetBasePath(Directory.GetCurrentDirectory())
                        .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

                _configuration = builder.Build();
            }

            public UserContext CreateDbContext(string[] args)
            {
                var optionsBuilder = new DbContextOptionsBuilder<UserContext>();
                optionsBuilder.UseMySQL(_configuration.GetConnectionString("UserCon"));

                return new UserContext(optionsBuilder.Options);
            }

        }

        public UserContext(DbContextOptions<UserContext> options)
        : base(options)
        {
            Database.SetCommandTimeout(600);
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {

            modelBuilder.Entity<TeamMember>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasIndex(x => new { x.Id, x.Organization, x.Status, x.UserId, x.ProfileId, x.SkypeName });
                e.Property(x => x.Organization).HasMaxLength(100);
                e.Property(x => x.Status).HasMaxLength(100);
                e.Property(x => x.SkypeName).HasMaxLength(100);
            });

            modelBuilder.Entity<SkypeObj>(e =>
            {
                e.Property(x => x.Sys_id).ValueGeneratedOnAdd();
            });

            modelBuilder.Entity<ProfileImage>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
            });

            modelBuilder.Entity<PodioRef>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.Property(x => x.UniqueId).HasMaxLength(100);
                e.HasIndex(x => new { x.Id, x.ItemId, x.UniqueId });
            });

            modelBuilder.Entity<TeamMemberDetail>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.HasMany(x => x.ListSkill);
            });

            modelBuilder.Entity<Skills>(e =>
            {
                e.Property(x => x.Id).ValueGeneratedOnAdd();
                e.Property(x => x.Skill).HasMaxLength(100);
                e.HasIndex(x => new { x.Id, x.Skill});
            });

        }
        public DbSet<TeamMember> TeamMember { get; set; }
        public DbSet<SkypeObj> SkypeObj { get; set; }
        public DbSet<ProfileImage> ProfileImage { get; set; }
        public DbSet<PodioRef> PodioRef { get; set; }
        public DbSet<TeamMemberDetail> TeamMemberDetail { get; set; }
        public DbSet<Skills> Skills { get; set; }
    }
}
