// <auto-generated />
using System;
using A2B_App.Server.Data;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Infrastructure;
using Microsoft.EntityFrameworkCore.Storage.ValueConversion;

namespace A2B_App.Server.Migrations
{
    [DbContext(typeof(UserContext))]
    partial class UserContextModelSnapshot : ModelSnapshot
    {
        protected override void BuildModel(ModelBuilder modelBuilder)
        {
#pragma warning disable 612, 618
            modelBuilder
                .HasAnnotation("ProductVersion", "3.1.6")
                .HasAnnotation("Relational:MaxIdentifierLength", 64);

            modelBuilder.Entity("A2B_App.Shared.Podio.PodioRef", b =>
                {
                    b.Property<int>("Id")
                        .ValueGeneratedOnAdd()
                        .HasColumnType("int");

                    b.Property<string>("CreatedBy")
                        .HasColumnType("text");

                    b.Property<DateTimeOffset?>("CreatedOn")
                        .HasColumnType("timestamp");

                    b.Property<int>("ItemId")
                        .HasColumnType("int");

                    b.Property<DateTimeOffset>("LastUpdate")
                        .ValueGeneratedOnAddOrUpdate()
                        .HasColumnType("timestamp");

                    b.Property<string>("Link")
                        .HasColumnType("text");

                    b.Property<int>("Revision")
                        .HasColumnType("int");

                    b.Property<string>("UniqueId")
                        .HasColumnType("varchar(100)")
                        .HasMaxLength(100);

                    b.HasKey("Id");

                    b.HasIndex("Id", "ItemId", "UniqueId");

                    b.ToTable("PodioRef");
                });

            modelBuilder.Entity("A2B_App.Shared.Skype.Conversation", b =>
                {
                    b.Property<int>("Sys_id")
                        .ValueGeneratedOnAdd()
                        .HasColumnType("int");

                    b.Property<string>("id")
                        .HasColumnType("text");

                    b.Property<bool>("isgroup")
                        .HasColumnType("bit");

                    b.HasKey("Sys_id");

                    b.ToTable("Conversation");
                });

            modelBuilder.Entity("A2B_App.Shared.Skype.SkypeObj", b =>
                {
                    b.Property<int>("Sys_id")
                        .ValueGeneratedOnAdd()
                        .HasColumnType("int");

                    b.Property<string>("channelId")
                        .HasColumnType("text");

                    b.Property<int?>("conversationSys_id")
                        .HasColumnType("int");

                    b.Property<string>("id")
                        .HasColumnType("text");

                    b.Property<string>("serviceUrl")
                        .HasColumnType("text");

                    b.HasKey("Sys_id");

                    b.HasIndex("conversationSys_id");

                    b.ToTable("SkypeObj");
                });

            modelBuilder.Entity("A2B_App.Shared.Time.ProfileImage", b =>
                {
                    b.Property<int>("Id")
                        .ValueGeneratedOnAdd()
                        .HasColumnType("int");

                    b.Property<string>("Filename")
                        .HasColumnType("text");

                    b.Property<string>("Link")
                        .HasColumnType("text");

                    b.HasKey("Id");

                    b.ToTable("ProfileImage");
                });

            modelBuilder.Entity("A2B_App.Shared.Time.Skills", b =>
                {
                    b.Property<int>("Id")
                        .ValueGeneratedOnAdd()
                        .HasColumnType("int");

                    b.Property<DateTimeOffset?>("CreatedOn")
                        .HasColumnType("timestamp");

                    b.Property<DateTimeOffset>("LastUpdate")
                        .ValueGeneratedOnAddOrUpdate()
                        .HasColumnType("timestamp");

                    b.Property<string>("Skill")
                        .HasColumnType("varchar(100)")
                        .HasMaxLength(100);

                    b.Property<int>("TeamMemberDetailId")
                        .HasColumnType("int");

                    b.HasKey("Id");

                    b.HasIndex("TeamMemberDetailId");

                    b.HasIndex("Id", "Skill");

                    b.ToTable("Skills");
                });

            modelBuilder.Entity("A2B_App.Shared.Time.TeamMember", b =>
                {
                    b.Property<int>("Id")
                        .ValueGeneratedOnAdd()
                        .HasColumnType("int");

                    b.Property<DateTimeOffset?>("CreatedOn")
                        .HasColumnType("timestamp");

                    b.Property<DateTimeOffset>("LastUpdate")
                        .ValueGeneratedOnAddOrUpdate()
                        .HasColumnType("timestamp");

                    b.Property<string>("Name")
                        .HasColumnType("text");

                    b.Property<string>("Organization")
                        .HasColumnType("varchar(100)")
                        .HasMaxLength(100);

                    b.Property<int?>("PodioRefId")
                        .HasColumnType("int");

                    b.Property<int>("ProfileId")
                        .HasColumnType("int");

                    b.Property<int?>("ProfileImageId")
                        .HasColumnType("int");

                    b.Property<string>("SkypeName")
                        .HasColumnType("varchar(100)")
                        .HasMaxLength(100);

                    b.Property<string>("SkypeObjRaw")
                        .HasColumnType("text");

                    b.Property<int?>("SkypeObjSys_id")
                        .HasColumnType("int");

                    b.Property<string>("Status")
                        .HasColumnType("varchar(100)")
                        .HasMaxLength(100);

                    b.Property<int?>("TeamMemberDetailId")
                        .HasColumnType("int");

                    b.Property<int>("UserId")
                        .HasColumnType("int");

                    b.HasKey("Id");

                    b.HasIndex("PodioRefId");

                    b.HasIndex("ProfileImageId");

                    b.HasIndex("SkypeObjSys_id");

                    b.HasIndex("TeamMemberDetailId");

                    b.HasIndex("Id", "Organization", "Status", "UserId", "ProfileId", "SkypeName");

                    b.ToTable("TeamMember");
                });

            modelBuilder.Entity("A2B_App.Shared.Time.TeamMemberDetail", b =>
                {
                    b.Property<int>("Id")
                        .ValueGeneratedOnAdd()
                        .HasColumnType("int");

                    b.Property<string>("About")
                        .HasColumnType("text");

                    b.Property<string>("Address")
                        .HasColumnType("text");

                    b.Property<DateTime?>("BirthDate")
                        .HasColumnType("datetime");

                    b.Property<string>("City")
                        .HasColumnType("text");

                    b.Property<string>("Country")
                        .HasColumnType("text");

                    b.Property<DateTime?>("HiredDate")
                        .HasColumnType("datetime");

                    b.Property<string>("JobTitle")
                        .HasColumnType("text");

                    b.Property<string>("LinkedInURL")
                        .HasColumnType("text");

                    b.Property<string>("Mobile")
                        .HasColumnType("text");

                    b.Property<string>("State")
                        .HasColumnType("text");

                    b.Property<string>("Url")
                        .HasColumnType("text");

                    b.Property<string>("Zip")
                        .HasColumnType("text");

                    b.HasKey("Id");

                    b.ToTable("TeamMemberDetail");
                });

            modelBuilder.Entity("A2B_App.Shared.Time.UserEmail", b =>
                {
                    b.Property<int>("Id")
                        .ValueGeneratedOnAdd()
                        .HasColumnType("int");

                    b.Property<DateTimeOffset?>("CreatedOn")
                        .HasColumnType("timestamp");

                    b.Property<string>("Email")
                        .HasColumnType("text");

                    b.Property<DateTimeOffset>("LastUpdate")
                        .ValueGeneratedOnAddOrUpdate()
                        .HasColumnType("timestamp");

                    b.Property<int?>("TeamMemberId")
                        .HasColumnType("int");

                    b.HasKey("Id");

                    b.HasIndex("TeamMemberId");

                    b.ToTable("UserEmail");
                });

            modelBuilder.Entity("A2B_App.Shared.Skype.SkypeObj", b =>
                {
                    b.HasOne("A2B_App.Shared.Skype.Conversation", "conversation")
                        .WithMany()
                        .HasForeignKey("conversationSys_id");
                });

            modelBuilder.Entity("A2B_App.Shared.Time.Skills", b =>
                {
                    b.HasOne("A2B_App.Shared.Time.TeamMemberDetail", null)
                        .WithMany("ListSkill")
                        .HasForeignKey("TeamMemberDetailId")
                        .OnDelete(DeleteBehavior.Cascade)
                        .IsRequired();
                });

            modelBuilder.Entity("A2B_App.Shared.Time.TeamMember", b =>
                {
                    b.HasOne("A2B_App.Shared.Podio.PodioRef", "PodioRef")
                        .WithMany()
                        .HasForeignKey("PodioRefId");

                    b.HasOne("A2B_App.Shared.Time.ProfileImage", "ProfileImage")
                        .WithMany()
                        .HasForeignKey("ProfileImageId");

                    b.HasOne("A2B_App.Shared.Skype.SkypeObj", "SkypeObj")
                        .WithMany()
                        .HasForeignKey("SkypeObjSys_id");

                    b.HasOne("A2B_App.Shared.Time.TeamMemberDetail", "TeamMemberDetail")
                        .WithMany()
                        .HasForeignKey("TeamMemberDetailId");
                });

            modelBuilder.Entity("A2B_App.Shared.Time.UserEmail", b =>
                {
                    b.HasOne("A2B_App.Shared.Time.TeamMember", null)
                        .WithMany("ListEmail")
                        .HasForeignKey("TeamMemberId");
                });
#pragma warning restore 612, 618
        }
    }
}
