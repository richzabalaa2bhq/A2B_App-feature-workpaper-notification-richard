using System;
using Microsoft.EntityFrameworkCore.Migrations;
using MySql.Data.EntityFrameworkCore.Metadata;

namespace A2B_App.Server.Migrations
{
    public partial class userv2 : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "Conversation",
                columns: table => new
                {
                    Sys_id = table.Column<int>(nullable: false)
                        .Annotation("MySQL:ValueGenerationStrategy", MySQLValueGenerationStrategy.IdentityColumn),
                    id = table.Column<string>(nullable: true),
                    isgroup = table.Column<bool>(nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Conversation", x => x.Sys_id);
                });

            migrationBuilder.CreateTable(
                name: "PodioRef",
                columns: table => new
                {
                    Id = table.Column<int>(nullable: false)
                        .Annotation("MySQL:ValueGenerationStrategy", MySQLValueGenerationStrategy.IdentityColumn),
                    ItemId = table.Column<int>(nullable: false),
                    UniqueId = table.Column<string>(maxLength: 100, nullable: true),
                    Revision = table.Column<int>(nullable: false),
                    Link = table.Column<string>(nullable: true),
                    CreatedBy = table.Column<string>(nullable: true),
                    CreatedOn = table.Column<DateTimeOffset>(nullable: true),
                    LastUpdate = table.Column<DateTimeOffset>(nullable: false)
                        .Annotation("MySQL:ValueGenerationStrategy", MySQLValueGenerationStrategy.ComputedColumn)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_PodioRef", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "ProfileImage",
                columns: table => new
                {
                    Id = table.Column<int>(nullable: false)
                        .Annotation("MySQL:ValueGenerationStrategy", MySQLValueGenerationStrategy.IdentityColumn),
                    Link = table.Column<string>(nullable: true),
                    Filename = table.Column<string>(nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_ProfileImage", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "TeamMemberDetail",
                columns: table => new
                {
                    Id = table.Column<int>(nullable: false)
                        .Annotation("MySQL:ValueGenerationStrategy", MySQLValueGenerationStrategy.IdentityColumn),
                    BirthDate = table.Column<DateTime>(nullable: true),
                    HiredDate = table.Column<DateTime>(nullable: true),
                    LinkedInURL = table.Column<string>(nullable: true),
                    About = table.Column<string>(nullable: true),
                    Address = table.Column<string>(nullable: true),
                    Zip = table.Column<string>(nullable: true),
                    City = table.Column<string>(nullable: true),
                    Country = table.Column<string>(nullable: true),
                    State = table.Column<string>(nullable: true),
                    Mobile = table.Column<string>(nullable: true),
                    JobTitle = table.Column<string>(nullable: true),
                    Url = table.Column<string>(nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_TeamMemberDetail", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "SkypeObj",
                columns: table => new
                {
                    Sys_id = table.Column<int>(nullable: false)
                        .Annotation("MySQL:ValueGenerationStrategy", MySQLValueGenerationStrategy.IdentityColumn),
                    id = table.Column<string>(nullable: true),
                    channelId = table.Column<string>(nullable: true),
                    serviceUrl = table.Column<string>(nullable: true),
                    conversationSys_id = table.Column<int>(nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_SkypeObj", x => x.Sys_id);
                    table.ForeignKey(
                        name: "FK_SkypeObj_Conversation_conversationSys_id",
                        column: x => x.conversationSys_id,
                        principalTable: "Conversation",
                        principalColumn: "Sys_id",
                        onDelete: ReferentialAction.Restrict);
                });

            migrationBuilder.CreateTable(
                name: "Skills",
                columns: table => new
                {
                    Id = table.Column<int>(nullable: false)
                        .Annotation("MySQL:ValueGenerationStrategy", MySQLValueGenerationStrategy.IdentityColumn),
                    Skill = table.Column<string>(maxLength: 100, nullable: true),
                    CreatedOn = table.Column<DateTimeOffset>(nullable: true),
                    LastUpdate = table.Column<DateTimeOffset>(nullable: false)
                        .Annotation("MySQL:ValueGenerationStrategy", MySQLValueGenerationStrategy.ComputedColumn),
                    TeamMemberDetailId = table.Column<int>(nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Skills", x => x.Id);
                    table.ForeignKey(
                        name: "FK_Skills_TeamMemberDetail_TeamMemberDetailId",
                        column: x => x.TeamMemberDetailId,
                        principalTable: "TeamMemberDetail",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateTable(
                name: "TeamMember",
                columns: table => new
                {
                    Id = table.Column<int>(nullable: false)
                        .Annotation("MySQL:ValueGenerationStrategy", MySQLValueGenerationStrategy.IdentityColumn),
                    Name = table.Column<string>(nullable: true),
                    Organization = table.Column<string>(maxLength: 100, nullable: true),
                    Status = table.Column<string>(maxLength: 100, nullable: true),
                    UserId = table.Column<int>(nullable: false),
                    ProfileId = table.Column<int>(nullable: false),
                    SkypeName = table.Column<string>(maxLength: 100, nullable: true),
                    SkypeObjRaw = table.Column<string>(nullable: true),
                    SkypeObjSys_id = table.Column<int>(nullable: true),
                    ProfileImageId = table.Column<int>(nullable: true),
                    PodioRefId = table.Column<int>(nullable: true),
                    CreatedOn = table.Column<DateTimeOffset>(nullable: true),
                    TeamMemberDetailId = table.Column<int>(nullable: true),
                    LastUpdate = table.Column<DateTimeOffset>(nullable: false)
                        .Annotation("MySQL:ValueGenerationStrategy", MySQLValueGenerationStrategy.ComputedColumn)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_TeamMember", x => x.Id);
                    table.ForeignKey(
                        name: "FK_TeamMember_PodioRef_PodioRefId",
                        column: x => x.PodioRefId,
                        principalTable: "PodioRef",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Restrict);
                    table.ForeignKey(
                        name: "FK_TeamMember_ProfileImage_ProfileImageId",
                        column: x => x.ProfileImageId,
                        principalTable: "ProfileImage",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Restrict);
                    table.ForeignKey(
                        name: "FK_TeamMember_SkypeObj_SkypeObjSys_id",
                        column: x => x.SkypeObjSys_id,
                        principalTable: "SkypeObj",
                        principalColumn: "Sys_id",
                        onDelete: ReferentialAction.Restrict);
                    table.ForeignKey(
                        name: "FK_TeamMember_TeamMemberDetail_TeamMemberDetailId",
                        column: x => x.TeamMemberDetailId,
                        principalTable: "TeamMemberDetail",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Restrict);
                });

            migrationBuilder.CreateTable(
                name: "UserEmail",
                columns: table => new
                {
                    Id = table.Column<int>(nullable: false)
                        .Annotation("MySQL:ValueGenerationStrategy", MySQLValueGenerationStrategy.IdentityColumn),
                    Email = table.Column<string>(nullable: true),
                    CreatedOn = table.Column<DateTimeOffset>(nullable: true),
                    LastUpdate = table.Column<DateTimeOffset>(nullable: false)
                        .Annotation("MySQL:ValueGenerationStrategy", MySQLValueGenerationStrategy.ComputedColumn),
                    TeamMemberId = table.Column<int>(nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_UserEmail", x => x.Id);
                    table.ForeignKey(
                        name: "FK_UserEmail_TeamMember_TeamMemberId",
                        column: x => x.TeamMemberId,
                        principalTable: "TeamMember",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Restrict);
                });

            migrationBuilder.CreateIndex(
                name: "IX_PodioRef_Id_ItemId_UniqueId",
                table: "PodioRef",
                columns: new[] { "Id", "ItemId", "UniqueId" });

            migrationBuilder.CreateIndex(
                name: "IX_Skills_TeamMemberDetailId",
                table: "Skills",
                column: "TeamMemberDetailId");

            migrationBuilder.CreateIndex(
                name: "IX_Skills_Id_Skill",
                table: "Skills",
                columns: new[] { "Id", "Skill" });

            migrationBuilder.CreateIndex(
                name: "IX_SkypeObj_conversationSys_id",
                table: "SkypeObj",
                column: "conversationSys_id");

            migrationBuilder.CreateIndex(
                name: "IX_TeamMember_PodioRefId",
                table: "TeamMember",
                column: "PodioRefId");

            migrationBuilder.CreateIndex(
                name: "IX_TeamMember_ProfileImageId",
                table: "TeamMember",
                column: "ProfileImageId");

            migrationBuilder.CreateIndex(
                name: "IX_TeamMember_SkypeObjSys_id",
                table: "TeamMember",
                column: "SkypeObjSys_id");

            migrationBuilder.CreateIndex(
                name: "IX_TeamMember_TeamMemberDetailId",
                table: "TeamMember",
                column: "TeamMemberDetailId");

            migrationBuilder.CreateIndex(
                name: "IX_TeamMember_Id_Organization_Status_UserId_ProfileId_SkypeName",
                table: "TeamMember",
                columns: new[] { "Id", "Organization", "Status", "UserId", "ProfileId", "SkypeName" });

            migrationBuilder.CreateIndex(
                name: "IX_UserEmail_TeamMemberId",
                table: "UserEmail",
                column: "TeamMemberId");
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "Skills");

            migrationBuilder.DropTable(
                name: "UserEmail");

            migrationBuilder.DropTable(
                name: "TeamMember");

            migrationBuilder.DropTable(
                name: "PodioRef");

            migrationBuilder.DropTable(
                name: "ProfileImage");

            migrationBuilder.DropTable(
                name: "SkypeObj");

            migrationBuilder.DropTable(
                name: "TeamMemberDetail");

            migrationBuilder.DropTable(
                name: "Conversation");
        }
    }
}
