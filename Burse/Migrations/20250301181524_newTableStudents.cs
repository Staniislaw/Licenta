using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace Burse.Migrations
{
    public partial class newTableStudents : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "StudentRecord",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    NrCrt = table.Column<int>(type: "int", nullable: false),
                    Emplid = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    CNP = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    NumeStudent = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    TaraCetatenie = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    An = table.Column<int>(type: "int", nullable: false),
                    Media = table.Column<decimal>(type: "decimal(18,2)", nullable: false),
                    PunctajAn = table.Column<int>(type: "int", nullable: false),
                    CO = table.Column<int>(type: "int", nullable: false),
                    RO = table.Column<int>(type: "int", nullable: false),
                    TC = table.Column<int>(type: "int", nullable: false),
                    TR = table.Column<int>(type: "int", nullable: false),
                    SursaFinantare = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Bursa = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    FondBurseMeritRepartizatId = table.Column<int>(type: "int", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_StudentRecord", x => x.Id);
                    table.ForeignKey(
                        name: "FK_StudentRecord_FondBurseMeritRepartizat_FondBurseMeritRepartizatId",
                        column: x => x.FondBurseMeritRepartizatId,
                        principalTable: "FondBurseMeritRepartizat",
                        principalColumn: "ID",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateIndex(
                name: "IX_StudentRecord_FondBurseMeritRepartizatId",
                table: "StudentRecord",
                column: "FondBurseMeritRepartizatId");
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "StudentRecord");
        }
    }
}
