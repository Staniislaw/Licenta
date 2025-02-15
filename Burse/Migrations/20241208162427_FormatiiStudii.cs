using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace Burse.Migrations
{
    public partial class FormatiiStudii : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "Student");

            migrationBuilder.CreateTable(
                name: "FormatiiStudii",
                columns: table => new
                {
                    id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    Facultatea = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    ProgramDeStudiu = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    An = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    FaraTaxaRomani = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    FaraTaxaRp = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    FaraTaxaUECEE = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    CuTaxaRomani = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    ElibiliB = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    CuTaxaRM = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    RMEligibil = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    CuTaxaUECEE = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    BursieriAIStatuluiRoman = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    CPV = table.Column<string>(type: "nvarchar(max)", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_FormatiiStudii", x => x.id);
                });
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "FormatiiStudii");

            migrationBuilder.CreateTable(
                name: "Student",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    Name = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    StudentId = table.Column<int>(type: "int", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Student", x => x.Id);
                });
        }
    }
}
