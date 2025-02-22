using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace Burse.Migrations
{
    public partial class AdaugareFondBurseMeritRepartizat : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "FondBurseMeritRepartizat",
                columns: table => new
                {
                    ID = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    domeniu = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    bursaAlocatata = table.Column<decimal>(type: "decimal(18,2)", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_FondBurseMeritRepartizat", x => x.ID);
                });
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "FondBurseMeritRepartizat");
        }
    }
}
