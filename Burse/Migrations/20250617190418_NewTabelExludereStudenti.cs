using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace Burse.Migrations
{
    public partial class NewTabelExludereStudenti : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "ExcludereStudent",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    CampExcludere = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Valoare = table.Column<string>(type: "nvarchar(max)", nullable: true)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_ExcludereStudent", x => x.Id);
                });
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "ExcludereStudent");
        }
    }
}
