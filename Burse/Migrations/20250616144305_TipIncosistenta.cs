using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace Burse.Migrations
{
    public partial class TipIncosistenta : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<string>(
                name: "TipInconsistenta",
                table: "StudentRecord",
                type: "nvarchar(max)",
                nullable: true);
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "TipInconsistenta",
                table: "StudentRecord");
        }
    }
}
