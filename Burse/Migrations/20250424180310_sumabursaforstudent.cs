using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace Burse.Migrations
{
    public partial class sumabursaforstudent : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "NrCrt",
                table: "StudentRecord");

            migrationBuilder.RenameColumn(
                name: "Bursa",
                table: "BursaIstoric",
                newName: "Bursa");

            migrationBuilder.AddColumn<decimal>(
                name: "SumaBursa",
                table: "StudentRecord",
                type: "decimal(18,2)",
                nullable: false,
                defaultValue: 0m);
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "SumaBursa",
                table: "StudentRecord");

            migrationBuilder.RenameColumn(
                name: "Bursa",
                table: "BursaIstoric",
                newName: "Bursa");

            migrationBuilder.AddColumn<int>(
                name: "NrCrt",
                table: "StudentRecord",
                type: "int",
                nullable: false,
                defaultValue: 0);
        }
    }
}
