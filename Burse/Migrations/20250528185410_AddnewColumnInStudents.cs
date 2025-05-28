using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace Burse.Migrations
{
    public partial class AddnewColumnInStudents : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.AddColumn<decimal>(
                name: "MEDG_ASL",
                table: "StudentRecord",
                type: "decimal(18,2)",
                nullable: false,
                defaultValue: 0m);

            migrationBuilder.AddColumn<decimal>(
                name: "MediaBac",
                table: "StudentRecord",
                type: "decimal(18,2)",
                nullable: false,
                defaultValue: 0m);

            migrationBuilder.AddColumn<decimal>(
                name: "MediaBacMat",
                table: "StudentRecord",
                type: "decimal(18,2)",
                nullable: false,
                defaultValue: 0m);

            migrationBuilder.AddColumn<decimal>(
                name: "MediaDL",
                table: "StudentRecord",
                type: "decimal(18,2)",
                nullable: false,
                defaultValue: 0m);

            migrationBuilder.AddColumn<decimal>(
                name: "MediaInterviu",
                table: "StudentRecord",
                type: "decimal(18,2)",
                nullable: false,
                defaultValue: 0m);
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropColumn(
                name: "MEDG_ASL",
                table: "StudentRecord");

            migrationBuilder.DropColumn(
                name: "MediaBac",
                table: "StudentRecord");

            migrationBuilder.DropColumn(
                name: "MediaBacMat",
                table: "StudentRecord");

            migrationBuilder.DropColumn(
                name: "MediaDL",
                table: "StudentRecord");

            migrationBuilder.DropColumn(
                name: "MediaInterviu",
                table: "StudentRecord");
        }
    }
}
