using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace Burse.Migrations
{
    public partial class BurseIstoric : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "BursaIstoric",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    StudentRecordId = table.Column<int>(type: "int", nullable: false),
                    TipBursa = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Motiv = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Actiune = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Etapa = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    Suma = table.Column<decimal>(type: "decimal(18,2)", nullable: false),
                    Comentarii = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    DataModificare = table.Column<DateTime>(type: "datetime2", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_BursaIstoric", x => x.Id);
                    table.ForeignKey(
                        name: "FK_BursaIstoric_StudentRecord_StudentRecordId",
                        column: x => x.StudentRecordId,
                        principalTable: "StudentRecord",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateIndex(
                name: "IX_BursaIstoric_StudentRecordId",
                table: "BursaIstoric",
                column: "StudentRecordId");
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "BursaIstoric");
        }
    }
}
