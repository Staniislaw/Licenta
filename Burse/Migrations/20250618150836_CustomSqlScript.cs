using Burse.Helpers;

using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace Burse.Migrations
{
    public partial class CustomSqlScript : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.SqlResource("20250618150836_CustomSqlScript.sql");
        }
        protected override void Down(MigrationBuilder migrationBuilder)
        {
                    // Șterge înregistrările din GrupProgramStudii adăugate în Up
                    migrationBuilder.Sql(@"
                DELETE FROM [GrupProgramStudii]
                WHERE [Grup] IN ('IETTI', 'IEN', 'IE', 'IS', 'SIA', 'CTI', 'IA')
                    AND [Domeniu] IN ('SC', 'IETTI', 'RCC', 'RST', 'SMCPE', 'IEN', 'ME', 'ETI', 'SE', 'TAMAE', 'SE-DUAL', 'AIA', 'ESM', 'SIC', 'C', 'C-DUAL', 'AIA-DUAL', 'ESCCA');
            ");

                    // Șterge înregistrările din GrupPDF adăugate în Up
                    migrationBuilder.Sql(@"
                DELETE FROM [GrupPDF]
                WHERE [Grup] IN ('Grup: Calculatoare, Calculatoare-DUAL', 'Gup: Automatica')
                    AND [Valoare] IN ('C', 'AIA', 'AIA-DUAL', 'C-DUAL');
            ");

                    // Șterge înregistrările din GrupDomeniu adăugate în Up
                    migrationBuilder.Sql(@"
                DELETE FROM [GrupDomeniu]
                WHERE [Grup] IN ('IEN/ME/ETI', 'IETTI/RST')
                    AND [Domeniu] IN ('ME(3)', 'ETI(4)', 'IETTI(1)(2)', 'RST(3)(4)', 'IEN (1)(2)');
            ");

                    // Șterge înregistrările din GrupBursa adăugate în Up
                    migrationBuilder.Sql(@"
                DELETE FROM [GrupBursa]
                WHERE [GrupBursa] IN ('G1', 'G2', 'G3')
                    AND [Domeniu] IN ('AIA', 'IEN', 'ETI', 'ME', 'C', 'C-DUAL', 'IETTI', 'RST', 'SIC', 'SE', 'SE-DUAL', 'TAMAE', 'RCC', 'SC', 'EA', 'SMCPE', 'ESCCA', 'ESM', 'AIA-DUAL');
            ");

                    // Șterge template-ul inserat
                    migrationBuilder.Sql(@"
                DELETE FROM [TemplateEntity]
                WHERE [Name] = 'Template Bursa';
            ");

                    // Șterge înregistrările din GrupAcronim
                    migrationBuilder.Sql(@"
                DELETE FROM [GrupAcronim]
                WHERE [Valoare] IN (
                    'SE', 'ESM', 'AIA', 'TAMAE', 'ME', 'C', 'ESCCA',
                    'SMCPE', 'IETTI', 'ETI', 'RST', 'SC', 'IEN', 'RCC', 'SIC'
                );
            ");
        }
    }
}
