using Microsoft.AspNetCore.Mvc;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using QuestPDF.Drawing;
using System.Text.Json;
using Burse.Models.TemplatePDF;
using Burse.Services.Abstractions;
using System.Text.RegularExpressions;
using System;
using Burse.Data;
using Microsoft.EntityFrameworkCore;
namespace Burse.Controllers
{
    [ApiController]
    [Route("api/pdf")]
    public class PdfController : ControllerBase
    {
        private readonly IFondBurseService _fondBurseService;
        private readonly BurseDBContext _context;

        public PdfController(IFondBurseService fondBurseService, BurseDBContext context)
        {
            _fondBurseService = fondBurseService;
            _context = context;

        }
        private static IContainer CellStyle(IContainer container) =>
    container.Padding(0)      // Eliminăm padding-ul pentru a face tabelul mai compact
             .Border(0.5f)       // Eliminăm borderul (sau îl facem foarte subțire)
             .AlignCenter();

        [HttpPost("generate")]
        public async Task<IActionResult> GeneratePdf([FromBody] PdfRequest request)
        {
            QuestPDF.Settings.License = LicenseType.Community;

            var studentiCuBursa0 = await _fondBurseService.GetStudentsWithBursaFromDatabaseAsync();

            var document = Document.Create(container =>
            {
                container.Page(page =>
                {
                    page.Size(PageSizes.A4);
                    page.Margin(30);
                    page.DefaultTextStyle(x => x.FontSize(9));

                    page.Content().Column(col =>
                    {
                        foreach (var el in request.Elements)
                        {
                            var style = el.Style ?? new PdfStyle();

                            if (el.Type == "table")
                            {
                                var domeniiSelectate = el.Domenii ?? new List<string>();

                                var filtrati = studentiCuBursa0
                                    .Where(s => domeniiSelectate
                                        .Any(d => RemoveYearFromDomain(s.FondBurseMeritRepartizat.domeniu).Equals(RemoveYearFromDomain(d))))
                                    .OrderBy(s => s.An).ThenByDescending(S=>S.Media)
                                    .ToList();
                                col.Item().Table(table =>
                                {
                                    table.ColumnsDefinition(columns =>
                                    {
                                        columns.ConstantColumn(20);   // Nr. crt.
                                        columns.ConstantColumn(60);   // ID
                                        columns.ConstantColumn(30);   // An de studii
                                        columns.ConstantColumn(35);   // Media
                                        columns.ConstantColumn(55);  // Sursa de finanțare
                                        columns.ConstantColumn(30);  // Categorie bursă
                                        columns.ConstantColumn(50);  // Valoarea / 12 luni

                                    });

                                    table.Header(header =>
                                    {
                                        header.Cell().Element(CellStyle).Text("Nr. crt.").FontSize(9).Bold();
                                        header.Cell().Element(CellStyle).Text("ID").FontSize(9).Bold();
                                        header.Cell().Element(CellStyle).Text("An de studii").FontSize(9).Bold();
                                        header.Cell().Element(CellStyle).Text("Media").FontSize(9).Bold();
                                        header.Cell().Element(CellStyle).Text("Sursa de finanțare").FontSize(9).Bold();
                                        header.Cell().Element(CellStyle).RotateLeft().Text(text =>
                                        {
                                            text.Span("Categorie").FontSize(9).Bold();
                                            text.EmptyLine();
                                            text.Span("bursă").FontSize(9).Bold();
                                        });
                                        header.Cell().Element(CellStyle).Text("Valoarea / 12 luni").FontSize(9).Bold();
                                    });

                                    int index = 1;
                                    foreach (var s in filtrati)
                                    {
                                        table.Cell().Element(CellStyle).Text(index++.ToString()).FontSize(9);
                                        table.Cell().Element(CellStyle).Text(s.Emplid).FontSize(9);
                                        table.Cell().Element(CellStyle).Text((s.An+1).ToString()).FontSize(9);
                                        table.Cell().Element(CellStyle).Text(s.Media.ToString("0.00")).FontSize(9);
                                        table.Cell().Element(CellStyle).Text(s.SursaFinantare).FontSize(9);
                                        table.Cell().Element(CellStyle).Text(s.Bursa).FontSize(9);
                                        table.Cell().Element(CellStyle).Text((s.SumaBursa).ToString("0")).FontSize(9);
                                    }
                                });



                                col.Item().Height(20);
                                continue;
                            }

                            col.Item().Text(text =>
                            {
                                // Aplică stilul textului
                                text.Span(el.Content)
                                    .FontSize(style.FontSize)
                                    .FontColor(style.Color);

                                switch (style.TextAlign)
                                {
                                    case "center": text.AlignCenter(); break;
                                    case "right": text.AlignRight(); break;
                                    default: text.AlignLeft(); break;
                                }
                            });

                            col.Item().Height(10);
                        }
                    });
                });
            });

            var stream = new MemoryStream();
            document.GeneratePdf(stream);
            stream.Position = 0;

            return File(stream, "application/pdf", "generated.pdf");
        }

        // Functia pentru eliminarea anului și a sufixului "-DUAL"
        private string RemoveYearFromDomain(string domeniu)
        {
            return Regex.Replace(domeniu, @"\s*\(\d+\)|\-DUAL", "").Trim().ToLower();
        }
    }
}
