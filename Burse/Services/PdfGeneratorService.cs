using Burse.Models.TemplatePDF;
using Burse.Services.Abstractions;

using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

using System.Text.RegularExpressions;

namespace Burse.Services
{
    public class PdfGeneratorService : IPdfGeneratorService
    {
        private readonly IFondBurseService _fondBurseService;
        private readonly IGrupuriService _grupuriService;
        public PdfGeneratorService(IFondBurseService fondBurseService, IGrupuriService grupuriService)
        {
            _fondBurseService = fondBurseService;
            _grupuriService = grupuriService;
        }

        public async Task<MemoryStream> GeneratePdfAsync(PdfRequest request)
        {
            if (request.Elements == null)
            {
                throw new ArgumentException("Request-ul este invalid sau nu conține niciun ID de template.");
            }

            QuestPDF.Settings.License = LicenseType.Community;

            var studentiCuBursa0 = await _fondBurseService.GetStudentsWithBursaFromDatabaseAsync();
            var acronymMappings = await _grupuriService.GetGrupuriAcronimeAsync();

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
                                var domeniiSelectate = new List<string>();

                                if (request.DynamicFields != null && request.DynamicFields.TryGetValue("ProgramStudiu.Dynamic", out var domeniiRaw))
                                {
                                    domeniiSelectate = domeniiRaw
                                        .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                                        .ToList();
                                }

                                var filtrati = studentiCuBursa0
                                    .Where(s => domeniiSelectate
                                        .Any(d => RemoveYearFromDomain(s.FondBurseMeritRepartizat.domeniu).Equals(RemoveYearFromDomain(d))))
                                    .OrderBy(s => s.An).ThenByDescending(s => s.Media)
                                    .ToList();

                                col.Item().Table(table =>
                                {
                                    table.ColumnsDefinition(columns =>
                                    {
                                        columns.ConstantColumn(20);
                                        columns.ConstantColumn(60);
                                        columns.ConstantColumn(30);
                                        columns.ConstantColumn(35);
                                        columns.ConstantColumn(80);
                                        columns.ConstantColumn(30);
                                        columns.ConstantColumn(50);
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
                                    var filtratiGrupati = filtrati
                                        .OrderBy(s => s.An + 1)
                                        .ThenBy(s => s.FondBurseMeritRepartizat.domeniu.Contains("DUAL") || s.FondBurseMeritRepartizat.domeniu.Contains("-DUAL") ? 1 : 0) 
                                        .ToList();

                                    foreach (var s in filtratiGrupati)
                                    {
                                        table.Cell().Element(CellStyle).Text(index++.ToString()).FontSize(9);
                                        table.Cell().Element(CellStyle).Text(s.Emplid).FontSize(9);
                                        string anText;
                                        if (s.FondBurseMeritRepartizat.domeniu.Contains("DUAL") || s.FondBurseMeritRepartizat.domeniu.Contains("-DUAL"))
                                        {
                                            anText = (s.An + 1).ToString() + "Dual";
                                        }
                                        else
                                        {
                                            anText = (s.An + 1).ToString();
                                        }
                                        table.Cell().Element(CellStyle).Text(anText).FontSize(9);
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
                                string content = ReplaceDynamicFields(el.Content, request.DynamicFields, acronymMappings);

                                text.Span(content)
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
            return stream;
        }

        private static IContainer CellStyle(IContainer container) =>
      container.Padding(0)      // Eliminăm padding-ul pentru a face tabelul mai compact
               .Border(0.5f)       // Eliminăm borderul (sau îl facem foarte subțire)
               .AlignCenter();

        // Functia pentru eliminarea anului și a sufixului "-DUAL"
        private string RemoveYearFromDomain(string domeniu)
        {
            return Regex.Replace(domeniu, @"\s*\(\d+\)|\-DUAL", "").Trim().ToLower();
        }
        private string ReplaceDynamicFields(string content, Dictionary<string, string> dynamicFields, Dictionary<string, List<string>> acronymMappings)
        {
            if (string.IsNullOrWhiteSpace(content) || dynamicFields == null)
                return content;

            foreach (var field in dynamicFields)
            {
                if (field.Key == "ProgramStudiu.Dynamic")
                {
                    var acronimeSelectate = field.Value
                        .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                        .ToList();

                    var domeniiSelectate = new List<string>();

                    foreach (var acronim in acronimeSelectate)
                    {
                        if (acronim.Contains("DUAL"))
                        {
                            // Pentru acronime cu DUAL, căutăm versiunea fără DUAL în dicționar
                            string acronimFaraDual = acronim.Replace("-DUAL", "");
                            var programeGasite = acronymMappings
                                .Where(kvp => kvp.Value.Contains(acronimFaraDual))
                                .Select(kvp => kvp.Key + "-DUAL") // Adăugăm -DUAL la cheia găsită
                                .ToList();
                            domeniiSelectate.AddRange(programeGasite);
                        }
                        else
                        {
                            // Pentru acronime normale, căutăm direct
                            var programeGasite = acronymMappings
                                .Where(kvp => kvp.Value.Contains(acronim))
                                .Select(kvp => kvp.Key)
                                .ToList();
                            domeniiSelectate.AddRange(programeGasite);
                        }
                    }

                    domeniiSelectate = domeniiSelectate.Distinct().ToList();
                    string domeniiText = string.Join("/ ", domeniiSelectate);

                    content = content.Replace(field.Key, domeniiText);
                }
                else
                {
                    content = content.Replace(field.Key, field.Value);
                }
            }

            return content;

        }
    }

}
