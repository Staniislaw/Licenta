using Burse.Data;
using Burse.Models;
using Microsoft.AspNetCore.Mvc;
using ExcelDataReader;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using Burse.Services.Abstractions;
using Burse.Helpers;
using System.Text.RegularExpressions;

namespace Burse.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class FondBurseController : ControllerBase
    {
        private readonly BurseDBContext _context;
        private readonly IFondBurseService _fondBurseService;
        public FondBurseController(BurseDBContext context, IFondBurseService fondBurseService)
        {
            _context = context;
            _fondBurseService = fondBurseService;
        }

        [HttpPost("AddFondBurse")]
        public async Task<IActionResult> AddFondBurse()
        {
            var filePath = "C:\\Users\\Stas\\Downloads\\Fond_burse_2024_2025 13noiembrie.xls"; 

            var excelReader = new FondBurseExcelReader();
            var fonduriBurse = excelReader.ReadFondBurseFromExcel(filePath);

            var fonduriBurseNoi = fonduriBurse.Where(f => !_context.FondBurse.Any(fb => fb.CategorieBurse == f.CategorieBurse)).ToList();

            var filePath2 = "C:\\Users\\Stas\\Downloads\\Formatii studii USV_1 octombrie 2024 finantare.xlsx";
            var excelReader2 = new FormatiiStudiiFromExcel();
            var fonduriBurse2 = excelReader2.ReadFormatiiStudiiFromExcel(filePath2);
            var fonduriBurse2Noi = fonduriBurse2
            .Where(f => !_context.FormatiiStudii.Any(fs =>
                fs.Facultatea == f.Facultatea &&
                fs.ProgramDeStudiu == f.ProgramDeStudiu &&
                fs.An == f.An))
            .ToList();


            try
            {
                if (fonduriBurseNoi.Any()) 
                {
                    _context.FondBurse.AddRange(fonduriBurseNoi);
                    await _context.SaveChangesAsync();
                    return Ok(fonduriBurseNoi);
                }
                else if (fonduriBurse2Noi.Any())
                {
                    _context.FormatiiStudii.AddRange(fonduriBurse2);
                    await _context.SaveChangesAsync();
                    return Ok(fonduriBurse2);
                }
                else
                {
                    return Ok("Nu au fost găsite fonduri noi de adăugat.");
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }


        [HttpGet("{id}")]
        public async Task<IActionResult> GetFondBurseById(int id)
        {
            var fondBurse = await _context.FondBurse.FindAsync(id);

            if (fondBurse == null)
            {
                return NotFound();
            }

            return Ok(fondBurse);
        }
        [HttpGet("generate")]
        public async Task<IActionResult> GenerateExcel(int disponibilBM = 1671770)
        {
            try
            {
                List<FondBurse> fonduri = await _fondBurseService.GetDateFromBursePerformanteAsync();
                List<FormatiiStudii> formatiiStudii = await _fondBurseService.GetAllFromFormatiiStudiiAsync();
                // 📌 Calea temporară unde fișierul va fi generat
                string filePath = Path.Combine(Path.GetTempPath(), "Burse_Studenți.xlsx");

                // Generăm fișierul Excel
                GenerateCustomLayout2(filePath, fonduri, formatiiStudii, disponibilBM);

                // Citim fișierul și îl returnăm ca răspuns HTTP
                byte[] fileBytes = await System.IO.File.ReadAllBytesAsync(filePath);
                 
                // Returnăm fișierul ca `FileContentResult` pentru descărcare
                return File(fileBytes,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            "Burse_Studenți.xlsx");
            }
            catch (Exception ex)
            {
                return BadRequest($"❌ Eroare la generarea fișierului: {ex.Message}");
            }
        }
        static async void GenerateCustomLayout2(string filePath, List<FondBurse> fonduri, List<FormatiiStudii> formatiiStudii, int disponibilBM)
        {
            // 1) Licență EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var generator = new AcronymGenerator();
            // 2) Ștergem fișierul vechi dacă există
            FileInfo fi = new FileInfo(filePath);
            if (fi.Exists)
            {
                fi.Delete();
            }
            int totalFiesc = formatiiStudii
             .Where(f => f.ProgramDeStudiu.Trim().Equals("Total FIESC", StringComparison.OrdinalIgnoreCase))
             .Sum(f => (int.TryParse(f.FaraTaxaRomani, out int rom) ? rom : 0) +
                       (int.TryParse(f.FaraTaxaRp, out int rp) ? rp : 0) +
                       (int.TryParse(f.FaraTaxaUECEE, out int ue) ? ue : 0));

            using (ExcelPackage package = new ExcelPackage(fi))
            {
                // 3) Creăm foaia de lucru
                var sheet = package.Workbook.Worksheets.Add("Burse 2024-2025");
                sheet.Cells.Style.WrapText = true;

                // ---------------------------------------------------------
                // A) Îmbinare și text pentru Program de studiu (A16:A19)
                // ---------------------------------------------------------
                sheet.Cells["A16:A19"].Merge = true;
                sheet.Cells["A16"].Value = "Program de studiu";
                sheet.Cells["A16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["A16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells["A16"].Style.Font.Bold = true;

                // ---------------------------------------------------------
                // B) Îmbinare și text pentru Nr. Studenți Buget (B16:B19)
                // ---------------------------------------------------------
                sheet.Cells["B16:B19"].Merge = true;
                sheet.Cells["B16"].Value = "Nr. Studenți Buget";
                sheet.Cells["B16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["B16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells["B16"].Style.Font.Bold = true;

                // Culoare galbenă
                sheet.Cells["B16:B19"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["B16:B19"].Style.Fill.BackgroundColor.SetColor(Color.Yellow);

                // Rotire text (90 de grade)
                sheet.Cells["B16"].Style.TextRotation = 90;

                // ---------------------------------------------------------
                // C) Fond rep.stud.buget... (C16:C19)
                // ---------------------------------------------------------
                sheet.Cells["C16:C19"].Merge = true;
                sheet.Cells["C16"].Value = "Fond rep.stud.buget pentru bursa de merit, TOTAL";
                sheet.Cells["C16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["C16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells["C16"].Style.Font.Bold = true;
                sheet.Cells["C16"].Style.TextRotation = 90;

                // (Dacă vrei textul și aici rotit, decomentează linia de mai jos)
                // sheet.Cells["C16"].Style.TextRotation = 90;

                // ---------------------------------------------------------
                // D) Burse acordate, 2024/2025 (D16:K16)
                // ---------------------------------------------------------
                sheet.Cells["D16:K16"].Merge = true;
                sheet.Cells["D16"].Value = "Burse acordate, 2024/2025";
                sheet.Cells["D16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["D16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells["D16"].Style.Font.Bold = true;

                // ---------------------------------------------------------
                // E) BM1 (B.Perf.1) (D17:E17)
                // ---------------------------------------------------------
                sheet.Cells["D17:E17"].Merge = true;
                sheet.Cells["D17"].Value = "BM1 (B.Perf.1)";
                sheet.Cells["D17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["D17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells["D17"].Style.Font.Bold = true;

                // ---------------------------------------------------------
                // F) BM2 (B.Perf.2) (F17:G17)
                // ---------------------------------------------------------
                sheet.Cells["F17:G17"].Merge = true;
                sheet.Cells["F17"].Value = "BM2 (B.Perf.2)";
                sheet.Cells["F17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["F17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells["F17"].Style.Font.Bold = true;

                // ---------------------------------------------------------
                // G) 15600 pe D18, 12155 pe D19
                // ---------------------------------------------------------
                sheet.Cells["D18"].Value = fonduri[0].ValoreaLunara * 12;
                sheet.Cells["D19"].Value = fonduri[0].ValoreaLunara * 9.35m;

                // ---------------------------------------------------------
                // H) 14400 pe F18, 11200 pe F19
                // ---------------------------------------------------------
                sheet.Cells["F18"].Value = fonduri[1].ValoreaLunara * 12;
                sheet.Cells["F19"].Value = fonduri[1].ValoreaLunara * 9.35m;

                // ---------------------------------------------------------
                // I) Cheltuit bursa de merit (H17:H19)
                // ---------------------------------------------------------
                sheet.Cells["H17:H19"].Merge = true;
                sheet.Cells["H17"].Value = "Cheltuit bursa de merit";
                sheet.Cells["H17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["H17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells["H17"].Style.Font.Bold = true;

                // ---------------------------------------------------------
                // J) Dif. (I17:I19)
                // ---------------------------------------------------------
                sheet.Cells["I17:I19"].Merge = true;
                sheet.Cells["I17"].Value = "Dif.";
                sheet.Cells["I17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["I17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells["I17"].Style.Font.Bold = true;

                // ---------------------------------------------------------
                // K) Burse acordate de merit (J17:J19)
                // ---------------------------------------------------------
                sheet.Cells["J17:J19"].Merge = true;
                sheet.Cells["J17"].Value = "Burse acordate de merit";
                sheet.Cells["J17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["J17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells["J17"].Style.Font.Bold = true;

                // ---------------------------------------------------------
                // L) Fond ramas pe program (K17:K19)
                // ---------------------------------------------------------
                sheet.Cells["K17:K19"].Merge = true;
                sheet.Cells["K17"].Value = "Fond ramas pe program";
                sheet.Cells["K17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["K17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells["K17"].Style.Font.Bold = true;

                // ---------------------------------------------------------
                // (Opțional) Ajustăm lățimile coloanelor
                // ---------------------------------------------------------
                for (int col = 1; col <= 11; col++)
                {
                    sheet.Column(col).AutoFit();
                }

                // ---------------------------------------------------------
                // (Opțional) Adăugăm borduri pe tot intervalul
                // ---------------------------------------------------------
                // Intervalul cuprinde A16:K19
                using (var range = sheet.Cells["A16:K19"])
                {
                    range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                }
                int startRow = 20;
                int currentRow = startRow;

                // Vom împărți datele în două grupuri: grup L și grup M.
                List<FormatiiStudii> groupL = new List<FormatiiStudii>();
                List<FormatiiStudii> groupM = new List<FormatiiStudii>();

                // Procesăm lista: 
                // Dacă ProgramDeStudiu este "Total FIESC", acesta marchează sfârșitul grupului L.
                // Dacă "An" este "An invalid", se trece peste rând.
                bool groupLCompleted = false;
                foreach (var record in formatiiStudii)
                {
                    // Dacă ProgramDeStudiu este "Total FIESC", marchez sfârșitul grupului L și nu îl adaug.
                    if (record.ProgramDeStudiu.Trim().Equals("Total FIESC", StringComparison.OrdinalIgnoreCase))
                    {
                        groupLCompleted = true;
                        continue;
                    }
                    // Dacă "An" este "An invalid", treci peste rând.
                    if (record.An.Trim().Equals("An invalid", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    if (!groupLCompleted)
                        groupL.Add(record);
                    else
                        groupM.Add(record);
                }
                #region TOTAL L
                // Scriem grupul L
                int groupLStartRow = currentRow;
                foreach (var rec in groupL)
                {
                    // Coloana A: ProgramDeStudiu
                    sheet.Cells[currentRow, 1].Value = generator.GenerateAcronym(rec.ProgramDeStudiu, rec.An);

                    sheet.Cells[currentRow, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[currentRow, 2].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);

                    // Coloana B: suma valorilor din FaraTaxaRomani, FaraTaxaRp, FaraTaxaUECEE
                    int valRom = int.TryParse(rec.FaraTaxaRomani, out int r) ? r : 0;
                    int valRp = int.TryParse(rec.FaraTaxaRp, out int rp) ? rp : 0;
                    int valU = int.TryParse(rec.FaraTaxaUECEE, out int u) ? u : 0;
                    sheet.Cells[currentRow, 2].Value = valRom + valRp + valU;
                    currentRow++;
                }
                int groupLEndRow = currentRow - 1;
                //int totalFiescRow = currentRow;
                for (int row = startRow; row < currentRow; row++)
                {
                    // Coloana C (Fond rep. stud. buget pentru bursa de merit, TOTAL)
                    sheet.Cells[row, 3].Formula = $"({disponibilBM}/{totalFiesc})*B{row}";

                    // Extragem anul din coloana A (ex. "C (4)")
                    string programStudiu = sheet.Cells[row, 1].Value?.ToString();
                    int an = 0;

                    // Extragem numărul anului folosind expresie regulată
                    Match match = Regex.Match(programStudiu ?? "", @"\((\d+)\)");
                    if (match.Success)
                    {
                        an = int.Parse(match.Groups[1].Value);
                    }

                    // Verificăm dacă suntem la rândurile Total L, Total M sau Total FIESC
                    if (sheet.Cells[row, 1].Value?.ToString() == "Total L")
                    {
                        // Formula specifică pentru Total L
                        sheet.Cells[row, 8].Formula = $"SUM(H{groupLStartRow}:H{groupLEndRow})";
                    }
                    else
                    {
                        // Dacă anul este 4, folosim referințe speciale ($D$19 etc.)
                        if (an == 4)
                        {
                            sheet.Cells[row, 8].Formula = $"D{row}*$D$19 + E{row}*$E$19 + F{row}*$F$19 + G{row}*$G$19";
                        }
                        else
                        {
                            // Formula standard pentru cheltuieli bursă
                            sheet.Cells[row, 8].Formula = $"D{row}*$D$18 + E{row}*$E$18 + F{row}*$F$18 + G{row}*$G$18";
                        }
                    }

                    // Coloana I (Diferența dintre fondurile alocate și cheltuite)
                    sheet.Cells[row, 9].Formula = $"C{row}-H{row}";

                    // Coloana J (Suma valorilor din D:G)
                    sheet.Cells[row, 10].Formula = $"SUM(D{row}:G{row})";
                }

                Dictionary<string, List<int>> programRowMap = new Dictionary<string, List<int>>();

                // Regex pentru a elimina doar anii între paranteze (1), (2), (3), (4), dar păstrând "-DUAL" intact
                Regex regex = new Regex(@"(.*?)\s\(\d+\)(-DUAL)?$");

                for (int row = startRow; row < currentRow; row++)
                {
                    string programFull = sheet.Cells[row, 1].Value?.ToString();
                    if (string.IsNullOrEmpty(programFull))
                        continue;
                    if (programFull.Contains("Total"))
                    {
                        if (!programRowMap.ContainsKey(programFull))
                        {
                            programRowMap[programFull] = new List<int>();
                        }
                        programRowMap[programFull].Add(row);
                        continue; // Sărim regex-ul pentru Total-uri
                    }
                    // Aplicăm regex-ul: eliminăm (1), (2), (3), (4), dar păstrăm "-DUAL" dacă există
                    Match match = regex.Match(programFull);
                    string programShort = match.Groups[1].Value.Trim(); // Extragem numele de bază
                    string dualSuffix = match.Groups[2].Value.Trim();  // Verificăm dacă are "-DUAL"

                    // Combinăm numele programului cu "-DUAL" dacă există
                    if (!string.IsNullOrEmpty(dualSuffix))
                    {
                        programShort += dualSuffix;
                    }

                    if (string.IsNullOrEmpty(programShort))
                        continue;

                    // Adăugăm rândul în grupul corespunzător
                    if (!programRowMap.ContainsKey(programShort))
                    {
                        programRowMap[programShort] = new List<int>();
                    }
                    programRowMap[programShort].Add(row);
                }


                // Aplicăm SUM() și MERGE() pentru fiecare grup
                foreach (var entry in programRowMap)
                {
                    List<int> rows = entry.Value;
                    if (entry.Key.Contains("Total"))
                    {
                        // 🔹 FORMULE SPECIALE PENTRU TOTALURI 🔹
                        int totalRow = rows.First(); // Totalurile au doar un singur rând

                        if (entry.Key == "Total L")
                            sheet.Cells[totalRow, 11].Formula = $"SUM(K{startRow}:K{totalRow - 1})";
                    }
                    else
                    {
                        if (rows.Count > 1) // Dacă sunt mai multe rânduri, facem SUM() și merge
                        {
                            int firstRow = rows.First();
                            int lastRow = rows.Last();

                            // Aplicăm formula SUM() în prima coloană de grup (Coloana K)
                            sheet.Cells[firstRow, 11].Formula = $"SUM(I{firstRow}:I{lastRow})";

                            // Facem merge pe toate rândurile
                            string mergeRange = $"K{firstRow}:K{lastRow}";
                            sheet.Cells[mergeRange].Merge = true;
                            sheet.Cells[mergeRange].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        }
                        else // Dacă avem un singur rând (ex. "-DUAL"), tot aplicăm formula
                        {
                            int singleRow = rows.First();

                            // Formula va fi identică cu valoarea din coloana "I"
                            sheet.Cells[singleRow, 11].Formula = $"I{singleRow}";
                        }
                    }
                }


                // Inserăm rândul Total L (o singură dată, dacă există date în grup L)
                if (groupL.Any())
                {
                    sheet.Cells[currentRow, 1].Value = "Total L";

                    // Aplicăm SUM() pentru toate coloanele de la B la K
                    for (char col = 'B'; col <= 'K'; col++)
                    {
                        sheet.Cells[currentRow, col - 'A' + 1].Formula = $"SUM({col}{groupLStartRow}:{col}{groupLEndRow})";
                    }
                    // Stilizare
                    sheet.Cells[currentRow, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[currentRow, 2].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
                    sheet.Cells[currentRow, 1, currentRow, sheet.Dimension.End.Column].Style.Font.Bold = true;
                    sheet.Cells[currentRow, 1, currentRow, sheet.Dimension.End.Column].Style.Font.Size = 12;

                    currentRow++;
                }
                #endregion TOTAL L



                #region TOTAL M
                // Scriem grupul M
                int groupMStartRow = currentRow;
                foreach (var rec in groupM)
                {
                    sheet.Cells[currentRow, 1].Value = generator.GenerateAcronym(rec.ProgramDeStudiu, rec.An);


                    sheet.Cells[currentRow, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[currentRow, 2].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);

                    int valRom = int.TryParse(rec.FaraTaxaRomani, out int r) ? r : 0;
                    int valRp = int.TryParse(rec.FaraTaxaRp, out int rp) ? rp : 0;
                    int valU = int.TryParse(rec.FaraTaxaUECEE, out int u) ? u : 0;
                    sheet.Cells[currentRow, 2].Value = valRom + valRp + valU;
                    currentRow++;
                }
                for (int row = groupMStartRow; row < currentRow; row++)
                {
                    // Coloana C (Fond rep. stud. buget pentru bursa de merit, TOTAL)
                    sheet.Cells[row, 3].Formula = $"({disponibilBM}/{totalFiesc})*B{row}";

                    // Extragem anul din coloana A (ex. "C (4)")
                    string programStudiu = sheet.Cells[row, 1].Value?.ToString();
                    int an = 0;

                    // Extragem numărul anului folosind expresie regulată
                    Match match = Regex.Match(programStudiu ?? "", @"\((\d+)\)");
                    if (match.Success)
                    {
                        an = int.Parse(match.Groups[1].Value);
                    }

                    // Verificăm dacă suntem la rândurile Total L, Total M sau Total FIESC
                    if (sheet.Cells[row, 1].Value?.ToString() == "Total L")
                    {
                        // Formula specifică pentru Total L
                        sheet.Cells[row, 8].Formula = $"SUM(H{groupLStartRow}:H{groupLEndRow})";
                    }
                    else
                    {
                        // Dacă anul este 2, folosim referințe speciale ($D$19 etc.)
                        if (an == 2)
                        {
                            sheet.Cells[row, 8].Formula = $"D{row}*$D$19 + E{row}*$E$19 + F{row}*$F$19 + G{row}*$G$19";
                        }
                        else
                        {
                            // Formula standard pentru cheltuieli bursă
                            sheet.Cells[row, 8].Formula = $"D{row}*$D$18 + E{row}*$E$18 + F{row}*$F$18 + G{row}*$G$18";
                        }
                    }

                    // Coloana I (Diferența dintre fondurile alocate și cheltuite)
                    sheet.Cells[row, 9].Formula = $"C{row}-H{row}";

                    // Coloana J (Suma valorilor din D:G)
                    sheet.Cells[row, 10].Formula = $"SUM(D{row}:G{row})";
                }

                for (int row = startRow; row < currentRow; row++)
                {
                    string programFull = sheet.Cells[row, 1].Value?.ToString();
                    if (string.IsNullOrEmpty(programFull))
                        continue;
                    if (programFull.Contains("Total"))
                    {
                        if (!programRowMap.ContainsKey(programFull))
                        {
                            programRowMap[programFull] = new List<int>();
                        }
                        programRowMap[programFull].Add(row);
                        continue; // Sărim regex-ul pentru Total-uri
                    }
                    // Aplicăm regex-ul: eliminăm (1), (2), (3), (4), dar păstrăm "-DUAL" dacă există
                    Match match = regex.Match(programFull);
                    string programShort = match.Groups[1].Value.Trim(); // Extragem numele de bază
                    string dualSuffix = match.Groups[2].Value.Trim();  // Verificăm dacă are "-DUAL"

                    // Combinăm numele programului cu "-DUAL" dacă există
                    if (!string.IsNullOrEmpty(dualSuffix))
                    {
                        programShort += dualSuffix;
                    }

                    if (string.IsNullOrEmpty(programShort))
                        continue;

                    // Adăugăm rândul în grupul corespunzător
                    if (!programRowMap.ContainsKey(programShort))
                    {
                        programRowMap[programShort] = new List<int>();
                    }
                    programRowMap[programShort].Add(row);
                }


                // Aplicăm SUM() și MERGE() pentru fiecare grup
                foreach (var entry in programRowMap)
                {
                    List<int> rows = entry.Value;
                    if (entry.Key.Contains("Total"))
                    {
                        // 🔹 FORMULE SPECIALE PENTRU TOTALURI 🔹
                        int totalRow = rows.First(); // Totalurile au doar un singur rând

                        if ( entry.Key == "Total M")
                            sheet.Cells[totalRow, 11].Formula = $"SUM(K{startRow}:K{totalRow - 1})";
                    }
                    else
                    {
                        if (rows.Count > 1) // Dacă sunt mai multe rânduri, facem SUM() și merge
                        {
                            int firstRow = rows.First();
                            int lastRow = rows.Last();

                            // Aplicăm formula SUM() în prima coloană de grup (Coloana K)
                            sheet.Cells[firstRow, 11].Formula = $"SUM(I{firstRow}:I{lastRow})";

                            // Facem merge pe toate rândurile
                            string mergeRange = $"K{firstRow}:K{lastRow}";
                            sheet.Cells[mergeRange].Merge = true;
                            sheet.Cells[mergeRange].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        }
                        else // Dacă avem un singur rând (ex. "-DUAL"), tot aplicăm formula
                        {
                            int singleRow = rows.First();

                            // Formula va fi identică cu valoarea din coloana "I"
                            sheet.Cells[singleRow, 11].Formula = $"I{singleRow}";
                        }
                    }
                }

                int groupMEndRow = currentRow - 1;
                // Inserăm rândul Total M, doar dacă există date în grup M
                if (groupL.Any())
                {
                    sheet.Cells[currentRow, 1].Value = "Total M";

                    // Aplicăm SUM() pentru toate coloanele de la B la K
                    for (char col = 'B'; col <= 'K'; col++)
                    {
                        sheet.Cells[currentRow, col - 'A' + 1].Formula = $"SUM({col}{groupMStartRow}:{col}{groupMEndRow})";
                    }
                    // Stilizare
                    sheet.Cells[currentRow, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[currentRow, 2].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
                    sheet.Cells[currentRow, 1, currentRow, sheet.Dimension.End.Column].Style.Font.Bold = true;
                    sheet.Cells[currentRow, 1, currentRow, sheet.Dimension.End.Column].Style.Font.Size = 12;

                    currentRow++;
                }
                #endregion TOTAL M

                #region  TOTAL FIESC

                // Inserăm rândul "Total FIESC" care însumează Total L și Total M
                sheet.Cells[currentRow, 1].Value = "Total FIESC";

                int totalLRow = groupLEndRow + 1; // rândul unde a fost scris Total L
                int totalMRow = groupMEndRow + 1; // rândul unde a fost scris Total M

                // Aplicăm SUM() pentru toate coloanele de la B la K
                for (char col = 'B'; col <= 'K'; col++)
                {
                    sheet.Cells[currentRow, col - 'A' + 1].Formula = $"{col}{totalLRow} + {col}{totalMRow}";
                }

                // Stilizare Total FIESC
                sheet.Cells[currentRow, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[currentRow, 2].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
                sheet.Cells[currentRow, 1, currentRow, sheet.Dimension.End.Column].Style.Font.Bold = true;
                sheet.Cells[currentRow, 1, currentRow, sheet.Dimension.End.Column].Style.Font.Size = 12;

                currentRow++;

                #endregion TOTAL FIESC
                /* // La final, inserăm rândul "Total FIESC" care însumează Total L și Total M
                 sheet.Cells[currentRow, 1].Value = "Total FIESC";
                 if (groupL.Any() && groupM.Any())
                 {
                     int totalLRow = groupLEndRow + 1; // rândul unde a fost scris Total L
                     int totalMRow = groupMEndRow + 1;


                     sheet.Cells[currentRow, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                     sheet.Cells[currentRow, 2].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);

                     sheet.Cells[currentRow, 2].Formula = $"B{totalLRow} + B{totalMRow}";
                     sheet.Cells[currentRow, 3].Formula = $"C{totalLRow} + C{totalMRow}"; // Total FIESC în C
                     sheet.Cells[currentRow, 1, currentRow, sheet.Dimension.End.Column].Style.Font.Bold = true;
                     sheet.Cells[currentRow, 1, currentRow, sheet.Dimension.End.Column].Style.Font.Size = 12;
                     sheet.Cells[currentRow, 1].Style.Font.Size = 12;
                 }
                 else if (groupL.Any())
                 {
                     int totalLRow = groupLEndRow + 1;


                     sheet.Cells[currentRow, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                     sheet.Cells[currentRow, 2].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);

                     sheet.Cells[currentRow, 2].Formula = $"B{totalLRow}";
                     sheet.Cells[currentRow, 3].Formula = $"C{totalLRow}";
                 }
                 else if (groupM.Any())
                 {
                     int totalMRow = groupMEndRow + 1;


                     sheet.Cells[currentRow, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                     sheet.Cells[currentRow, 2].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);

                     sheet.Cells[currentRow, 2].Formula = $"B{totalMRow}";
                     sheet.Cells[currentRow, 3].Formula = $"C{totalMRow}";
                 }
                 int totalFiescRow = currentRow;
                 for (int row = startRow; row < currentRow; row++)
                 {
                     // Coloana C (Fond rep. stud. buget pentru bursa de merit, TOTAL)
                     sheet.Cells[row, 3].Formula = $"({disponibilBM}/B{totalFiescRow})*B{row}";

                     // Extragem anul din coloana A (ex. "C (4)")
                     string programStudiu = sheet.Cells[row, 1].Value?.ToString();
                     int an = 0;

                     // Extragem numărul anului folosind expresie regulată
                     Match match = Regex.Match(programStudiu ?? "", @"\((\d+)\)");
                     if (match.Success)
                     {
                         an = int.Parse(match.Groups[1].Value);
                     }

                     // Verificăm dacă suntem la rândurile Total L, Total M sau Total FIESC
                     if (sheet.Cells[row, 1].Value?.ToString() == "Total L")
                     {
                         // Formula specifică pentru Total L
                         sheet.Cells[row, 8].Formula = $"SUM(H{groupLStartRow}:H{groupLEndRow})";
                     }
                     else if (sheet.Cells[row, 1].Value?.ToString() == "Total M")
                     {
                         // Formula specifică pentru Total M
                         sheet.Cells[row, 8].Formula = $"SUM(H{groupMStartRow}:H{groupMEndRow})";
                     }
                     else if (sheet.Cells[row, 1].Value?.ToString() == "Total FIESC")
                     {
                         // Formula specifică pentru Total FIESC
                         sheet.Cells[row, 8].Formula = $"H{groupLEndRow + 1} + H{groupMEndRow + 1}";
                     }
                     else
                     {
                         // Dacă anul este 4, folosim referințe speciale ($D$19 etc.)
                         if (an == 4)
                         {
                             sheet.Cells[row, 8].Formula = $"D{row}*$D$19 + E{row}*$E$19 + F{row}*$F$19 + G{row}*$G$19";
                         }
                         else
                         {
                             // Formula standard pentru cheltuieli bursă
                             sheet.Cells[row, 8].Formula = $"D{row}*$D$18 + E{row}*$E$18 + F{row}*$F$18 + G{row}*$G$18";
                         }
                     }

                     // Coloana I (Diferența dintre fondurile alocate și cheltuite)
                     sheet.Cells[row, 9].Formula = $"C{row}-H{row}";

                     // Coloana J (Suma valorilor din D:G)
                     sheet.Cells[row, 10].Formula = $"SUM(D{row}:G{row})";
                 }
                 // Dicționar pentru a reține programul de studiu și rândurile aferente
                 Dictionary<string, List<int>> programRowMap = new Dictionary<string, List<int>>();

                 // Regex pentru a elimina doar anii între paranteze (1), (2), (3), (4), dar păstrând "-DUAL" intact
                 Regex regex = new Regex(@"(.*?)\s\(\d+\)(-DUAL)?$");

                 for (int row = startRow; row < currentRow; row++)
                 {
                     string programFull = sheet.Cells[row, 1].Value?.ToString();
                     if (string.IsNullOrEmpty(programFull))
                         continue;
                     if (programFull.Contains("Total"))
                     {
                         if (!programRowMap.ContainsKey(programFull))
                         {
                             programRowMap[programFull] = new List<int>();
                         }
                         programRowMap[programFull].Add(row);
                         continue; // Sărim regex-ul pentru Total-uri
                     }
                     // Aplicăm regex-ul: eliminăm (1), (2), (3), (4), dar păstrăm "-DUAL" dacă există
                     Match match = regex.Match(programFull);
                     string programShort = match.Groups[1].Value.Trim(); // Extragem numele de bază
                     string dualSuffix = match.Groups[2].Value.Trim();  // Verificăm dacă are "-DUAL"

                     // Combinăm numele programului cu "-DUAL" dacă există
                     if (!string.IsNullOrEmpty(dualSuffix))
                     {
                         programShort += dualSuffix;
                     }

                     if (string.IsNullOrEmpty(programShort))
                         continue;

                     // Adăugăm rândul în grupul corespunzător
                     if (!programRowMap.ContainsKey(programShort))
                     {
                         programRowMap[programShort] = new List<int>();
                     }
                     programRowMap[programShort].Add(row);
                 }


                 // Aplicăm SUM() și MERGE() pentru fiecare grup
                 foreach (var entry in programRowMap)
                 {
                     List<int> rows = entry.Value;
                     if (entry.Key.Contains("Total"))
                     {
                         // 🔹 FORMULE SPECIALE PENTRU TOTALURI 🔹
                         int totalRow = rows.First(); // Totalurile au doar un singur rând

                         if (entry.Key == "Total L" || entry.Key == "Total M")
                             sheet.Cells[totalRow, 11].Formula = $"SUM(K{startRow}:K{totalRow - 1})";

                         else if (entry.Key == "Total FIESC")
                         {
                             int totalLRow = programRowMap["Total L"].First();
                             int totalMRow = programRowMap["Total M"].First();
                             sheet.Cells[totalRow, 11].Formula = $"K{totalLRow} + K{totalMRow}";
                         }
                     }
                     else
                     {
                         if (rows.Count > 1) // Dacă sunt mai multe rânduri, facem SUM() și merge
                         {
                             int firstRow = rows.First();
                             int lastRow = rows.Last();

                             // Aplicăm formula SUM() în prima coloană de grup (Coloana K)
                             sheet.Cells[firstRow, 11].Formula = $"SUM(I{firstRow}:I{lastRow})";

                             // Facem merge pe toate rândurile
                             string mergeRange = $"K{firstRow}:K{lastRow}";
                             sheet.Cells[mergeRange].Merge = true;
                             sheet.Cells[mergeRange].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                         }
                         else // Dacă avem un singur rând (ex. "-DUAL"), tot aplicăm formula
                         {
                             int singleRow = rows.First();

                             // Formula va fi identică cu valoarea din coloana "I"
                             sheet.Cells[singleRow, 11].Formula = $"I{singleRow}";
                         }
                     }
                 }*/


                 sheet.Cells[currentRow, 1, currentRow, sheet.Dimension.End.Column].Style.Font.Bold = true;
                 sheet.Cells[currentRow, 1, currentRow, sheet.Dimension.End.Column].Style.Font.Size = 12;
                 currentRow++;

                 sheet.Cells[startRow, 3, currentRow - 1, 3].Style.Numberformat.Format = "#,##0.00";
                 sheet.Column(3).AutoFit();
                 sheet.Column(3).Width = 13.57; // Setează lățimea exactă în Excel units


                 // Ajustăm coloanele
                 //sheet.Cells[sheet.Dimension.Address].AutoFitColumns();

                 for (int row = sheet.Dimension.Start.Row; row <= sheet.Dimension.End.Row; row++)
                 {
                     sheet.Row(row).CustomHeight = false;
                 }
                 using (var range = sheet.Cells["A20:C63"])
                 {
                     range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                     range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                     range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                     range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                 }
                sheet.Cells[startRow, 2, currentRow, 11].Style.Numberformat.Format = "#,##0.00";
                // 3) Salvăm fișierul
                package.Save();
            }
        }
        static async void GenerateCustomLayout(string filePath, List<FondBurse> fonduri, List<FormatiiStudii> formatiiStudii,int disponibilBM)
        {
            // 1) Licență EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var generator = new AcronymGenerator();
            // 2) Ștergem fișierul vechi dacă există
            FileInfo fi = new FileInfo(filePath);
            if (fi.Exists)
            {
                fi.Delete();
            }

            using (ExcelPackage package = new ExcelPackage(fi))
            {
                // 3) Creăm foaia de lucru
                var sheet = package.Workbook.Worksheets.Add("Burse 2024-2025");
                sheet.Cells.Style.WrapText = true;

                // ---------------------------------------------------------
                // A) Îmbinare și text pentru Program de studiu (A16:A19)
                // ---------------------------------------------------------
                sheet.Cells["A16:A19"].Merge = true;
                sheet.Cells["A16"].Value = "Program de studiu";
                sheet.Cells["A16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["A16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells["A16"].Style.Font.Bold = true;

                // ---------------------------------------------------------
                // B) Îmbinare și text pentru Nr. Studenți Buget (B16:B19)
                // ---------------------------------------------------------
                sheet.Cells["B16:B19"].Merge = true;
                sheet.Cells["B16"].Value = "Nr. Studenți Buget";
                sheet.Cells["B16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["B16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells["B16"].Style.Font.Bold = true;

                // Culoare galbenă
                sheet.Cells["B16:B19"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["B16:B19"].Style.Fill.BackgroundColor.SetColor(Color.Yellow);

                // Rotire text (90 de grade)
                sheet.Cells["B16"].Style.TextRotation = 90;

                // ---------------------------------------------------------
                // C) Fond rep.stud.buget... (C16:C19)
                // ---------------------------------------------------------
                sheet.Cells["C16:C19"].Merge = true;
                sheet.Cells["C16"].Value = "Fond rep.stud.buget pentru bursa de merit, TOTAL";
                sheet.Cells["C16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["C16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells["C16"].Style.Font.Bold = true;
                sheet.Cells["C16"].Style.TextRotation = 90;

                // (Dacă vrei textul și aici rotit, decomentează linia de mai jos)
                // sheet.Cells["C16"].Style.TextRotation = 90;

                // ---------------------------------------------------------
                // D) Burse acordate, 2024/2025 (D16:K16)
                // ---------------------------------------------------------
                sheet.Cells["D16:K16"].Merge = true;
                sheet.Cells["D16"].Value = "Burse acordate, 2024/2025";
                sheet.Cells["D16"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["D16"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells["D16"].Style.Font.Bold = true;

                // ---------------------------------------------------------
                // E) BM1 (B.Perf.1) (D17:E17)
                // ---------------------------------------------------------
                sheet.Cells["D17:E17"].Merge = true;
                sheet.Cells["D17"].Value = "BM1 (B.Perf.1)";
                sheet.Cells["D17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["D17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells["D17"].Style.Font.Bold = true;

                // ---------------------------------------------------------
                // F) BM2 (B.Perf.2) (F17:G17)
                // ---------------------------------------------------------
                sheet.Cells["F17:G17"].Merge = true;
                sheet.Cells["F17"].Value = "BM2 (B.Perf.2)";
                sheet.Cells["F17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["F17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells["F17"].Style.Font.Bold = true;

                // ---------------------------------------------------------
                // G) 15600 pe D18, 12155 pe D19
                // ---------------------------------------------------------
                sheet.Cells["D18"].Value = fonduri[0].ValoreaLunara * 12;
                sheet.Cells["D19"].Value = fonduri[0].ValoreaLunara * 9.35m;

                // ---------------------------------------------------------
                // H) 14400 pe F18, 11200 pe F19
                // ---------------------------------------------------------
                sheet.Cells["F18"].Value = fonduri[1].ValoreaLunara * 12;
                sheet.Cells["F19"].Value = fonduri[1].ValoreaLunara * 9.35m;

                // ---------------------------------------------------------
                // I) Cheltuit bursa de merit (H17:H19)
                // ---------------------------------------------------------
                sheet.Cells["H17:H19"].Merge = true;
                sheet.Cells["H17"].Value = "Cheltuit bursa de merit";
                sheet.Cells["H17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["H17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells["H17"].Style.Font.Bold = true;

                // ---------------------------------------------------------
                // J) Dif. (I17:I19)
                // ---------------------------------------------------------
                sheet.Cells["I17:I19"].Merge = true;
                sheet.Cells["I17"].Value = "Dif.";
                sheet.Cells["I17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["I17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells["I17"].Style.Font.Bold = true;

                // ---------------------------------------------------------
                // K) Burse acordate de merit (J17:J19)
                // ---------------------------------------------------------
                sheet.Cells["J17:J19"].Merge = true;
                sheet.Cells["J17"].Value = "Burse acordate de merit";
                sheet.Cells["J17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["J17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells["J17"].Style.Font.Bold = true;

                // ---------------------------------------------------------
                // L) Fond ramas pe program (K17:K19)
                // ---------------------------------------------------------
                sheet.Cells["K17:K19"].Merge = true;
                sheet.Cells["K17"].Value = "Fond ramas pe program";
                sheet.Cells["K17"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["K17"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells["K17"].Style.Font.Bold = true;

                // ---------------------------------------------------------
                // (Opțional) Ajustăm lățimile coloanelor
                // ---------------------------------------------------------
                for (int col = 1; col <= 11; col++)
                {
                    sheet.Column(col).AutoFit();
                }

                // ---------------------------------------------------------
                // (Opțional) Adăugăm borduri pe tot intervalul
                // ---------------------------------------------------------
                // Intervalul cuprinde A16:K19
                using (var range = sheet.Cells["A16:K19"])
                {
                    range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                }
                int startRow = 20;
                int currentRow = startRow;

                // Vom împărți datele în două grupuri: grup L și grup M.
                List<FormatiiStudii> groupL = new List<FormatiiStudii>();
                List<FormatiiStudii> groupM = new List<FormatiiStudii>();

                // Procesăm lista: 
                // Dacă ProgramDeStudiu este "Total FIESC", acesta marchează sfârșitul grupului L.
                // Dacă "An" este "An invalid", se trece peste rând.
                bool groupLCompleted = false;
                foreach (var record in formatiiStudii)
                {
                    // Dacă ProgramDeStudiu este "Total FIESC", marchez sfârșitul grupului L și nu îl adaug.
                    if (record.ProgramDeStudiu.Trim().Equals("Total FIESC", StringComparison.OrdinalIgnoreCase))
                    {
                        groupLCompleted = true;
                        continue;
                    }
                    // Dacă "An" este "An invalid", treci peste rând.
                    if (record.An.Trim().Equals("An invalid", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    if (!groupLCompleted)
                        groupL.Add(record);
                    else
                        groupM.Add(record);
                }

                // Scriem grupul L
                int groupLStartRow = currentRow;
                foreach (var rec in groupL)
                {
                    // Coloana A: ProgramDeStudiu
                    sheet.Cells[currentRow, 1].Value = generator.GenerateAcronym(rec.ProgramDeStudiu, rec.An);

                    sheet.Cells[currentRow,2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[currentRow,2].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);

                    // Coloana B: suma valorilor din FaraTaxaRomani, FaraTaxaRp, FaraTaxaUECEE
                    int valRom = int.TryParse(rec.FaraTaxaRomani, out int r) ? r : 0;
                    int valRp = int.TryParse(rec.FaraTaxaRp, out int rp) ? rp : 0;
                    int valU = int.TryParse(rec.FaraTaxaUECEE, out int u) ? u : 0;
                    sheet.Cells[currentRow, 2].Value = valRom + valRp + valU;
                    currentRow++;
                }
                int groupLEndRow = currentRow - 1;

                // Inserăm rândul Total L (o singură dată, dacă există date în grup L)
                if (groupL.Any())
                {
                    sheet.Cells[currentRow, 1].Value = "Total L";
                    sheet.Cells[currentRow, 2].Formula = $"SUM(B{groupLStartRow}:B{groupLEndRow})";
                    sheet.Cells[currentRow, 3].Formula = $"SUM(C{groupLStartRow}:C{groupLEndRow})"; // SUM pentru fonduri alocate


                    sheet.Cells[currentRow, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[currentRow, 2].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);

                    sheet.Cells[currentRow, 1, currentRow, sheet.Dimension.End.Column].Style.Font.Bold = true;
                    sheet.Cells[currentRow, 1, currentRow, sheet.Dimension.End.Column].Style.Font.Size = 12;
                    currentRow++;
                }

                // Scriem grupul M
                int groupMStartRow = currentRow;
                foreach (var rec in groupM)
                {
                    sheet.Cells[currentRow, 1].Value = generator.GenerateAcronym(rec.ProgramDeStudiu,rec.An);


                    sheet.Cells[currentRow, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[currentRow, 2].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);

                    int valRom = int.TryParse(rec.FaraTaxaRomani, out int r) ? r : 0;
                    int valRp = int.TryParse(rec.FaraTaxaRp, out int rp) ? rp : 0;
                    int valU = int.TryParse(rec.FaraTaxaUECEE, out int u) ? u : 0;
                    sheet.Cells[currentRow, 2].Value = valRom + valRp + valU;
                    currentRow++;
                }
                int groupMEndRow = currentRow - 1;

                // Inserăm rândul Total M, doar dacă există date în grup M
                if (groupM.Any())
                {
                    sheet.Cells[currentRow, 1].Value = "Total M";
                    sheet.Cells[currentRow, 2].Formula = $"SUM(B{groupMStartRow}:B{groupMEndRow})";
                    sheet.Cells[currentRow, 3].Formula = $"SUM(C{groupMStartRow}:C{groupMEndRow})"; // SUM pentru fonduri alocate


                    sheet.Cells[currentRow, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[currentRow, 2].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);

                    sheet.Cells[currentRow, 1, currentRow, sheet.Dimension.End.Column].Style.Font.Bold = true;
                    sheet.Cells[currentRow, 1, currentRow, sheet.Dimension.End.Column].Style.Font.Size = 12;
                    currentRow++;
                }

                // La final, inserăm rândul "Total FIESC" care însumează Total L și Total M
                sheet.Cells[currentRow, 1].Value = "Total FIESC";
                if (groupL.Any() && groupM.Any())
                {
                    int totalLRow = groupLEndRow + 1; // rândul unde a fost scris Total L
                    int totalMRow = groupMEndRow + 1;


                    sheet.Cells[currentRow, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[currentRow, 2].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);

                    sheet.Cells[currentRow, 2].Formula = $"B{totalLRow} + B{totalMRow}";
                    sheet.Cells[currentRow, 3].Formula = $"C{totalLRow} + C{totalMRow}"; // Total FIESC în C
                    sheet.Cells[currentRow, 1, currentRow, sheet.Dimension.End.Column].Style.Font.Bold = true;
                    sheet.Cells[currentRow, 1, currentRow, sheet.Dimension.End.Column].Style.Font.Size = 12;
                    sheet.Cells[currentRow, 1].Style.Font.Size = 12;
                }
                else if (groupL.Any())
                {
                    int totalLRow = groupLEndRow + 1;


                    sheet.Cells[currentRow, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[currentRow, 2].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);

                    sheet.Cells[currentRow, 2].Formula = $"B{totalLRow}";
                    sheet.Cells[currentRow, 3].Formula = $"C{totalLRow}";
                }
                else if (groupM.Any())
                {
                    int totalMRow = groupMEndRow + 1;


                    sheet.Cells[currentRow, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[currentRow, 2].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);

                    sheet.Cells[currentRow, 2].Formula = $"B{totalMRow}";
                    sheet.Cells[currentRow, 3].Formula = $"C{totalMRow}";
                }
                int totalFiescRow = currentRow;
                for (int row = startRow; row < currentRow; row++)
                {
                    // Coloana C (Fond rep. stud. buget pentru bursa de merit, TOTAL)
                    sheet.Cells[row, 3].Formula = $"({disponibilBM}/B{totalFiescRow})*B{row}";

                    // Extragem anul din coloana A (ex. "C (4)")
                    string programStudiu = sheet.Cells[row, 1].Value?.ToString();
                    int an = 0;

                    // Extragem numărul anului folosind expresie regulată
                    Match match = Regex.Match(programStudiu ?? "", @"\((\d+)\)");
                    if (match.Success)
                    {
                        an = int.Parse(match.Groups[1].Value);
                    }

                    // Verificăm dacă suntem la rândurile Total L, Total M sau Total FIESC
                    if (sheet.Cells[row, 1].Value?.ToString() == "Total L")
                    {
                        // Formula specifică pentru Total L
                        sheet.Cells[row, 8].Formula = $"SUM(H{groupLStartRow}:H{groupLEndRow})";
                    }
                    else if (sheet.Cells[row, 1].Value?.ToString() == "Total M")
                    {
                        // Formula specifică pentru Total M
                        sheet.Cells[row, 8].Formula = $"SUM(H{groupMStartRow}:H{groupMEndRow})";
                    }
                    else if (sheet.Cells[row, 1].Value?.ToString() == "Total FIESC")
                    {
                        // Formula specifică pentru Total FIESC
                        sheet.Cells[row, 8].Formula = $"H{groupLEndRow + 1} + H{groupMEndRow + 1}";
                    }
                    else
                    {
                        // Dacă anul este 4, folosim referințe speciale ($D$19 etc.)
                        if (an == 4)
                        {
                            sheet.Cells[row, 8].Formula = $"D{row}*$D$19 + E{row}*$E$19 + F{row}*$F$19 + G{row}*$G$19";
                        }
                        else
                        {
                            // Formula standard pentru cheltuieli bursă
                            sheet.Cells[row, 8].Formula = $"D{row}*$D$18 + E{row}*$E$18 + F{row}*$F$18 + G{row}*$G$18";
                        }
                    }

                    // Coloana I (Diferența dintre fondurile alocate și cheltuite)
                    sheet.Cells[row, 9].Formula = $"C{row}-H{row}";

                    // Coloana J (Suma valorilor din D:G)
                    sheet.Cells[row, 10].Formula = $"SUM(D{row}:G{row})";
                }
                // Dicționar pentru a reține programul de studiu și rândurile aferente
                Dictionary<string, List<int>> programRowMap = new Dictionary<string, List<int>>();

                // Regex pentru a elimina doar anii între paranteze (1), (2), (3), (4), dar păstrând "-DUAL" intact
                Regex regex = new Regex(@"(.*?)\s\(\d+\)(-DUAL)?$");

                for (int row = startRow; row < currentRow; row++)
                {
                    string programFull = sheet.Cells[row, 1].Value?.ToString();
                    if (string.IsNullOrEmpty(programFull))
                        continue;
                    if (programFull.Contains("Total"))
                    {
                        if (!programRowMap.ContainsKey(programFull))
                        {
                            programRowMap[programFull] = new List<int>();
                        }
                        programRowMap[programFull].Add(row);
                        continue; // Sărim regex-ul pentru Total-uri
                    }
                    // Aplicăm regex-ul: eliminăm (1), (2), (3), (4), dar păstrăm "-DUAL" dacă există
                    Match match = regex.Match(programFull);
                    string programShort = match.Groups[1].Value.Trim(); // Extragem numele de bază
                    string dualSuffix = match.Groups[2].Value.Trim();  // Verificăm dacă are "-DUAL"

                    // Combinăm numele programului cu "-DUAL" dacă există
                    if (!string.IsNullOrEmpty(dualSuffix))
                    {
                        programShort += dualSuffix;
                    }

                    if (string.IsNullOrEmpty(programShort))
                        continue;

                    // Adăugăm rândul în grupul corespunzător
                    if (!programRowMap.ContainsKey(programShort))
                    {
                        programRowMap[programShort] = new List<int>();
                    }
                    programRowMap[programShort].Add(row);
                }


                // Aplicăm SUM() și MERGE() pentru fiecare grup
                foreach (var entry in programRowMap)
                {
                    List<int> rows = entry.Value;
                    if (entry.Key.Contains("Total")) 
                    {
                        // 🔹 FORMULE SPECIALE PENTRU TOTALURI 🔹
                        int totalRow = rows.First(); // Totalurile au doar un singur rând

                        if (entry.Key == "Total L" || entry.Key == "Total M")
                            sheet.Cells[totalRow, 11].Formula = $"SUM(K{startRow}:K{totalRow - 1})";

                        else if (entry.Key == "Total FIESC")
                        {
                            int totalLRow = programRowMap["Total L"].First();
                            int totalMRow = programRowMap["Total M"].First();
                            sheet.Cells[totalRow, 11].Formula = $"K{totalLRow} + K{totalMRow}";
                        }
                    }
                    else
                    {
                        if (rows.Count > 1) // Dacă sunt mai multe rânduri, facem SUM() și merge
                        {
                            int firstRow = rows.First();
                            int lastRow = rows.Last();

                            // Aplicăm formula SUM() în prima coloană de grup (Coloana K)
                            sheet.Cells[firstRow, 11].Formula = $"SUM(I{firstRow}:I{lastRow})";

                            // Facem merge pe toate rândurile
                            string mergeRange = $"K{firstRow}:K{lastRow}";
                            sheet.Cells[mergeRange].Merge = true;
                            sheet.Cells[mergeRange].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        }
                        else // Dacă avem un singur rând (ex. "-DUAL"), tot aplicăm formula
                        {
                            int singleRow = rows.First();

                            // Formula va fi identică cu valoarea din coloana "I"
                            sheet.Cells[singleRow, 11].Formula = $"I{singleRow}";
                        }
                    }
                }
              

                sheet.Cells[currentRow, 1, currentRow, sheet.Dimension.End.Column].Style.Font.Bold = true;
                sheet.Cells[currentRow, 1, currentRow, sheet.Dimension.End.Column].Style.Font.Size = 12;
                currentRow++;

                sheet.Cells[startRow, 3, currentRow - 1, 3].Style.Numberformat.Format = "#,##0.00";
                sheet.Column(3).AutoFit();
                sheet.Column(3).Width = 13.57; // Setează lățimea exactă în Excel units


                // Ajustăm coloanele
                //sheet.Cells[sheet.Dimension.Address].AutoFitColumns();

                for (int row = sheet.Dimension.Start.Row; row <= sheet.Dimension.End.Row; row++)
                {
                    sheet.Row(row).CustomHeight = false;
                }
                using (var range = sheet.Cells["A20:C63"])
                {
                    range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                }

                //sheet.Cells[startRow, 2, currentRow, 3].Style.Numberformat.Format = "#,##0.00";
                // 3) Salvăm fișierul
                package.Save();
            }
        }
    }
}
