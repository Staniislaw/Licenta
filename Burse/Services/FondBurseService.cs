using Burse.Data;
using Burse.Helpers;
using Burse.Models;
using Burse.Services.Abstractions;

using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml.Style;
using OfficeOpenXml;
using System.Drawing;

using System.Text.RegularExpressions;

namespace Burse.Services
{
    public class FondBurseService : IFondBurseService
    {
        private readonly BurseDBContext _context;
        private readonly IFondBurseMeritRepartizatService _fondBurseMeritRepartizatService;

        public FondBurseService(BurseDBContext context, IFondBurseMeritRepartizatService fondBurseMeritRepartizatService)
        {
            _context = context;
            _fondBurseMeritRepartizatService = fondBurseMeritRepartizatService;
        }

        public async Task<List<FondBurse>> GetDateFromBursePerformanteAsync()
        {
            var fonduri = await _context.FondBurse
                .Where(f => f.CategorieBurse == "Bursa de performanță 1" ||
                            f.CategorieBurse == "Bursa de performanță 2")
                .ToListAsync();

            return fonduri;
        }
        public async Task<List<FormatiiStudii>> GetAllFromFormatiiStudiiAsync()
        {
            var formatiiStudii = await _context.FormatiiStudii.ToListAsync();

            return formatiiStudii;
        }
        public async Task<byte[]> GenerateCustomLayout2(string filePath, List<FondBurse> fonduri, List<FormatiiStudii> formatiiStudii, decimal disponibilBM)
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
                    // ✅ Generate domain name
                    string domeniu = generator.GenerateAcronym(AcronymGenerator.RemoveDiacritics(rec.ProgramDeStudiu), rec.An);

                    // ✅ Compute the sum for the scholarship fund
                    int valRom = int.TryParse(rec.FaraTaxaRomani, out int r) ? r : 0;
                    int valRp = int.TryParse(rec.FaraTaxaRp, out int rp) ? rp : 0;
                    int valU = int.TryParse(rec.FaraTaxaUECEE, out int u) ? u : 0;
                    decimal totalFond = valRom + valRp + valU;

                    // ✅ Write values to Excel
                    sheet.Cells[currentRow, 1].Value = domeniu;

                    sheet.Cells[currentRow, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[currentRow, 2].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);

                    sheet.Cells[currentRow, 2].Value = totalFond;

                    // ✅ Create and add record to the database
                    string grupa = GrupuriDomeniiHelper.GetGrupa(domeniu);

                    var fondBursa = new FondBurseMeritRepartizat
                    {
                        domeniu = domeniu,
                        bursaAlocatata = disponibilBM / totalFiesc * totalFond,
                        programStudiu = "licenta",
                        Grupa = grupa,
                    };
                    await _fondBurseMeritRepartizatService.AddAsync(fondBursa);

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
                    /*Match match = regex.Match(programFull);
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
                    programRowMap[programShort].Add(row);*/
                    string grupa = GrupuriDomeniiHelper.GetGrupa(programFull);
                    if (!programRowMap.ContainsKey(grupa))
                    {
                        programRowMap[grupa] = new List<int>();
                    }
                    programRowMap[grupa].Add(row);
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
                    string domeniu = generator.GenerateAcronym(AcronymGenerator.RemoveDiacritics(rec.ProgramDeStudiu), rec.An);

                    // ✅ Compute the sum for the scholarship fund
                    int valRom = int.TryParse(rec.FaraTaxaRomani, out int r) ? r : 0;
                    int valRp = int.TryParse(rec.FaraTaxaRp, out int rp) ? rp : 0;
                    int valU = int.TryParse(rec.FaraTaxaUECEE, out int u) ? u : 0;
                    decimal totalFond = valRom + valRp + valU;

                    sheet.Cells[currentRow, 1].Value = domeniu;


                    sheet.Cells[currentRow, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[currentRow, 2].Style.Fill.BackgroundColor.SetColor(Color.LightYellow);

                   
                    sheet.Cells[currentRow, 2].Value = totalFond;
                    string grupa = GrupuriDomeniiHelper.GetGrupa(domeniu);
                    var fondBursa = new FondBurseMeritRepartizat
                    {
                        domeniu = domeniu,
                        bursaAlocatata = disponibilBM / totalFiesc * totalFond,
                        programStudiu = "master",
                        Grupa = grupa,
                    };



                    await _fondBurseMeritRepartizatService.AddAsync(fondBursa);
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
                    string grupa = GrupuriDomeniiHelper.GetGrupa(programFull);
                    if (!programRowMap.ContainsKey(grupa))
                    {
                        programRowMap[grupa] = new List<int>();
                    }
                    programRowMap[grupa].Add(row);

                    /*Match match = regex.Match(programFull);
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
                    programRowMap[programShort].Add(row);*/
                }


                // Aplicăm SUM() și MERGE() pentru fiecare grup
                foreach (var entry in programRowMap)
                {
                    List<int> rows = entry.Value;
                    if (entry.Key.Contains("Total"))
                    {
                        // 🔹 FORMULE SPECIALE PENTRU TOTALURI 🔹
                        int totalRow = rows.First(); // Totalurile au doar un singur rând

                        if (entry.Key == "Total M")
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

                sheet.Cells[currentRow, 1, currentRow, sheet.Dimension.End.Column].Style.Font.Bold = true;
                sheet.Cells[currentRow, 1, currentRow, sheet.Dimension.End.Column].Style.Font.Size = 12;
                currentRow++;

                sheet.Cells[startRow, 3, currentRow - 1, 3].Style.Numberformat.Format = "#,##0.00";
                sheet.Column(3).AutoFit();
                sheet.Column(3).Width = 13.57; // Setează lățimea exactă în Excel units
                sheet.Column(9).AutoFit();
                sheet.Column(9).Width = 13.57; // Setează lățimea exactă în Excel units
                sheet.Column(11).AutoFit();
                sheet.Column(11).Width = 13.57; // Setează lățimea exactă în Excel units


                // Ajustăm coloanele
                //sheet.Cells[sheet.Dimension.Address].AutoFitColumns();

                for (int row = sheet.Dimension.Start.Row; row <= sheet.Dimension.End.Row; row++)
                {
                    sheet.Row(row).CustomHeight = false;
                }
                using (var range = sheet.Cells["A20:K63"])
                {
                    range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                }
                sheet.Cells[startRow, 2, currentRow, 11].Style.Numberformat.Format = "#,##0.00";
                // 3) Salvăm fișierul
                //package.Save();
                using (var memoryStream = new MemoryStream())
                {
                    await package.SaveAsAsync(memoryStream);
                    return memoryStream.ToArray(); // ✅ Return file as byte[]
                }
            }
        }
        public async Task SaveNewStudentsAsync(List<StudentRecord> students)
        {
            // Listă pentru studenții care trebuie adăugați
            var studentsToAdd = new List<StudentRecord>();

            foreach (var student in students)
            {
                // Verifică dacă studentul există deja în baza de date
                bool studentExists = await _context.StudentRecord
                    .AnyAsync(s => s.Emplid == student.Emplid);

                if (!studentExists)
                {
                    // Adaugă studentul în lista de studenți de adăugat
                    studentsToAdd.Add(student);
                }
            }

            // Adaugă studenții care nu există deja
            if (studentsToAdd.Any())
            {
                await _context.StudentRecord.AddRangeAsync(studentsToAdd);
                await _context.SaveChangesAsync();
            }
        }
        public async Task<List<StudentRecord>> GetStudentsWithBursaFromDatabaseAsync()
        {
            return await _context.StudentRecord
                .Where(s => !string.IsNullOrEmpty(s.Bursa) && s.Bursa.ToLower() != "nicio bursă")
                .Include(s => s.FondBurseMeritRepartizat) 
                .ToListAsync();
        }
        public async Task<Dictionary<string, List<StudentRecord>>> GetStudentiEligibiliPeGrupaAsync()
        {
            var studenti = await _context.StudentRecord
                .Include(s => s.FondBurseMeritRepartizat)
                .Where(s => s.FondBurseMeritRepartizat != null)
                .Where(s => s.Bursa == null || s.Bursa == "nicio bursă")
                .ToListAsync();

            return studenti
                .GroupBy(s => s.FondBurseMeritRepartizat.Grupa)
                .ToDictionary(
                    g => g.Key,
                    g => g.OrderByDescending(s => s.Media).ToList()
                );
        }
        public async Task<Dictionary<string, List<StudentRecord>>> GetStudentiEligibiliPeProgramAsync()
        {
            var studenti = await _context.StudentRecord
                .Include(s => s.FondBurseMeritRepartizat)
                .Where(s => s.FondBurseMeritRepartizat != null)
                .Where(s => s.Bursa == null || s.Bursa == "nicio bursă")
                .ToListAsync();

            return studenti
                .GroupBy(s => s.FondBurseMeritRepartizat.programStudiu) // "licenta" / "master"
                .ToDictionary(
                    g => g.Key,
                    g => g.OrderByDescending(s => s.Media).ToList()
                );
        }
        public async Task<List<StudentRecord>> GetStudentiEligibiliPeDomeniiAsync(List<string> domenii)
        {
            return await _context.StudentRecord
                .Include(s => s.FondBurseMeritRepartizat)
                .Where(s => s.FondBurseMeritRepartizat != null)
                .Where(s => domenii.Contains(s.FondBurseMeritRepartizat.Grupa))
                .Where(s => s.FondBurseMeritRepartizat.programStudiu == "licenta")
                .Where(s => s.Bursa == null || s.Bursa == "nicio bursă")
                .ToListAsync();
        }



    }
}
