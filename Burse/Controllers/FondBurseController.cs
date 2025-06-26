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
using Burse.Services;
using Microsoft.EntityFrameworkCore;
using DocumentFormat.OpenXml.Vml.Office;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Authorization;

namespace Burse.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class FondBurseController : ControllerBase
    {
        private readonly BurseDBContext _context;
        private readonly IFondBurseService _fondBurseService;
        private readonly IFondBurseMeritRepartizatService _fondBurseMeritRepartizatService;
        private readonly IBurseIstoricService _burseIstoricService;
        private readonly AppLogger _logger;
        private readonly GrupuriDomeniiHelper _grupuriHelper;
        private readonly IGrupuriService _grupuriService;
        public FondBurseController(BurseDBContext context, IFondBurseService fondBurseService, IFondBurseMeritRepartizatService fondBurseMeritRepartizatService, GrupuriDomeniiHelper 
            grupuriHelper, IBurseIstoricService burseIstoricService, AppLogger logger, IGrupuriService grupuriService)
        {
            _context = context;
            _fondBurseService = fondBurseService;
            _fondBurseMeritRepartizatService = fondBurseMeritRepartizatService;
            _grupuriHelper = grupuriHelper;
            _burseIstoricService = burseIstoricService;
            _logger = logger;
            _grupuriService = grupuriService;
        }
        [Authorize]
        [HttpPost("AddFondBurse")]
        public async Task<IActionResult> AddFondBurse(List<IFormFile> files)
        {
            if (files == null || files.Count == 0)
                return BadRequest("Nu s-au primit fișiere.");
            var fonduriBurseFile = files[0];
            var excelReader = new FondBurseExcelReader();
            List<FondBurse> fonduriBurse;
            using (var stream1 = fonduriBurseFile.OpenReadStream())
            {
                fonduriBurse = excelReader.ReadFondBurseFromExcel(stream1);
            }
            var fonduriBurseNoi = fonduriBurse.Where(f => !_context.FondBurse.Any(fb => fb.CategorieBurse == f.CategorieBurse)).ToList();
            var formatiiStudiiFile = files[1];
            var excelReader2 = new FormatiiStudiiFromExcel();
            List<FormatiiStudii> fonduriBurse2;
            using (var stream2 = formatiiStudiiFile.OpenReadStream())
            {
                fonduriBurse2 = excelReader2.ReadFormatiiStudiiFromExcel(stream2);
            }
            var fonduriBurse2Noi = fonduriBurse2
            .Where(f => !_context.FormatiiStudii.Any(fs =>
                fs.Facultatea == f.Facultatea &&
                fs.ProgramDeStudiu == f.ProgramDeStudiu &&
                fs.An == f.An))
            .ToList();
            try
            {
                bool hasChanges = false;
                if (fonduriBurseNoi.Any())
                {
                    _context.FondBurse.AddRange(fonduriBurseNoi);
                    hasChanges = true;
                }
                if (fonduriBurse2Noi.Any())
                {
                    _context.FormatiiStudii.AddRange(fonduriBurse2Noi);
                    hasChanges = true;
                }
                if (hasChanges)
                {
                    await _context.SaveChangesAsync();
                    return Ok(new { message = "Fondurile noi au fost adăugate cu succes." });

                }
                return Ok(new { message = "Nu au fost găsite fonduri noi de adăugat." });
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }
        [Authorize]
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
        [Authorize]
        [HttpGet("generate")]
        public async Task<IActionResult> GenerateExcel(decimal disponibilBM = 1671770.95m, double? valoareRomaniDePretutindeni = null)
        {
            try
            {
                List<FondBurse> fonduri = await _fondBurseService.GetDateFromBursePerformanteAsync();
                List<FormatiiStudii> formatiiStudii = await _fondBurseService.GetAllFromFormatiiStudiiAsync();
                string filePath = Path.Combine(Path.GetTempPath(), "Burse_Studenți.xlsx");
                byte[] fileBytes= await _fondBurseService.GenerateCustomLayout2(filePath, fonduri, formatiiStudii, disponibilBM, valoareRomaniDePretutindeni);
                return File(fileBytes,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            "Burse_Studenți.xlsx");
            }
            catch (Exception ex)
            {
                return BadRequest($"❌ Eroare la generarea fișierului: {ex.Message}");
            }
        }
        [Authorize]
        [HttpPost("process")]
        public async Task<IActionResult> ProcessExcelFiles([FromForm] List<IFormFile> pathStudentiList, [FromForm] IFormFile burseFile,
            [FromQuery] decimal? epsilonValue = 0.05M, [FromQuery] double? valoareRomaniDePretutindeni = null)
        {
            decimal epsilon = epsilonValue ?? 0.05M;
            if (burseFile == null)
            {
                return BadRequest("Fișierul Burse_Studenti.xlsx nu a fost găsit.");
            }
            var grupuriHelper = new GrupuriDomeniiHelper(_context);
            var grupuriProgramStudii = await grupuriHelper.GetGrupuriProgramStudiiAsync();
            var domeniiDinDb = grupuriProgramStudii.SelectMany(g => g.Value).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            Dictionary<string, List<FormatiiStudii>> groupedFormatii = await _fondBurseService.GetGroupedFormatiiStudiiAsync();
            var programeDeStudii = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            Dictionary<string, List<string>> domenii = await _grupuriService.GetGrupuriAsync();
            Dictionary<string, List<string>> excludereStudenti = await _grupuriService.GetExcluderiStudentAsync();
            foreach (var file in pathStudentiList)
            {
                string programPrincipal = Path.GetFileNameWithoutExtension(file.FileName).ToUpper();
                programeDeStudii.Add(programPrincipal);
                using var stream = file.OpenReadStream();
                using var reader = ExcelReaderFactory.CreateReader(stream);
                do
                {
                    string sheetName = reader.Name?.Trim();
                    if (string.IsNullOrWhiteSpace(sheetName))
                        continue;
                    bool handledSpecial = false;
                    string grupa = await _grupuriHelper.GetGrupaAsync(programPrincipal);
                    List<string> domeniiByGrup;
                    if (domenii.TryGetValue(grupa, out var listaDomenii))
                    {
                        domeniiByGrup = listaDomenii;
                    }
                    else
                    {
                        domeniiByGrup = new List<string>();
                    }
                    string? domeniuPotrivit = domeniiByGrup.FirstOrDefault(d => d.Contains($"({sheetName})"));
                    if (!string.IsNullOrEmpty(domeniuPotrivit))
                    {
                        string domeniuCurat = new string(domeniuPotrivit.TakeWhile(c => c != '(').ToArray()).Trim();
                        programeDeStudii.Add(domeniuCurat);
                        handledSpecial = true;
                    }
                    if (!handledSpecial)
                    {
                        var match = Regex.Match(sheetName, @"^(\d*)([a-zA-Z]+)$", RegexOptions.IgnoreCase);
                        if (match.Success)
                        {
                            string raw = match.Groups[2].Value.ToUpper();
                            string subProgram = raw.EndsWith("DUAL")
                                ? $"{(raw[..^4].Length > 0 ? raw[..^4] : programPrincipal)}-DUAL"
                                : raw;
                            programeDeStudii.Add(subProgram);
                        }
                    }
                } while (reader.NextResult());
            }
            var domeniiLipsa = domeniiDinDb
                    .Where(db => !programeDeStudii.Contains(db, StringComparer.OrdinalIgnoreCase))
                    .ToList();
            if (domeniiLipsa.Any())
            {
                var msg = $"Nu au fost găsite toate domeniile de studii in listele cu studenti incarcate. Lipsesc: {string.Join(", ", domeniiLipsa)}.";
                _logger.LogFormatiiInfo(msg);
                return BadRequest("Eroare: unele domenii de studii lipsesc. Consultați logurile Formatii pentru detalii.");
            }
            var streamBurseFile = burseFile.OpenReadStream();
            StudentExcelReader excelReader = new StudentExcelReader();
            List<FondBurse> fonduri = await _fondBurseService.GetDateFromBursePerformanteAsync();
            var allStudentRecordsList = new List<Dictionary<string, List<StudentRecord>>>();
            bool toateCoincid = true;
            var discrepante = new List<string>();
            foreach (var pathStudenti in pathStudentiList)
            {
                using var stream = pathStudenti.OpenReadStream();
                var (studentRecords, excluderiPeDomeniu) = excelReader.ReadStudentRecordsFromExcel(stream, 
                    pathStudenti.FileName, domenii, excludereStudenti, _logger);
                var processed = new Dictionary<string, List<StudentRecord>>();
                foreach (var kvp in studentRecords)
                {
                    string processedKeyNormalized = AcronymGenerator.RemoveDiacritics(kvp.Key).ToUpperInvariant();
                    var matchedKey = groupedFormatii.Keys
                        .FirstOrDefault(k => AcronymGenerator.RemoveDiacritics(k).ToUpperInvariant() == processedKeyNormalized);
                    if (matchedKey == null)
                    {
                        _logger.LogFormatiiInfo($"Procesare date Studenti -> Domeniul '{kvp.Key} studentului' nu există în formatii studii.");
                        continue;
                    }
                    int processedCount = kvp.Value.Count;
                    int excludedCount = excluderiPeDomeniu.TryGetValue(kvp.Key, out int excl) ? excl : 0;
                    int groupedCount = groupedFormatii[matchedKey].Sum(f =>
                        ParseIntOrZero(f.FaraTaxaRomani) +
                        ParseIntOrZero(f.FaraTaxaRp) +
                        ParseIntOrZero(f.FaraTaxaUECEE) +
                        ParseIntOrZero(f.CuTaxaRomani) +
                        ParseIntOrZero(f.CuTaxaRM) +
                        ParseIntOrZero(f.CuTaxaUECEE) +
                        ParseIntOrZero(f.BursieriAIStatuluiRoman) +
                        ParseIntOrZero(f.CPV)
                    );
                    if (processedCount + excludedCount != groupedCount)
                    {
                        var msg = $"Numărul studenților pentru Domeniul '{kvp.Key}' în fișierul '{pathStudenti.FileName}'" +
                            $" nu coincide. Procesați: {processedCount}, Excluși: {excludedCount}, Așteptați: {groupedCount}";
                        _logger.LogStudentsExcels(msg);
                        discrepante.Add(msg);
                    }
                    var processedStudents = ProcessStudents(kvp.Value);
                    processed[kvp.Key] = processedStudents;
                }
                allStudentRecordsList.Add(processed);
            }
            if (discrepante.Any())
            {
                var msg = "A apărut o eroare. Consultați logurile din Students-Excels pentru detalii.";
                _logger.LogError(msg);
                throw new Exception(msg);
            }
            foreach (var studentRecords in allStudentRecordsList)
            {   
                var istoricList = new List<(string Emplid, BursaIstoric Istoric)>();
                foreach (var entry in studentRecords)
                {
                    string domeniu = entry.Key;
                    List<StudentRecord> students = ProcessStudents(entry.Value);
                    FondBurseMeritRepartizat? fondRepartizatByDomeniu = await _fondBurseMeritRepartizatService.GetByDomeniuAsync(domeniu);
                    students
                        .Where(s => s.Media == 0 || s.Media == null)
                        .ToList()
                        .ForEach(s => _logger.LogStudentInfo($"Studentul {s.Emplid} Program: {entry.Key} are media 0"));
                    if (fondRepartizatByDomeniu == null) continue;
                    (decimal valoareAnualBP1, decimal valoareAnualBP2, decimal valoareAnualBP1RP,decimal valoareAnualBP2RP) =
                        CalculateScholarshipValues(domeniu, fonduri, fondRepartizatByDomeniu, valoareRomaniDePretutindeni);
                    decimal sumaDisponibila = fondRepartizatByDomeniu.bursaAlocatata;
                    if (sumaDisponibila < 0)
                        continue;
                    decimal sumaRamasa;
                    var istoricePerDomeniu = new List<(string Emplid, BursaIstoric Istoric)>();
                    if (sumaDisponibila <= 31500)
                    {
                        (sumaRamasa, istoricePerDomeniu) = AssignScholarshipsOptimizedV2(
                            students,
                            sumaDisponibila,
                            valoareAnualBP1,
                            valoareAnualBP2,
                            valoareAnualBP1RP,
                            valoareAnualBP2RP,
                            epsilon,
                            fondRepartizatByDomeniu,
                            "0",
                            valoareRomaniDePretutindeni
                        );
                    }
                    else
                    {
                        (sumaRamasa, istoricePerDomeniu) = AssignScholarshipsOptimezedWithCriteriaMediilorAceleasi(
                            students,
                            sumaDisponibila,
                            valoareAnualBP1,
                            valoareAnualBP2,
                            valoareAnualBP1RP,
                            valoareAnualBP2RP,
                            epsilon,
                            fondRepartizatByDomeniu,
                            "0",
                            valoareRomaniDePretutindeni
                        );
                    }
                    sumaDisponibila = sumaRamasa;
                    istoricList.AddRange(istoricePerDomeniu);
                    students.ForEach(s => s.FondBurseMeritRepartizatId = fondRepartizatByDomeniu.ID);
                    var groupedByMedia = students.GroupBy(s => s.Media);
                    foreach (var group in groupedByMedia)
                    {
                        var valoriDistincteBursa = group
                            .Select(s => string.IsNullOrWhiteSpace(s.Bursa) ? null : s.Bursa.Trim())
                            .Distinct()
                            .ToList();
                        if (valoriDistincteBursa.Count > 1)
                        {
                            foreach (var student in group)
                            {
                                if (!string.IsNullOrWhiteSpace(student.Bursa))
                                {
                                    var altStudent = group.FirstOrDefault(s => s.Emplid != student.Emplid);

                                    if (altStudent != null)
                                    {
                                        student.TipInconsistenta = $"Etapa :0 ,Studentul Emplid: {student.Emplid} Nume: {student.NumeStudent}, Bursa: {(string.IsNullOrWhiteSpace(student.Bursa) ? "NU" : student.Bursa)}, Program: {fondRepartizatByDomeniu.domeniu}, cu media {group.Key} are aceeași medie ca studentul Emplid: {altStudent.Emplid}, Nume: {altStudent.NumeStudent}, Bursa: {(string.IsNullOrWhiteSpace(altStudent.Bursa) ? "NU" : altStudent.Bursa)}, Program: {fondRepartizatByDomeniu.domeniu}, dar nu a primit bursă.";
                                    }
                                }
                            }
                            var studentiCuAceeasiMedia = group.Select(s =>
                                $"Emplid: {s.Emplid}, Nume: {s.NumeStudent}, Bursa: {(string.IsNullOrWhiteSpace(s.Bursa) ? "NU" : s.Bursa)}, Program: {fondRepartizatByDomeniu.domeniu}"
                            );
                            var mesaj = $"⚠️ Atenție! etapa 0: Studenți cu media {group.Key} au situație aceeasi la bursă (valori diferite):\n" +
                                        string.Join("\n", studentiCuAceeasiMedia);
                            _logger.LogStudentInfo(mesaj);
                        }
                    }
                    var studentiCuIdCorect = await _fondBurseService.SaveNewStudentsAsync(students);
                    foreach (var (emplid, istoric) in istoricList)
                    {
                        var match = studentiCuIdCorect.FirstOrDefault(s => s.Emplid == emplid);
                        if (match != null)
                            istoric.StudentRecordId = match.Id;
                    }
                    fondRepartizatByDomeniu.SumaRamasa = sumaDisponibila;
                    await _fondBurseMeritRepartizatService.UpdateAsync(fondRepartizatByDomeniu);
                }
                try
                {
                    foreach (var item in istoricList)
                    {
                        var existing = await _context.BursaIstoric.FirstOrDefaultAsync(x =>
                            x.StudentRecordId == item.Istoric.StudentRecordId
                        );
                        if (existing != null)
                        {
                            existing.Motiv = item.Istoric.Motiv;
                            existing.Actiune = item.Istoric.Actiune;
                            existing.Suma = item.Istoric.Suma;
                            existing.Etapa = "0";
                            existing.Comentarii = item.Istoric.Comentarii;
                        }
                        else
                        {
                            _context.BursaIstoric.Add(item.Istoric);
                        }
                    }
                    await _context.SaveChangesAsync();

                }
                catch (Exception ex)
                {
                    _logger.LogError( "Eroare la salvarea în BursaIstoric "+ex.Message);
                }
            }
            List<StudentRecord> studentiCuBursa0 = await _fondBurseService.GetStudentsWithBursaFromDatabaseAsync();
            List<StudentScholarshipData> studentiClasificati0 = studentiCuBursa0
                .GroupBy(s => new { s.FondBurseMeritRepartizatId, s.FondBurseMeritRepartizat.domeniu })
                .Select(group => new StudentScholarshipData
                {
                    FondBurseId = group.Key.FondBurseMeritRepartizatId,
                    Domeniu = group.Key.domeniu,
                    BP1Count = group.Count(s => s.Bursa.ToLower().Contains("bp1")),
                    BP2Count = group.Count(s => s.Bursa.ToLower().Contains("bp2"))
                }).ToList();
             var studentiPeGrupa = await _fondBurseService.GetStudentiEligibiliPeGrupaAsync();
             var fonduriRepartizate = await _fondBurseMeritRepartizatService.GetAllAsync();
             var sumaDisponibilaPeGrupa = fonduriRepartizate
                 .GroupBy(f => f.Grupa)
                 .ToDictionary(
                     g => g.Key,
                     g => g.Sum(f => f.SumaRamasa)
                 );
             var fonduriDict = fonduriRepartizate.ToDictionary(f => f.ID, f => f);
             var fonduriPeDomeniu = fonduriRepartizate.ToDictionary(f => f.domeniu, f => f);
             foreach (var entry in studentiPeGrupa)
             {
                 string grupa = entry.Key;
                 List<StudentRecord> students = entry.Value;
                 var fonduriGrupa = fonduriRepartizate
                     .Where(f => f.Grupa == grupa)
                     .ToList();
                 if (!fonduriGrupa.Any()) continue;
                 var sumaRamasaPeFond = fonduriGrupa.ToDictionary(f => f.ID, f => f.SumaRamasa);
                 decimal sumaDisponibila = sumaDisponibilaPeGrupa[grupa];
                 if (sumaDisponibila < 0)
                     continue;
                 (decimal sumaNoua, var istoricBP2) = AssignOnlyBP2(students, sumaDisponibila, fonduri, sumaRamasaPeFond,"1", valoareRomaniDePretutindeni);
                 sumaDisponibila = sumaNoua;
                var groupedByMedia = students.GroupBy(s => s.Media);
                foreach (var group in groupedByMedia)
                {
                    var valoriDistincteBursa = group
                        .Select(s => string.IsNullOrWhiteSpace(s.Bursa) ? null : s.Bursa.Trim())
                        .Distinct()
                        .ToList();
                    if (valoriDistincteBursa.Count > 1)
                    {
                        foreach (var student in group)
                        {
                            if (!string.IsNullOrWhiteSpace(student.Bursa))
                            {
                                var altStudent = group.FirstOrDefault(s => s.Emplid != student.Emplid);

                                if (altStudent != null)
                                {
                                    student.TipInconsistenta = $"Etapa :1 ,Studentul Emplid: {student.Emplid} Nume: {student.NumeStudent}, Bursa: {(string.IsNullOrWhiteSpace(student.Bursa) ? "NU" : student.Bursa)}, Program: {student.FondBurseMeritRepartizat.domeniu}, cu media {group.Key} are aceeași medie ca studentul Emplid: {altStudent.Emplid}, Nume: {altStudent.NumeStudent}, Bursa: {(string.IsNullOrWhiteSpace(altStudent.Bursa) ? "NU" : altStudent.Bursa)}, Program: {altStudent.FondBurseMeritRepartizat.domeniu}, dar nu a primit bursă.";
                                }
                            }
                        }
                        var studentiCuAceeasiMedia = group.Select(s =>
                            $"Emplid: {s.Emplid}, Nume: {s.NumeStudent}, Bursa: {(string.IsNullOrWhiteSpace(s.Bursa) ? "NU" : s.Bursa)}, Program: {s.FondBurseMeritRepartizat.domeniu}"
                        );
                        var mesaj = $"⚠️ Atenție! etapa1: Studenți cu media {group.Key} au situație mixtă la bursă (valori diferite):\n" +
                                    string.Join("\n", studentiCuAceeasiMedia);
                        _logger.LogStudentInfo(mesaj);
                    }
                }
                await _fondBurseService.SaveNewStudentsAsync(students);
                 foreach (var fond in fonduriGrupa)
                 {
                     fond.SumaRamasa = sumaRamasaPeFond[fond.ID];
                     await _fondBurseMeritRepartizatService.UpdateAsync(fond);
                 }
                 foreach (var hist in istoricBP2)
                 {
                     var existing = await _context.BursaIstoric.FirstOrDefaultAsync(x =>
                         x.StudentRecordId == hist.Istoric.StudentRecordId 
                     );
                     if (existing != null)
                     {
                         existing.Motiv = hist.Istoric.Motiv;
                         existing.Actiune = hist.Istoric.Actiune;
                         existing.Suma = hist.Istoric.Suma;
                         existing.Etapa = "1";
                         existing.Comentarii = hist.Istoric.Comentarii;
                     }
                     else
                     {
                         await _context.BursaIstoric.AddAsync(hist.Istoric);
                     }
                 }
                 await _context.SaveChangesAsync();
             }
            List<StudentRecord> studentiCuBursa1 = await _fondBurseService.GetStudentsWithBursaFromDatabaseAsync();
            List<StudentScholarshipData> studentiClasificati1 = studentiCuBursa1
                .GroupBy(s => new { s.FondBurseMeritRepartizatId, s.FondBurseMeritRepartizat.domeniu })
                .Select(group => new StudentScholarshipData
                {
                    FondBurseId = group.Key.FondBurseMeritRepartizatId,
                    Domeniu = group.Key.domeniu,
                    BP1Count = group.Count(s => s.Bursa.ToLower().Contains("bp1")),
                    BP2Count = group.Count(s => s.Bursa.ToLower().Contains("bp2"))
                }).ToList();
            var studentiPeGrup = await _fondBurseService.GetStudentiEligibiliPeGrupProgramStudiiAsync();
            var fonduriPeGrup = studentiPeGrup.Keys
                .ToDictionary(
                    grup => grup,
                    grup =>
                    {
                        var domeniiGrup = studentiPeGrup[grup]
                            .Select(s => s.FondBurseMeritRepartizat.domeniu.Split(' ')[0].Trim())
                            .Distinct()
                            .ToList();

                        return fonduriRepartizate
                            .Where(f => domeniiGrup.Contains(f.domeniu.Split(' ')[0].Trim()))
                            .ToList();
                    });
            var sumaDisponibilaPeGrup = fonduriPeGrup
                .ToDictionary(
                    g => g.Key,
                    g => g.Value.Sum(f => f.SumaRamasa)
                );
            var studentiLicentaPeGrup = studentiPeGrup
                .ToDictionary(
                    g => g.Key,
                    g => g.Value
                        .Where(s => !s.FondBurseMeritRepartizat.programStudiu
                            .ToLowerInvariant()
                            .Contains("master"))
                        .ToList()
                );
            foreach (var entry in studentiLicentaPeGrup)
            {
                var grup = entry.Key;
                var students = entry.Value;
                if (!fonduriPeGrup.ContainsKey(grup)) continue;
                var fonduriGrupa = fonduriPeGrup[grup];
                var sumaRamasaPeFond = fonduriGrupa.ToDictionary(f => f.ID, f => f.SumaRamasa);
                decimal sumaDisponibila = sumaDisponibilaPeGrup[grup];
                if (sumaDisponibila <= 0) continue;
                (decimal sumaNoua, var istoricBP2) = AssignOnlyBP2(students, sumaDisponibila, fonduri, sumaRamasaPeFond, "2", valoareRomaniDePretutindeni);
                sumaDisponibila = sumaNoua;
                var groupedByMedia = students.GroupBy(s => s.Media);
                foreach (var group in groupedByMedia)
                {
                    var valoriDistincteBursa = group
                        .Select(s => string.IsNullOrWhiteSpace(s.Bursa) ? null : s.Bursa.Trim())
                        .Distinct()
                        .ToList();
                    if (valoriDistincteBursa.Count > 1)
                    {
                        foreach (var student in group)
                        {
                            if (!string.IsNullOrWhiteSpace(student.Bursa))
                            {
                                var altStudent = group.FirstOrDefault(s => s.Emplid != student.Emplid);
                                if (altStudent != null)
                                {
                                    student.TipInconsistenta = $"Etapa :2 ,Studentul Emplid: {student.Emplid} Nume: {student.NumeStudent}, Bursa: {(string.IsNullOrWhiteSpace(student.Bursa) ? "NU" : student.Bursa)}, Program: {student.FondBurseMeritRepartizat.domeniu}, cu media {group.Key} are aceeași medie ca studentul Emplid: {altStudent.Emplid}, Nume: {altStudent.NumeStudent}, Bursa: {(string.IsNullOrWhiteSpace(altStudent.Bursa) ? "NU" : altStudent.Bursa)}, Program: {altStudent.FondBurseMeritRepartizat.domeniu}, dar nu a primit bursă.";
                                }
                            }
                        }
                        var studentiCuAceeasiMedia = group.Select(s =>
                            $"Emplid: {s.Emplid}, Nume: {s.NumeStudent}, Bursa: {(string.IsNullOrWhiteSpace(s.Bursa) ? "NU" : s.Bursa)}, Program: {s.FondBurseMeritRepartizat.domeniu}"
                        );
                        var mesaj = $"⚠️ Atenție!etapa2: Studenți cu media {group.Key} au situație mixtă la bursă (valori diferite):\n" +
                                    string.Join("\n", studentiCuAceeasiMedia);
                        _logger.LogStudentInfo(mesaj);
                    }
                }
                await _fondBurseService.SaveNewStudentsAsync(students);
                foreach (var fond in fonduriGrupa)
                {
                    fond.SumaRamasa = sumaRamasaPeFond[fond.ID];
                    await _fondBurseMeritRepartizatService.UpdateAsync(fond);
                }
                foreach (var hist in istoricBP2)
                {
                    var existing = await _context.BursaIstoric.FirstOrDefaultAsync(x =>
                        x.StudentRecordId == hist.Istoric.StudentRecordId &&
                        x.TipBursa == hist.Istoric.TipBursa &&
                        x.DataModificare == hist.Istoric.DataModificare
                    );
                    if (existing != null)
                    {
                        existing.Motiv = hist.Istoric.Motiv;
                        existing.Actiune = hist.Istoric.Actiune;
                        existing.Suma = hist.Istoric.Suma;
                        existing.Etapa = "2";
                        existing.Comentarii = hist.Istoric.Comentarii;
                    }
                    else
                    {
                        await _context.BursaIstoric.AddAsync(hist.Istoric);
                    }
                }
                await _context.SaveChangesAsync();
            }
            List<StudentRecord> studentiCuBursa2 = await _fondBurseService.GetStudentsWithBursaFromDatabaseAsync();
            List<StudentScholarshipData> studentiClasificati2 = studentiCuBursa2
                .GroupBy(s => new { s.FondBurseMeritRepartizatId, s.FondBurseMeritRepartizat.domeniu })
                .Select(group => new StudentScholarshipData
                {
                    FondBurseId = group.Key.FondBurseMeritRepartizatId,
                    Domeniu = group.Key.domeniu,
                    BP1Count = group.Count(s => s.Bursa.ToLower().Contains("bp1")),
                    BP2Count = group.Count(s => s.Bursa.ToLower().Contains("bp2"))
                }).ToList();
            var grupuriBurse = await _grupuriHelper.GetGrupuriBurseAsync();
            foreach (var grup in grupuriBurse)
            {
                string numeGrup = grup.Key;
                List<string> domeniiGrup = grup.Value;
                var fonduriInGrup = fonduriRepartizate
                    .Where(f => GetDomeniiDinGrupa(f.Grupa)
                        .Any(domeniu => domeniiGrup.Contains(domeniu)))
                    .ToList();
                if (!fonduriInGrup.Any()) continue;
                var fonduriPeProgram = fonduriInGrup
                        .Where(f => f.programStudiu?.ToLower() == "licenta")
                        .GroupBy(f => f.Grupa) // Grupez după Grupa (AIA(1), AIA(2), C(1), etc.)
                        .Select(grupProgram => new
                        {
                            ProgramStudiu = grupProgram.Key, // AIA(1), AIA(2), C(1), etc.
                            SumaRamasa = grupProgram.Sum(f => f.SumaRamasa),
                            SumaInitiala = grupProgram.Sum(f => f.bursaAlocatata),
                            Fonduri = grupProgram.ToList()
                        })
                        .Where(g => g.SumaInitiala > 0)
                        .Select(g => new
                        {
                            g.ProgramStudiu,
                            g.Fonduri,
                            Fractiune = g.SumaRamasa
                        })
                        .OrderByDescending(g => g.Fractiune)
                        .FirstOrDefault();
                if (fonduriPeProgram == null || fonduriPeProgram.Fonduri.Sum(f => f.SumaRamasa) <= 0)
                    continue;
                decimal sumaDisponibila = fonduriInGrup.Sum(f => f.SumaRamasa);
                if (sumaDisponibila <= 0) continue;
                var sumaRamasaPeFond = fonduriPeProgram.Fonduri.ToDictionary(f => f.ID, f => f.SumaRamasa);
                var studentiEligibili = fonduriPeProgram.Fonduri
                    .SelectMany(f => f.Studenti)
                    .Where(s => string.IsNullOrWhiteSpace(s.Bursa) || s.Bursa.Trim().ToLower() == "nicio bursă")
                    .OrderByDescending(s => s.Media)
                    .ToList();
                (decimal sumaNoua, var istoricBP2) = AssignOnlyBP2(
                    studentiEligibili,
                    sumaDisponibila,
                    fonduri,
                    sumaRamasaPeFond,
                    "3", valoareRomaniDePretutindeni
                );
                var groupedByMedia = studentiEligibili.GroupBy(s => s.Media);
                foreach (var group in groupedByMedia)
                {
                    var valoriDistincteBursa = group
                        .Select(s => string.IsNullOrWhiteSpace(s.Bursa) ? null : s.Bursa.Trim())
                        .Distinct()
                        .ToList();
                    if (valoriDistincteBursa.Count > 1)
                    {
                        foreach (var student in group)
                        {
                            if (!string.IsNullOrWhiteSpace(student.Bursa))
                            {
                                var altStudent = group.FirstOrDefault(s => s.Emplid != student.Emplid);

                                if (altStudent != null)
                                {
                                    student.TipInconsistenta = $"Etapa :3 ,Studentul Emplid: {student.Emplid} Nume: {student.NumeStudent}, Bursa: {(string.IsNullOrWhiteSpace(student.Bursa) ? "NU" : student.Bursa)}, Program: {student.FondBurseMeritRepartizat.domeniu}, cu media {group.Key} are aceeași medie ca studentul Emplid: {altStudent.Emplid}, Nume: {altStudent.NumeStudent}, Bursa: {(string.IsNullOrWhiteSpace(altStudent.Bursa) ? "NU" : altStudent.Bursa)}, Program: {altStudent.FondBurseMeritRepartizat.domeniu}, dar nu a primit bursă.";
                                }
                            }
                        }
                        var studentiCuAceeasiMedia = group.Select(s =>
                            $"Emplid: {s.Emplid}, Nume: {s.NumeStudent}, Bursa: {(string.IsNullOrWhiteSpace(s.Bursa) ? "NU" : s.Bursa)}, Program: {s.FondBurseMeritRepartizat.domeniu}"
                        );
                        var mesaj = $"⚠️ Atenție! etapa3: Studenți cu media {group.Key} au situație mixtă la bursă (valori diferite):\n" +
                                    string.Join("\n", studentiCuAceeasiMedia);
                        _logger.LogStudentInfo(mesaj);
                    }
                }
                await _fondBurseService.SaveNewStudentsAsync(studentiEligibili);
                foreach (var fond in fonduriPeProgram.Fonduri)
                {
                    fond.SumaRamasa = sumaRamasaPeFond[fond.ID];
                    await _fondBurseMeritRepartizatService.UpdateAsync(fond);
                }
                foreach (var entry in istoricBP2)
                {
                    var existing = await _context.BursaIstoric.FirstOrDefaultAsync(x =>
                        x.StudentRecordId == entry.Istoric.StudentRecordId);
                    if (existing != null)
                    {
                        existing.Motiv = entry.Istoric.Motiv;
                        existing.Actiune = entry.Istoric.Actiune;
                        existing.Suma = entry.Istoric.Suma;
                        existing.Etapa = "3";
                        existing.Comentarii = entry.Istoric.Comentarii;
                    }
                    else
                    {
                        await _context.BursaIstoric.AddAsync(entry.Istoric);
                    }
                }
                await _context.SaveChangesAsync();
            }
            List<StudentRecord> studentiCuBursa3 = await _fondBurseService.GetStudentsWithBursaFromDatabaseAsync();
            List<StudentScholarshipData> studentiClasificati3 = studentiCuBursa3
                .GroupBy(s => new { s.FondBurseMeritRepartizatId, s.FondBurseMeritRepartizat.domeniu })
                .Select(group => new StudentScholarshipData
                {
                    FondBurseId = group.Key.FondBurseMeritRepartizatId,
                    Domeniu = group.Key.domeniu,
                    BP1Count = group.Count(s => s.Bursa.ToLower().Contains("bp1")),
                    BP2Count = group.Count(s => s.Bursa.ToLower().Contains("bp2"))
                }).ToList();
            foreach (var item in studentiClasificati3)
            {
                Console.WriteLine($"Domeniu: {item.Domeniu}, BP1: {item.BP1Count}, BP2: {item.BP2Count}");
            }
            var grupuriCuSumaRamasa = grupuriBurse
                .Select(grup =>
                {
                    var fonduriGrup = fonduriRepartizate
                        .Where(f =>
                            GetDomeniiDinGrupa(f.Grupa)
                                .Any(d => grup.Value.Contains(d)))
                        .ToList();
                    var sumaRamasa = fonduriGrup.Sum(f => f.SumaRamasa);
                    var sumaInitiala = fonduriGrup.Sum(f => f.bursaAlocatata);
                    decimal fractiune = sumaRamasa; 
                    return new
                    {
                        NumeGrup = grup.Key,
                        Domenii = grup.Value,
                        SumaRamasa = sumaRamasa,
                        Fractiune = fractiune,
                        Fonduri = fonduriGrup
                    };
                })
                .OrderByDescending(g => g.Fractiune)
                .ToList();
            var grupCuSumaMaxima = grupuriCuSumaRamasa.FirstOrDefault();
            if (grupCuSumaMaxima != null && grupCuSumaMaxima.SumaRamasa > 0)
            {
                var studentiGrup = await _fondBurseService.GetStudentiEligibiliPeDomeniiAsync(grupCuSumaMaxima.Domenii);
                if (studentiGrup.Any())
                {
                    var fonduriGrup = fonduriRepartizate
                        .Where(f =>
                            GetDomeniiDinGrupa(f.Grupa)
                                .Any(d => grupCuSumaMaxima.Domenii.Contains(d)))
                        .ToList();
                    var sumaRamasaPeFond = fonduriGrup.ToDictionary(f => f.ID, f => f.SumaRamasa);
                    decimal sumaDisponibila = fonduriRepartizate.Sum(f => f.SumaRamasa);
                    studentiGrup = studentiGrup
                        .OrderByDescending(s => s.Media)
                        .ToList();
                    (decimal sumaNoua, var istoricBP2) = AssignOnlyBP2(studentiGrup, sumaDisponibila, fonduri, sumaRamasaPeFond, "4", valoareRomaniDePretutindeni);
                    sumaDisponibila = sumaNoua;
                    var groupedByMedia = studentiGrup.GroupBy(s => s.Media);
                    foreach (var group in groupedByMedia)
                    {
                        var valoriDistincteBursa = group
                            .Select(s => string.IsNullOrWhiteSpace(s.Bursa) ? null : s.Bursa.Trim())
                            .Distinct()
                            .ToList();
                        if (valoriDistincteBursa.Count > 1)
                        {
                            foreach (var student in group)
                            {
                                if (!string.IsNullOrWhiteSpace(student.Bursa))
                                {
                                    var altStudent = group.FirstOrDefault(s => s.Emplid != student.Emplid);

                                    if (altStudent != null)
                                    {
                                        student.TipInconsistenta = $" Etapa :4 :Studentul Emplid: {student.Emplid} Nume: {student.NumeStudent}, Bursa: {(string.IsNullOrWhiteSpace(student.Bursa) ? "NU" : student.Bursa)}, Program: {student.FondBurseMeritRepartizat.domeniu}, cu media {group.Key} are aceeași medie ca studentul Emplid: {altStudent.Emplid}, Nume: {altStudent.NumeStudent}, Bursa: {(string.IsNullOrWhiteSpace(altStudent.Bursa) ? "NU" : altStudent.Bursa)}, Program: {altStudent.FondBurseMeritRepartizat.domeniu}, dar nu a primit bursă.";
                                    }
                                }
                            }
                            var studentiCuAceeasiMedia = group.Select(s =>
                                $"Emplid: {s.Emplid}, Nume: {s.NumeStudent}, Bursa: {(string.IsNullOrWhiteSpace(s.Bursa) ? "NU" : s.Bursa)}, Program: {s.FondBurseMeritRepartizat.domeniu}"
                            );
                            var mesaj = $"⚠️ Atenție! etapa 4: Studenți cu media {group.Key} au situație mixtă la bursă (valori diferite):\n" +
                                        string.Join("\n", studentiCuAceeasiMedia);
                            _logger.LogStudentInfo(mesaj);
                        }
                    }
                    await _fondBurseService.SaveNewStudentsAsync(studentiGrup);
                    foreach (var fond in fonduriGrup)
                    {
                        fond.SumaRamasa = sumaRamasaPeFond[fond.ID];
                        await _fondBurseMeritRepartizatService.UpdateAsync(fond);
                    }
                    foreach (var entry in istoricBP2)
                    {
                        var existing = await _context.BursaIstoric.FirstOrDefaultAsync(x =>
                            x.StudentRecordId == entry.Istoric.StudentRecordId 
                        );
                        if (existing != null)
                        {
                            existing.Motiv = entry.Istoric.Motiv;
                            existing.Actiune = entry.Istoric.Actiune;
                            existing.Suma = entry.Istoric.Suma;
                            existing.Etapa = "4";
                            existing.Comentarii = entry.Istoric.Comentarii;
                        }
                        else
                        {
                            await _context.BursaIstoric.AddAsync(entry.Istoric);
                        }
                    }
                    await _context.SaveChangesAsync();
                    Console.WriteLine("✅ Redistribuire finală aplicată cu succes.");
                }
                else
                {
                    Console.WriteLine("⚠️ Nu există studenți eligibili în grupul selectat.");
                }
            }
            else
            {
                Console.WriteLine("⚠️ Nu există fonduri rămase suficiente pentru redistribuire (punctul 4).");
            }
            List<StudentRecord> studentiCuBursa4 = await _fondBurseService.GetStudentsWithBursaFromDatabaseAsync();
            bool EsteRP(StudentRecord s) => !string.IsNullOrWhiteSpace(s.TaraCetatenie) && !string.Equals(s.TaraCetatenie, "ROU", StringComparison.OrdinalIgnoreCase);
            var studentiClasificati4 = studentiCuBursa4
                .GroupBy(s => new { s.FondBurseMeritRepartizatId, s.FondBurseMeritRepartizat.domeniu })
                .Select(group =>
                {
                    int bp1Count, bp1CountRP, bp2Count, bp2CountRP;

                    if (valoareRomaniDePretutindeni.HasValue)
                    {
                        bp1Count = group.Count(s => s.Bursa.ToLower().Contains("bp1") && !EsteRP(s));
                        bp1CountRP = group.Count(s => s.Bursa.ToLower().Contains("bp1") && EsteRP(s));
                        bp2Count = group.Count(s => s.Bursa.ToLower().Contains("bp2") && !EsteRP(s));
                        bp2CountRP = group.Count(s => s.Bursa.ToLower().Contains("bp2") && EsteRP(s));
                    }
                    else
                    {
                        bp1Count = group.Count(s => s.Bursa.ToLower().Contains("bp1"));
                        bp1CountRP = 0;
                        bp2Count = group.Count(s => s.Bursa.ToLower().Contains("bp2"));
                        bp2CountRP = 0;
                    }

                    return new StudentScholarshipData
                    {
                        FondBurseId = group.Key.FondBurseMeritRepartizatId,
                        Domeniu = group.Key.domeniu,
                        BP1Count = bp1Count,
                        BP1CountRP = bp1CountRP,
                        BP2Count = bp2Count,
                        BP2CountRP = bp2CountRP,
                        TipInconsistenta = string.Join(", ", group
                            .Select(s => s.TipInconsistenta)
                            .Where(t => !string.IsNullOrWhiteSpace(t))
                            .Distinct())
                    };
                })
                .ToList();
            string licentaFolder = Path.Combine(Environment.CurrentDirectory, "Licenta");
            if (!Directory.Exists(licentaFolder))
            {
                Directory.CreateDirectory(licentaFolder);
            }
            string etapa0Path = Path.Combine(licentaFolder, $"Etapa_0.xlsx");
        using (var fileStream = new FileStream(etapa0Path, FileMode.Create, FileAccess.Write))
            {
                using var initialStream = burseFile.OpenReadStream();
                await initialStream.CopyToAsync(fileStream);
            }
            List<List<StudentScholarshipData>> toateEtapele = new()
            {
                studentiClasificati0,
                studentiClasificati1,
                studentiClasificati2,
                studentiClasificati3,
                studentiClasificati4
            };
            string previousPath = etapa0Path;
            for (int i = 0; i < toateEtapele.Count; i++)
            {
                string etapaInputPath = previousPath;
                string etapaOutputPath = Path.Combine(licentaFolder, $"Etapa_{i + 1}.xlsx");
                using var input = new FileStream(etapaInputPath, FileMode.Open, FileAccess.Read);
                using var output = new FileStream(etapaOutputPath, FileMode.Create, FileAccess.Write);
                var updatedStream = ExcelUpdater.UpdateScholarshipCounts(input, toateEtapele[i]);
                updatedStream.Position = 0;
                await updatedStream.CopyToAsync(output);
                previousPath = etapaOutputPath; // pentru următoarea rundă
            }
            string finalFilePath = previousPath;
            var finalBytes = await System.IO.File.ReadAllBytesAsync(finalFilePath);
            return File(
                finalBytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                Path.GetFileName(finalFilePath)
            );

        }
        public static List<string> GetDomeniiDinGrupa(string grupa)
        {
            if (string.IsNullOrEmpty(grupa))
                return new List<string>();
            return grupa.Split('/')
                .Select(d => d.Trim().ToUpper()) // sau păstrezi lowercase, după cum ai în GrupuriBurse
                .ToList();
        }
        private List<StudentRecord> ProcessStudents(List<StudentRecord> students)
        {
            return EliminaStudentiNeeligibili(students).OrderByDescending(s => s.Media).ToList();
        }
        private (decimal, decimal, decimal, decimal) CalculateScholarshipValues(
    string domeniu,
    List<FondBurse> fonduri,
    FondBurseMeritRepartizat fondRepartizat,
    double? valoareRomaniDePretutindeni = null)
        {
            decimal valoareBP1, valoareBP2, valoareBP1RP, valoareBP2RP;

            if (domeniu.Contains("4") || (fondRepartizat.programStudiu == "master" && domeniu.Contains("2")))
            {
                valoareBP1 = fonduri[0].ValoreaLunara * 9.35M;
                valoareBP2 = fonduri[1].ValoreaLunara * 9.35M;

                if (valoareRomaniDePretutindeni != null)
                {
                    decimal valoareRP = (decimal)valoareRomaniDePretutindeni;
                    valoareBP1RP = valoareBP1 - (valoareRP * 9.35M);
                    valoareBP2RP = valoareBP2 - (valoareRP * 9.35M);
                }
                else
                {
                    valoareBP1RP = valoareBP1;
                    valoareBP2RP = valoareBP2;
                }
            }
            else
            {
                valoareBP1 = fonduri[0].ValoreaLunara * 12;
                valoareBP2 = fonduri[1].ValoreaLunara * 12;

                if (valoareRomaniDePretutindeni != null)
                {
                    decimal valoareRP = (decimal)valoareRomaniDePretutindeni;
                    valoareBP1RP = valoareBP1 - (valoareRP * 12);
                    valoareBP2RP = valoareBP2 - (valoareRP * 12);
                }
                else
                {
                    valoareBP1RP = valoareBP1;
                    valoareBP2RP = valoareBP2;
                }
            }

            return (valoareBP1, valoareBP2, valoareBP1RP, valoareBP2RP);
        }
       
        private (decimal, List<(string Emplid, BursaIstoric Istoric)>) AssignScholarshipsOptimezedWithCriteriaMediilorAceleasi(
    List<StudentRecord> students,
    decimal sumaDisponibila,
    decimal valoareAnualBP1,
    decimal valoareAnualBP2,
    decimal valoareAnualBP1RP,
    decimal valoareAnualBP2RP,
    decimal epsilon,
    FondBurseMeritRepartizat fondBurseMeritRepartizat,
    string etapa,
    double? valoareRomaniDePretutindeni = null)
        {
            var istoricList = new List<(string Emplid, BursaIstoric Istoric)>();

            // --- NEW: Sort students based on Media and tie-breaking criteria ---
            // This is the crucial change to apply your tie-breaking rules.
            students.Sort(new StudentScholarshipComparer(fondBurseMeritRepartizat));
            // --- END NEW ---

            // The rest of your existing logic
            decimal? primaMedie = students.FirstOrDefault()?.Media; // Note: primaMedie will now be the media of the first student *after* sorting.
            bool aFostAcordatBP2 = false;
            StudentRecord studentAnterior = null;

            foreach (var student in students)
            {
                //decimal diferenta = studentAnterior != null ? Math.Abs(studentAnterior.Media - student.Media) : 0;

                decimal diferenta = primaMedie.HasValue ? Math.Abs(primaMedie.Value - student.Media) : 0;
                // The 'diferenta' calculation will now reflect the media of the previously processed student
                // in the *sorted* list. If two students had the same primary Media and were ordered by tie-breakers,
                // their 'diferenta' will be 0, correctly triggering the logic for close averages.
                bool esteRP = !string.Equals(student.TaraCetatenie, "ROU", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(student.TaraCetatenie);
                string bursaAtribuita = null;
                decimal suma = 0;
                string motiv = "";
                string explicatie = "";
                string fallback = "";

                if (sumaDisponibila <= 0)
                {
                    student.Bursa = null;
                    // No need to continue if no funds, but consider what "continue" does.
                    // If you want to log why no scholarship is given, do it here.
                    continue;
                }

                bool eligibilPentruBP1 = ("licenta".Equals(fondBurseMeritRepartizat.programStudiu) && student.Media >= 9.00M)
                                       || ("master".Equals(fondBurseMeritRepartizat.programStudiu) && student.Media >= 9.5M);
                decimal valoareBP1Actuala = esteRP ? valoareAnualBP1RP : valoareAnualBP1;
                decimal valoareBP2Actuala = esteRP ? valoareAnualBP2RP : valoareAnualBP2;
                if (eligibilPentruBP1)
                {
                    if (!aFostAcordatBP2 && diferenta <= epsilon)
                    {
                        if (sumaDisponibila >= valoareBP1Actuala)
                        {
                            bursaAtribuita = "BP1";
                            suma = valoareBP1Actuala;
                            motiv = "Media ≥ 9.00 și Δ ≤ ε – BP1 acordat";
                        }
                        else if (sumaDisponibila >= valoareBP2Actuala)
                        {
                            bursaAtribuita = "BP2";
                            suma = valoareBP2Actuala;
                            motiv = "Fond insuficient pentru BP1 – fallback la BP2";
                            fallback = $"(necesar BP1: {valoareAnualBP1:F2} lei, dar disponibil doar: {sumaDisponibila:F2} lei)";
                            aFostAcordatBP2 = true;
                        }
                    }
                    else if (sumaDisponibila >= valoareAnualBP2)
                    {
                        bursaAtribuita = "BP2";
                        suma = valoareBP2Actuala;
                        motiv = "Δ > ε – fallback la BP2";
                        fallback = $"(Δ = {diferenta:F2} > ε = {epsilon:F2})";
                        aFostAcordatBP2 = true;
                    }
                }
                else if (sumaDisponibila >= valoareAnualBP2 && student.Media >= 8.00M)
                {
                    bursaAtribuita = "BP2";
                    suma = valoareBP2Actuala;
                    motiv = "Media < 9.00 – BP2 acordat";
                    fallback = "(criteriu media)";
                    aFostAcordatBP2 = true;
                }

                // Logic for explicatie and Comentarii AI as you have it
                if (primaMedie == null) // This will now apply to the first student in the *sorted* list
                {
                    explicatie = "Primul student – fără comparație anterioară";
                }
                else
                {
                    explicatie = $"Media primului student: {primaMedie:F2} → Δ = {diferenta:F2} {(diferenta <= epsilon ? "(Δ ≤ ε)" : "(Δ > ε)")}";
                }

                if (!string.IsNullOrEmpty(bursaAtribuita))
                {
                    student.Bursa = bursaAtribuita;
                    student.SumaBursa = suma;
                    sumaDisponibila -= suma;

                    string anterior = studentAnterior != null
                        ? $"Studentul anterior: {studentAnterior.NumeStudent} (media {studentAnterior.Media:F2}, bursă {studentAnterior.Bursa})"
                        : "Acesta este primul student care primește bursă.";

                    // Filter for students who haven't received a scholarship yet, after the current one
                    var urmatorii = students
                        .Where(s => string.IsNullOrWhiteSpace(s.Bursa) && s != student) // Ensure 's != student' is used for the current iteration
                        .Take(5)
                        .Select(s => $"(Emplid: {s.Emplid}, Media: {s.Media:F2}, An: {s.An+1}, Bursa: {s.Bursa ?? "—"})")
                        .ToList();

                    string urmatoriiText = urmatorii.Count > 0
                        ? $"Următorii studenți eligibili: {string.Join(", ", urmatorii)}"
                        : "Nu mai sunt studenți eligibili în acest moment.";

                    string comentariuAI = $"Studentul {student.NumeStudent} cu media {student.Media:F2} a primit bursa de tip {bursaAtribuita} pentru că {motiv.ToLower()}. " +
                                          $"{(string.IsNullOrEmpty(fallback) ? "" : fallback + " ")}{anterior}. {urmatoriiText}. " +
                                          $"Fonduri rămase: {sumaDisponibila:F2} lei.";

                    string comentariu = $"Etapa: {etapa} | Media: {student.Media:F2} | {motiv} {fallback} | " +
                                        $"{explicatie} | Suma acordată: {suma:F2} lei | Rămas fond: {sumaDisponibila:F2} lei";

                    istoricList.Add((student.Emplid, new BursaIstoric
                    {
                        StudentRecordId = student.Id,
                        TipBursa = bursaAtribuita,
                        Actiune = "Acordare",
                        Suma = suma,
                        Motiv = motiv,
                        Comentarii = comentariu,
                        ComentariiAI = comentariuAI,
                        DataModificare = DateTime.Now,
                        Etapa = etapa
                    }));

                    studentAnterior = student;
                }
                else
                {
                    student.Bursa = null;
                }
            }

            return (sumaDisponibila, istoricList);
        }
        private (decimal, List<(string Emplid, BursaIstoric Istoric)>) AssignScholarshipsOptimizedV2(
    List<StudentRecord> students,
    decimal sumaDisponibila,
    decimal valoareAnualBP1,
    decimal valoareAnualBP2,
    decimal valoareAnualBP1RP,
    decimal valoareAnualBP2RP,
    decimal epsilon,
    FondBurseMeritRepartizat fondBurseMeritRepartizat,
    string etapa,
    double? valoareRomaniDePretutindeni = null)
        {
            var istoricList = new List<(string Emplid, BursaIstoric Istoric)>();
            decimal sumaDisponibilaInitiala = sumaDisponibila;
            var sortedStudents = students.OrderByDescending(s => s.Media).ToList();
            var eligibleBP1 = new List<StudentRecord>(); 
            var eligibleOnlyBP2 = new List<StudentRecord>();
            foreach (var student in sortedStudents)
            {
                bool isEligibleForBP1Criterion = ("licenta".Equals(fondBurseMeritRepartizat.programStudiu) && student.Media >= 9.00M) ||
                                                 ("master".Equals(fondBurseMeritRepartizat.programStudiu) && student.Media >= 9.5M);
                bool isEligibleForBP2Criterion = student.Media >= 8.00M;
                if (isEligibleForBP1Criterion)
                {
                    eligibleBP1.Add(student);
                }
                else if (isEligibleForBP2Criterion)
                {
                    eligibleOnlyBP2.Add(student);
                }
            }
            int maxBP1Possible = eligibleBP1.Count;
            int maxBP2Possible = eligibleOnlyBP2.Count;
            int bestNumBP1 = 0;
            int bestNumBP2 = 0;
            int maxTotalScholarships = -1;
            for (int numBP1 = maxBP1Possible; numBP1 >= 0; numBP1--)
            {
                if (numBP1 > eligibleBP1.Count) continue;
                decimal costBP1 = numBP1 * valoareAnualBP1;
                decimal remainingFundsAfterBP1 = sumaDisponibila - costBP1;
                if (remainingFundsAfterBP1 < 0) continue;
                int numBP2 = 0;
                decimal tempRemaining = remainingFundsAfterBP1;
                var studentsForBP2 = eligibleBP1.Skip(numBP1).Concat(eligibleOnlyBP2)
                                               .OrderByDescending(s => s.Media);
                foreach (var student in studentsForBP2)
                {
                    if (tempRemaining >= valoareAnualBP2)
                    {
                        numBP2++;
                        tempRemaining -= valoareAnualBP2;
                    }
                    else
                    {
                        break;
                    }
                }
                int currentTotalScholarships = numBP1 + numBP2;
                if (currentTotalScholarships > maxTotalScholarships)
                {
                    maxTotalScholarships = currentTotalScholarships;
                    bestNumBP1 = numBP1;
                    bestNumBP2 = numBP2;
                }
                else if (currentTotalScholarships == maxTotalScholarships)
                {
                    if (numBP1 > bestNumBP1)
                    {
                        bestNumBP1 = numBP1;
                        bestNumBP2 = numBP2;
                    }
                }
            }
            sumaDisponibila = sumaDisponibilaInitiala;
            int bp1AllocatedCount = 0;
            foreach (var student in eligibleBP1.OrderByDescending(s => s.Media))
            {
                bool esteRP = !string.Equals(student.TaraCetatenie, "ROU", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(student.TaraCetatenie);

                decimal valoareBP1Actuala = esteRP ? valoareAnualBP1RP : valoareAnualBP1;

                if (bp1AllocatedCount < bestNumBP1 && sumaDisponibila >= valoareAnualBP1)
                {
                    student.Bursa = "BP1";
                    student.SumaBursa = valoareBP1Actuala;
                    sumaDisponibila -= valoareBP1Actuala;
                    bp1AllocatedCount++;

                    string motiv = "Media eligibilă pentru BP1.";
                    string explicatie = $"Media: {student.Media:F2}";
                    string comentariu = $"Etapa: {etapa} | Media: {student.Media:F2} | {motiv} | {explicatie} | Suma acordată: {valoareAnualBP1:F2} lei | Rămas fond: {sumaDisponibila:F2} lei";
                    string comentariuAI = GenerateAIComment(student, null, sumaDisponibila, "BP1", motiv, "", etapa, sortedStudents);

                    istoricList.Add((student.Emplid, new BursaIstoric
                    {
                        StudentRecordId = student.Id,
                        TipBursa = "BP1",
                        Actiune = "Acordare",
                        Suma = valoareBP1Actuala,
                        Motiv = motiv,
                        Comentarii = comentariu,
                        ComentariiAI = comentariuAI,
                        DataModificare = DateTime.Now,
                        Etapa = etapa
                    }));
                }
                else
                {
                    student.Bursa = null;
                    student.SumaBursa = 0;
                }
            }
            int bp2AllocatedCount = 0;
            foreach (var student in sortedStudents.Where(s => String.IsNullOrEmpty(s.Bursa)).OrderByDescending(s => s.Media))
            {
                bool esteRP = !string.Equals(student.TaraCetatenie, "ROU", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(student.TaraCetatenie);
                decimal valoareBP2Actuala = esteRP ? valoareAnualBP2RP : valoareAnualBP2;
                bool isEligibleForBP2Criterion = student.Media >= 8.00M;

                if (isEligibleForBP2Criterion && bp2AllocatedCount < bestNumBP2 && sumaDisponibila >= valoareAnualBP2)
                {
                    student.Bursa = "BP2";
                    student.SumaBursa = valoareBP2Actuala;
                    sumaDisponibila -= valoareBP2Actuala;
                    bp2AllocatedCount++;

                    string motiv = "Media eligibilă pentru BP2.";
                    string explicatie = $"Media: {student.Media:F2}";
                    string comentariu = $"Etapa: {etapa} | Media: {student.Media:F2} | {motiv} | {explicatie} | Suma acordată: {valoareAnualBP2:F2} lei | Rămas fond: {sumaDisponibila:F2} lei";
                    string comentariuAI = GenerateAIComment(student, null, sumaDisponibila, "BP2", motiv, "", etapa, sortedStudents);

                    istoricList.Add((student.Emplid, new BursaIstoric
                    {
                        StudentRecordId = student.Id,
                        TipBursa = "BP2",
                        Actiune = "Acordare",
                        Suma = valoareBP2Actuala,
                        Motiv = motiv,
                        Comentarii = comentariu,
                        ComentariiAI = comentariuAI,
                        DataModificare = DateTime.Now
                    }));
                }
                else
                {
                    student.Bursa = null;
                    student.SumaBursa = 0;
                }
            }
            return (sumaDisponibila, istoricList);
        }
        private string GenerateAIComment(StudentRecord currentStudent, StudentRecord previousStudent, decimal remainingFunds, string scholarshipType, string reason, string fallback, string etapa, List<StudentRecord> allStudents)
        {
            string previousStudentInfo = previousStudent != null
                ? $"Studentul anterior: {previousStudent.NumeStudent} (media {previousStudent.Media:F2}, bursă {previousStudent.Bursa ?? "N/A"})"
                : "Acesta este primul student care primește bursă sau primul din categoria sa.";

            var nextEligibleStudents = allStudents
                .Where(s => string.IsNullOrWhiteSpace(s.Bursa) && s.Media >= 8.00M && s != currentStudent) // Considerăm toți studenții eligibili pentru o bursă merit
                .OrderByDescending(s => s.Media)
                .Take(3) // Afișăm următorii 3 studenți relevanți
                .Select(s => $"(Emplid: {s.Emplid}, Media: {s.Media:F2})")
                .ToList();

            string nextStudentsText = nextEligibleStudents.Any()
                ? $"Următorii potențiali beneficiari: {string.Join(", ", nextEligibleStudents)}"
                : "Nu mai sunt studenți eligibili cu bursă merit în acest moment.";

            return $"Studentul {currentStudent.NumeStudent} (Emplid: {currentStudent.Emplid}) cu media {currentStudent.Media:F2} a primit bursa de tip {scholarshipType} pentru că {reason.ToLower()}. " +
                   $"{(string.IsNullOrEmpty(fallback) ? "" : fallback + " ")}{previousStudentInfo}. {nextStudentsText}. " +
                   $"Fonduri rămase: {remainingFunds:F2} lei.";
        }
        private (decimal, List<(string Emplid, BursaIstoric Istoric)>) AssignOnlyBP2(
    List<StudentRecord> students,
    decimal sumaDisponibila,
    List<FondBurse> fonduri,
    Dictionary<int, decimal> sumaRamasaPeFond,
    string etapa,
    double? valoareRomaniDePretutindeni = null)
        {
            var istoricList = new List<(string Emplid, BursaIstoric Istoric)>();
            StudentRecord studentAnterior = null;

            foreach (var student in students)
            {
                bool esteRP = !string.Equals(student.TaraCetatenie, "ROU", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(student.TaraCetatenie);

                string domeniu = student.FondBurseMeritRepartizat?.domeniu;
                if (string.IsNullOrEmpty(domeniu))
                {
                    student.Bursa = null;
                    break;
                }

                Match match = Regex.Match(domeniu, @"\((\d+)\)");
                if (!match.Success)
                {
                    student.Bursa = null;
                    break;
                }

                int an = int.Parse(match.Groups[1].Value);
                string program = student.FondBurseMeritRepartizat?.programStudiu;
                int? fondId = student.FondBurseMeritRepartizatId;
                decimal valoareBP2 = (an == 4 || (program == "master" && an == 2))
                    ? fonduri[1].ValoreaLunara * 9.35M
                    : fonduri[1].ValoreaLunara * 12;
                decimal valoareAnualBP2RP = 0;
                if (valoareRomaniDePretutindeni != null)
                {
                    decimal valoareRP = (decimal)valoareRomaniDePretutindeni;
                    if (an == 4 || (program == "master" && an == 2))
                    {
                        valoareAnualBP2RP = valoareBP2 - (valoareRP * 9.35M);
                    }
                    else
                    {
                        valoareAnualBP2RP = valoareBP2 - (valoareRP * 12);
                    }
                }
                else
                {
                    valoareAnualBP2RP = valoareBP2; 
                }
                decimal valoareBP2Actuala = esteRP ? valoareAnualBP2RP : valoareBP2;
                if (sumaDisponibila >= valoareBP2Actuala && student.Media >= 8.00M)
                {
                    student.Bursa = "BP2";
                    student.SumaBursa = valoareBP2Actuala;
                    sumaDisponibila -= valoareBP2Actuala;
                    if (fondId.HasValue)
                        sumaRamasaPeFond[fondId.Value] -= valoareBP2Actuala;
                    string infoDurata = (an == 4 || (program == "master" && an == 2)) ? "9.35 luni" : "12 luni";
                    string comentariu = $"Etapa: {etapa} | Media: {student.Media:F2} | " +
                        $"Acordare BP2 ({valoareBP2Actuala:F2} lei) – fond ID {fondId?.ToString() ?? "—"} | " +
                        $"Necesari: {valoareBP2Actuala:F2} lei | " +
                        $"Program: {program}, An: {an}, Durată: {infoDurata}";
                    var urmatorii = students
                        .Where(s => string.IsNullOrWhiteSpace(s.Bursa) && s.Id != student.Id)
                        .Take(5)
                        .Select(s => $"{s.Emplid} (Media: {s.Media:F2}, An: {s.An+1}, Domeniu: {s.FondBurseMeritRepartizat?.domeniu ?? "—"})")
                        .ToList();
                    string urmatoriiText = urmatorii.Count > 0
                        ? $"Următorii studenți eligibili: {string.Join(", ", urmatorii)}"
                        : "Nu mai sunt studenți eligibili în acest moment.";
                    string anterior = studentAnterior != null
                        ? $"Studentul anterior: {studentAnterior.NumeStudent} (media {studentAnterior.Media:F2}, bursă {studentAnterior.Bursa}, Domeniu: {studentAnterior.FondBurseMeritRepartizat?.domeniu ?? "—"})"
                        : "Acesta este primul student care primește bursă.";
                    string comentariuAI = $"Studentul {student.NumeStudent} (media {student.Media:F2}) a primit bursa de tip BP2 " +
                        $"pentru anul {an}, program {program}, cu durata {infoDurata}. " +
                        $"Fondurile disponibile au permis acordarea integrală a bursei din fondul #{fondId?.ToString() ?? "—"}. " +
                        $"{anterior}. {urmatoriiText}. Fonduri rămase: {sumaDisponibila:F2} lei.";
                    istoricList.Add((student.Emplid, new BursaIstoric
                    {
                        StudentRecordId = student.Id,
                        TipBursa = "BP2",
                        Actiune = "Acordare",
                        Suma = valoareBP2Actuala,
                        Motiv = "Acordare BP2 – fonduri suficiente",
                        Comentarii = comentariu,
                        ComentariiAI = comentariuAI,
                        DataModificare = DateTime.Now,
                        Etapa = etapa
                    }));
                    studentAnterior = student;
                }
                else
                {
                    student.Bursa = null;
                    break;
                }
            }

            return (sumaDisponibila, istoricList);
        }
        public static List<StudentRecord> EliminaStudentiNeeligibili(List<StudentRecord> students)
        {
            return students.Where(s => s.RO == 0 && s.TR == 0).ToList();
        }
        [Authorize]
        [HttpPost("EvaluateAcurracy")]
        public async void EvaluateAcurracy()
        {
            string filePath1 = @"C:\Users\grati\Downloads\New Microsoft Excel Worksheet.xlsx";
            string filePath2 = @"C:\Users\grati\Downloads\Burse_Studenti_Actualizat (5).xlsx";
            string outputPath = @"C:\Users\grati\Downloads\OutputFile_Differences.xlsx";

            using var wb1 = new XLWorkbook(filePath1);
            using var wb2 = new XLWorkbook(filePath2);
            using var wbOut = new XLWorkbook();

            var ws1 = wb1.Worksheet(1); // Domenii + BM1
            var ws2 = wb2.Worksheet(1); // Domenii + BM2
            var wsOut = wbOut.AddWorksheet("Comparatie");

            int startRow = 20;
            int colDomeniu = 1; // A
            int colBM1 = 4;     // D din Excel 1
            int colBM2 = 6;     // F din Excel 2

            // Header
            wsOut.Cell(1, 1).Value = "Domeniu";
            wsOut.Cell(1, 2).Value = "BM1 (B.Perf.1)";
            wsOut.Cell(1, 3).Value = "BM2 (B.Perf.2)";
            wsOut.Row(1).Style.Font.Bold = true;

            int maxRows = Math.Max(ws1.LastRowUsed().RowNumber(), ws2.LastRowUsed().RowNumber());

            for (int i = 0; i <= maxRows - startRow; i++)
            {
                int currentRow = startRow + i;
                int outputRow = i + 2;
                string domeniu = ws1.Cell(currentRow, colDomeniu).GetValue<string>().Trim();
                double bm1_excel1 = GetNumericValue(ws1.Cell(currentRow, colBM1));
                double bm1_excel2 = GetNumericValue(ws2.Cell(currentRow, colBM1));
                double bm2_excel1 = GetNumericValue(ws1.Cell(currentRow, colBM2));
                double bm2_excel2 = GetNumericValue(ws2.Cell(currentRow, colBM2));
                wsOut.Cell(outputRow, 1).Value = domeniu;
                var cellBM1 = wsOut.Cell(outputRow, 2);
                if (bm1_excel1 == bm1_excel2)
                {
                    cellBM1.Value = bm1_excel1;
                }
                else
                {
                    cellBM1.Value = $"{bm1_excel1} -> {bm1_excel2}";
                    cellBM1.Style.Fill.BackgroundColor = XLColor.LightPink;
                }
                var cellBM2 = wsOut.Cell(outputRow, 3);
                if (bm2_excel1 == bm2_excel2)
                {
                    cellBM2.Value = bm2_excel2; 
                }
                else
                {
                    cellBM2.Value = $"{bm2_excel2} -> {bm2_excel1}";
                    cellBM2.Style.Fill.BackgroundColor = XLColor.LightPink;
                }
            }
            wbOut.SaveAs(outputPath);
        }
        private double GetNumericValue(IXLCell cell)
        {
            if (cell == null || string.IsNullOrWhiteSpace(cell.GetValue<string>()))
                return 0;

            double.TryParse(cell.GetValue<string>(), out double result);
            return result;
        }
        private int ParseIntOrZero(string s)
        {
            return int.TryParse(s, out int result) ? result : 0;
        }
        [Authorize]
        [HttpGet("situatie-studenti")]
        public async Task<IActionResult> GetSituatieStudenti()
        {
            try
            {
                List<FondBurse> fonduri = await _fondBurseService.GetDateFromBursePerformanteAsync();
                List<FormatiiStudii> formatiiStudii = await _fondBurseService.GetAllFromFormatiiStudiiAsync();
                string etapa0Path = Path.Combine(Path.GetTempPath(), "SituatieStudenti_modificati.xlsx");
                byte[] initialFileBytes = await _fondBurseService.GenerateCustomLayout2(etapa0Path, fonduri, formatiiStudii, 1671770.95m);
                await System.IO.File.WriteAllBytesAsync(etapa0Path, initialFileBytes);
                List<StudentRecord> totiStudentii = await _fondBurseService.GetStudentsWithBursaFromDatabaseAsync();
                List<StudentScholarshipData> studentiClasificati0 = totiStudentii
                    .GroupBy(s => new { s.FondBurseMeritRepartizatId, s.FondBurseMeritRepartizat?.domeniu })
                    .Select(group => new StudentScholarshipData
                    {
                        FondBurseId = group.Key.FondBurseMeritRepartizatId,
                        Domeniu = group.Key.domeniu,
                        BP1Count = group.Count(s => s.Bursa?.ToLower().Contains("bp1") ?? false),
                        BP2Count = group.Count(s => s.Bursa?.ToLower().Contains("bp2") ?? false)
                    }).ToList();
                using var input = new FileStream(etapa0Path, FileMode.Open, FileAccess.Read);
                var updatedStream = ExcelUpdater.UpdateScholarshipCounts(input, studentiClasificati0);
                updatedStream.Position = 0;
                using var workbook = new ClosedXML.Excel.XLWorkbook(updatedStream);
                var worksheet = workbook.Worksheets.Add("Toți studenții");
                var headers = new[]
                {
                    "Nr. crt.", "Emplid", "CNP", "Nume Student", "Țară Cetățenie",
                    "An", "Media", "Punctaj An", "CO", "RO", "TC", "TR",
                    "Sursa de finanțare", "Domeniu", "Bursa", "Suma Bursă"
                };
                for (int i = 0; i < headers.Length; i++)
                {
                    var cell = worksheet.Cell(1, i + 1);
                    cell.Value = headers[i];
                    cell.Style.Font.Bold = true;
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    cell.Style.Fill.BackgroundColor = XLColor.LightGray;
                }
                int row = 2;
                int nrCrt = 1;
                var studentiSortati = totiStudentii
                    .Where(s => !string.IsNullOrEmpty(s.Bursa) && s.Bursa != "NU")
                    .OrderBy(s => s.FondBurseMeritRepartizat?.domeniu)
                    .ThenBy(s => s.Media)
                    .ToList();
                string domeniuCurent = null;
                foreach (var s in studentiSortati)
                {
                    if (domeniuCurent != s.FondBurseMeritRepartizat?.domeniu)
                    {
                        if (row > 2)
                            row++;
                        domeniuCurent = s.FondBurseMeritRepartizat?.domeniu;
                        worksheet.Cell(row, 1).Value = domeniuCurent;
                        worksheet.Range(row, 1, row, headers.Length).Merge();
                        var domeniuCell = worksheet.Cell(row, 1);
                        domeniuCell.Style.Font.Bold = true;
                        domeniuCell.Style.Fill.BackgroundColor = XLColor.LightBlue;
                        domeniuCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                        row++;
                    }
                    worksheet.Cell(row, 1).Value = nrCrt++;
                    worksheet.Cell(row, 2).Value = s.Emplid;
                    worksheet.Cell(row, 3).Value = s.CNP;
                    worksheet.Cell(row, 4).Value = s.NumeStudent;
                    worksheet.Cell(row, 5).Value = s.TaraCetatenie; 
                    worksheet.Cell(row, 6).Value = s.An + 1;
                    worksheet.Cell(row, 7).Value = s.Media;
                    worksheet.Cell(row, 8).Value = s.PunctajAn;
                    worksheet.Cell(row, 9).Value = s.CO;
                    worksheet.Cell(row, 10).Value = s.RO;
                    worksheet.Cell(row, 11).Value = s.TC;
                    worksheet.Cell(row, 12).Value = s.TR;
                    worksheet.Cell(row, 13).Value = s.SursaFinantare;
                    worksheet.Cell(row, 14).Value = domeniuCurent;
                    worksheet.Cell(row, 15).Value = s.Bursa ?? "";
                    worksheet.Cell(row, 16).Value = s.SumaBursa.ToString("0.00");
                    var range = worksheet.Range(row, 1, row, headers.Length);
                    range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                    range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    row++;
                }
                worksheet.Columns().AdjustToContents();
                string ExtractGrup(string domeniu)
                {
                    if (string.IsNullOrEmpty(domeniu))
                        return "Fără grup";
                    var match = System.Text.RegularExpressions.Regex.Match(domeniu.Trim(), @"^[A-Za-z]+");
                    return match.Success ? match.Value.ToUpper() : domeniu.ToUpper();
                }
                var grupuri = studentiSortati
                    .GroupBy(s => ExtractGrup(s.FondBurseMeritRepartizat?.domeniu))
                    .OrderBy(g => g.Key);
                foreach (var grup in grupuri)
                {
                    string denumireFoaie = grup.Key.Length > 31 ? grup.Key.Substring(0, 31) : grup.Key;
                    worksheet = workbook.Worksheets.Add(denumireFoaie);
                    for (int i = 0; i < headers.Length; i++)
                    {
                        var cell = worksheet.Cell(1, i + 1);
                        cell.Value = headers[i];
                        cell.Style.Font.Bold = true;
                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        cell.Style.Fill.BackgroundColor = XLColor.LightGray;
                    }
                     row = 2;
                     nrCrt = 1;
                    foreach (var s in grup)
                    {
                        worksheet.Cell(row, 1).Value = nrCrt++;
                        worksheet.Cell(row, 2).Value = s.Emplid;
                        worksheet.Cell(row, 3).Value = s.CNP;
                        worksheet.Cell(row, 4).Value = s.NumeStudent;
                        worksheet.Cell(row, 5).Value = s.TaraCetatenie;
                        worksheet.Cell(row, 6).Value = s.An + 1;
                        worksheet.Cell(row, 7).Value = s.Media;
                        worksheet.Cell(row, 8).Value = s.PunctajAn;
                        worksheet.Cell(row, 9).Value = s.CO;
                        worksheet.Cell(row, 10).Value = s.RO;
                        worksheet.Cell(row, 11).Value = s.TC;
                        worksheet.Cell(row, 12).Value = s.TR;
                        worksheet.Cell(row, 13).Value = s.SursaFinantare;
                        worksheet.Cell(row, 14).Value = s.FondBurseMeritRepartizat?.domeniu ?? "";
                        worksheet.Cell(row, 15).Value = s.Bursa ?? "";
                        worksheet.Cell(row, 16).Value = s.SumaBursa.ToString("0.00");
                        var range = worksheet.Range(row, 1, row, headers.Length);
                        range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                        range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                        range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        row++;
                    }
                    worksheet.Columns().AdjustToContents();
                }
                using var finalStream = new MemoryStream();
                workbook.SaveAs(finalStream);
                finalStream.Position = 0;
                return File(finalStream.ToArray(),
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            "SituatieStudenti_modificati.xlsx");
            }
            catch (Exception ex)
            {
                return BadRequest($"❌ Eroare la generarea fișierului: {ex.Message}");
            }
        }
        [Authorize]
        [HttpPost("reset-burse")]
        public async Task<IActionResult> ResetBurseAsync()
        {
            try
            {
                await _context.Database.ExecuteSqlRawAsync("DELETE FROM FondBurseMeritRepartizat");
                return Ok(new { message = "înregistrările au fost resetate." });
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"❌ Eroare la resetarea burselor: {ex.Message}");
            }
        }
    }
}
