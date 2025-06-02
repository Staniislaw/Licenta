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

        public FondBurseController(BurseDBContext context, IFondBurseService fondBurseService, IFondBurseMeritRepartizatService fondBurseMeritRepartizatService, GrupuriDomeniiHelper grupuriHelper, IBurseIstoricService burseIstoricService, AppLogger logger, IGrupuriService grupuriService)
        {
            _context = context;
            _fondBurseService = fondBurseService;
            _fondBurseMeritRepartizatService = fondBurseMeritRepartizatService;
            _grupuriHelper = grupuriHelper;
            _burseIstoricService = burseIstoricService;
            _logger = logger;
            _grupuriService = grupuriService;
        }

        [HttpPost("AddFondBurse")]
        public async Task<IActionResult> AddFondBurse(List<IFormFile> files)
        {
            //var filePath = "C:\\Users\\Stas\\Downloads\\Fond_burse_2024_2025 13noiembrie.xls"; 
            //var filePath = "D:\\Licenta\\Fond_burse_2024_2025 13noiembrie.xls"; 
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

            //var filePath2 = "C:\\Users\\Stas\\Downloads\\Formatii studii USV_1 octombrie 2024 finantare.xlsx";
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
        public async Task<IActionResult> GenerateExcel(decimal disponibilBM = 1671770.95m)
        {
            try
            {
                List<FondBurse> fonduri = await _fondBurseService.GetDateFromBursePerformanteAsync();
                List<FormatiiStudii> formatiiStudii = await _fondBurseService.GetAllFromFormatiiStudiiAsync();
                // 📌 Calea temporară unde fișierul va fi generat
                string filePath = Path.Combine(Path.GetTempPath(), "Burse_Studenți.xlsx");

                // Generăm fișierul Excel
                byte[] fileBytes= await _fondBurseService.GenerateCustomLayout2(filePath, fonduri, formatiiStudii, disponibilBM);

                // Citim fișierul și îl returnăm ca răspuns HTTP
                //byte[] fileBytes = await System.IO.File.ReadAllBytesAsync(filePath);
               
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

        [HttpPost("process")]
        public async Task<IActionResult> ProcessExcelFiles([FromForm] List<IFormFile> pathStudentiList,IFormFile burseFile, [FromQuery]decimal? epsilonValue = 0.05M)
        {
            //await _fondBurseService.ResetSumaRamasaAsync();
            //await _fondBurseService.ResetStudentiAsync();
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

                    // Tratare exactă pentru foi numerice în IEN/IETTI
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
                        // Tratăm foi de genul: 1rcc, 2sc, etc.
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



            // 3. Verifică ce domenii lipsesc
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
                var studentRecords =excelReader.ReadStudentRecordsFromExcel(stream, pathStudenti.FileName, domenii);

                // Procesăm fiecare listă de studenți înainte de a o adăuga
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

                    if (processedCount != groupedCount)
                    {
                        if (processedCount != groupedCount)
                        {
                            var msg = $"Numărul studenților pentru Domeniul '{kvp.Key}' în fișierul '{pathStudenti.FileName}' nu coincide. Procesați: {processedCount}, Preluați din fișierul FormatiiStudii: {groupedCount}";
                            _logger.LogStudentsExcels(msg);
                            discrepante.Add(msg); 
                        }
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
               /// using var stream = pathStudenti.OpenReadStream();

               // Dictionary<string, List<StudentRecord>> studentRecords = excelReader.ReadStudentRecordsFromExcel(stream, pathStudenti.FileName);
                var istoricList = new List<(string Emplid, BursaIstoric Istoric)>();

                foreach (var entry in studentRecords)
                {

                    string domeniu = entry.Key;
                    List<StudentRecord> students = ProcessStudents(entry.Value);
                    FondBurseMeritRepartizat? fondRepartizatByDomeniu = await _fondBurseMeritRepartizatService.GetByDomeniuAsync(domeniu);

                    if (fondRepartizatByDomeniu == null) continue;

                    (decimal valoareAnualBP1, decimal valoareAnualBP2) = CalculateScholarshipValues(domeniu, fonduri, fondRepartizatByDomeniu);

                    decimal sumaDisponibila = fondRepartizatByDomeniu.bursaAlocatata;
                    if (sumaDisponibila < 0)
                        continue;
                    (var sumaRamasa, var istoricePerDomeniu) = AssignScholarshipsOptimezedWithCriteriaMediilorAceleasi(
                        students,
                        sumaDisponibila,
                        valoareAnualBP1,
                        valoareAnualBP2,
                        epsilon,
                        fondRepartizatByDomeniu,
                        "0"
                    );
                    sumaDisponibila = sumaRamasa;

                    istoricList.AddRange(istoricePerDomeniu);


                    students.ForEach(s => s.FondBurseMeritRepartizatId = fondRepartizatByDomeniu.ID);
                    var groupedByMedia = students.GroupBy(s => s.Media);

                    foreach (var group in groupedByMedia)
                    {
                        // Colectăm toate valorile burselor distincte (inclusiv null)
                        var valoriDistincteBursa = group
                            .Select(s => string.IsNullOrWhiteSpace(s.Bursa) ? null : s.Bursa.Trim())
                            .Distinct()
                            .ToList();


                        // Dacă sunt 2 sau mai multe valori distincte, înseamnă inconsistență
                        if (valoriDistincteBursa.Count > 1)
                        {
                            var studentiCuAceeasiMedia = group.Select(s =>
                                $"Emplid: {s.Emplid}, Nume: {s.NumeStudent}, Bursa: {(string.IsNullOrWhiteSpace(s.Bursa) ? "NU" : s.Bursa)}, Program: {fondRepartizatByDomeniu.domeniu}"
                            );

                            var mesaj = $"⚠️ Atenție! etapa 0: Studenți cu media {group.Key} au situație mixtă la bursă (valori diferite):\n" +
                                        string.Join("\n", studentiCuAceeasiMedia);

                            _logger.LogStudentInfo(mesaj);
                        }
                    }
                    var studentiCuIdCorect = await _fondBurseService.SaveNewStudentsAsync(students);

                    // ✅ actualizezi StudentRecordId în istoricul generat anterior
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

                    foreach (var i in istoricList)
                    {
                        Console.WriteLine($"Emplid: {i.Emplid}, StudentRecordId: {i.Istoric.StudentRecordId}, Bursa: {i.Istoric.TipBursa}, Suma: {i.Istoric.Suma}");
                    }

                    // Opțional: aruncă mai departe excepția dacă vrei să o tratezi mai sus
                    // throw;
                }

            }
            //verificare 



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

            foreach (var item in studentiClasificati0)
            {
                Console.WriteLine($"Domeniu: {item.Domeniu}, BP1: {item.BP1Count}, BP2: {item.BP2Count}");
            }

            /*using var inputStream = burseFile.OpenReadStream();
            var updatedStream = ExcelUpdater.UpdateScholarshipCounts(inputStream, studentiClasificati1);
            */


            //verificare 


            //INCEPERE ALGORITM DE RE REPARTIZARE A BURSELOR ADICA OFERI DIN NOU BURSELE IN FUNCTIE DE GRUPURI DE DOMENII
             var studentiPeGrupa = await _fondBurseService.GetStudentiEligibiliPeGrupaAsync();
             var fonduriRepartizate = await _fondBurseMeritRepartizatService.GetAllAsync();

             var sumaDisponibilaPeGrupa = fonduriRepartizate
                 .GroupBy(f => f.Grupa)
                 .ToDictionary(
                     g => g.Key,
                     g => g.Sum(f => f.SumaRamasa)
                 );

             var fonduriDict = fonduriRepartizate.ToDictionary(f => f.ID, f => f);

             // 📌 Domeniile sunt unice aici, deci putem păstra și o mapare Domeniu → Fond
             var fonduriPeDomeniu = fonduriRepartizate.ToDictionary(f => f.domeniu, f => f);

             foreach (var entry in studentiPeGrupa)
             {
                 string grupa = entry.Key;
                 List<StudentRecord> students = entry.Value;

                 // 🧮 Obținem toate fondurile care aparțin grupei
                 var fonduriGrupa = fonduriRepartizate
                     .Where(f => f.Grupa == grupa)
                     .ToList();

                 if (!fonduriGrupa.Any()) continue;

                 var sumaRamasaPeFond = fonduriGrupa.ToDictionary(f => f.ID, f => f.SumaRamasa);

                 // Luăm suma disponibilă per grupă
                 decimal sumaDisponibila = sumaDisponibilaPeGrupa[grupa];
                 if (sumaDisponibila < 0)
                     continue;
                 // ✅ Atribuim DOAR BP2 pe această grupă
                 //AssignOnlyBP2   (students, ref sumaDisponibila, fonduri, sumaRamasaPeFond);

                 (decimal sumaNoua, var istoricBP2) = AssignOnlyBP2(students, sumaDisponibila, fonduri, sumaRamasaPeFond,"1");
                 sumaDisponibila = sumaNoua;

                var groupedByMedia = students.GroupBy(s => s.Media);

                foreach (var group in groupedByMedia)
                {
                    // Colectăm toate valorile burselor distincte (inclusiv null)
                    var valoriDistincteBursa = group
                        .Select(s => string.IsNullOrWhiteSpace(s.Bursa) ? null : s.Bursa.Trim())
                        .Distinct()
                        .ToList();


                    // Dacă sunt 2 sau mai multe valori distincte, înseamnă inconsistență
                    if (valoriDistincteBursa.Count > 1)
                    {
                        var studentiCuAceeasiMedia = group.Select(s =>
                            $"Emplid: {s.Emplid}, Nume: {s.NumeStudent}, Bursa: {(string.IsNullOrWhiteSpace(s.Bursa) ? "NU" : s.Bursa)}, Program: {s.FondBurseMeritRepartizat.domeniu}"
                        );

                        var mesaj = $"⚠️ Atenție! etapa1: Studenți cu media {group.Key} au situație mixtă la bursă (valori diferite):\n" +
                                    string.Join("\n", studentiCuAceeasiMedia);

                        _logger.LogStudentInfo(mesaj);
                    }
                }


                await _fondBurseService.SaveNewStudentsAsync(students);

                 // 🔁 Update la suma rămasă pentru TOATE domeniile din acea grupă

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
            
            // COUNT BURSE BP1 SI BPS2 SI INTRODUCEARE IN EXCEL
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


            //ExcelUpdater.UpdateScholarshipCounts("D:\\Licenta\\Burse_Studenți (1).xlsx", studentiClasificati4);

            //PASUL 2 PRELUAM SUMELE DISPONIBILE PENTRU LICENTA/MASTER SI OFERIM BURSE IN FUNCTIE DE MEDIE



            //GRESIT  CRED SE FACE REPARTIZAREA DUPA LICENTA/MASTER
            /*  var sumaDisponibilaPeProgram = fonduriRepartizate
                  .GroupBy(f => f.programStudiu)
                  .ToDictionary(
                      g => g.Key,
                      g => g.Sum(f => f.SumaRamasa)
                  );
              var studentiPeProgram = await _fondBurseService.GetStudentiEligibiliPeProgramAsync();

              foreach (var entry in studentiPeProgram)
              {
                  string programStudiu = entry.Key;
                  List<StudentRecord> students = entry.Value;

                  // 🧮 Obținem toate fondurile care aparțin grupei
                  var fonduriGrupa = fonduriRepartizate
                      .Where(f => f.programStudiu == programStudiu)
                      .ToList();

                  if (!fonduriGrupa.Any()) continue;

                  var sumaRamasaPeProgramStudiu = fonduriGrupa.ToDictionary(f => f.ID, f => f.SumaRamasa);

                  // Luăm suma disponibilă per grupă
                  decimal sumaDisponibila = sumaDisponibilaPeProgram[programStudiu];
                  if (sumaDisponibila < 0)
                      continue;
                  // ✅ Atribuim DOAR BP2 pe această grupă
                  //AssignOnlyBP2(students, ref sumaDisponibila, fonduri, sumaRamasaPeProgramStudiu);
                  (decimal sumaNoua, var istoricBP2) = AssignOnlyBP2(students, sumaDisponibila, fonduri, sumaRamasaPeProgramStudiu, "3");
                  sumaDisponibila = sumaNoua;
                  await _fondBurseService.SaveNewStudentsAsync(students);

                  // 🔁 Update la suma rămasă pentru TOATE domeniile din acea grupă

                  foreach (var fond in fonduriGrupa)
                  {
                      fond.SumaRamasa = sumaRamasaPeProgramStudiu[fond.ID];
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
                          existing.Comentarii = hist.Istoric.Comentarii;
                      }
                      else
                      {
                          await _context.BursaIstoric.AddAsync(hist.Istoric);
                      }
                  }

                  await _context.SaveChangesAsync();
              }*/

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

                //AssignOnlyBP2(students, ref sumaDisponibila, fonduri, sumaRamasaPeFond);
                (decimal sumaNoua, var istoricBP2) = AssignOnlyBP2(students, sumaDisponibila, fonduri, sumaRamasaPeFond, "2");
                sumaDisponibila = sumaNoua;


                var groupedByMedia = students.GroupBy(s => s.Media);

                foreach (var group in groupedByMedia)
                {
                    // Colectăm toate valorile burselor distincte (inclusiv null)
                    var valoriDistincteBursa = group
                        .Select(s => string.IsNullOrWhiteSpace(s.Bursa) ? null : s.Bursa.Trim())
                        .Distinct()
                        .ToList();


                    // Dacă sunt 2 sau mai multe valori distincte, înseamnă inconsistență
                    if (valoriDistincteBursa.Count > 1)
                    {
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

            // COUNT BURSE BP1 SI BPS2 SI INTRODUCEARE IN EXCEL
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

            foreach (var item in studentiClasificati2)
            {
                Console.WriteLine($"Domeniu: {item.Domeniu}, BP1: {item.BP1Count}, BP2: {item.BP2Count}");
            }

            //ExcelUpdater.UpdateScholarshipCounts("D:\\Licenta\\Burse_Studenți (3).xlsx", studentiClasificati2);


            //PASUL 3 OFERIM BURSE PE GRUPUIRIDE DOMENII
            // PASUL 3 – Repartizare burse pe grupuri de domenii
            var grupuriBurse = await _grupuriHelper.GetGrupuriBurseAsync();

            foreach (var grup in grupuriBurse)
            {
                string numeGrup = grup.Key;
                List<string> domeniiGrup = grup.Value;

                // Fondurile din acest grup
                var fonduriInGrup = fonduriRepartizate
                    .Where(f => GetDomeniiDinGrupa(f.Grupa)
                        .Any(domeniu => domeniiGrup.Contains(domeniu)))
                    .ToList();

                if (!fonduriInGrup.Any()) continue;

                // Grupare pe program de studiu (sau domeniu), și determinare fracțiune
                var fonduriPeProgram = fonduriInGrup
                    .Where(f => f.programStudiu?.ToLower() == "licenta")
                    .GroupBy(f => f.domeniu)
                    .Select(grupProgram => new
                    {
                        ProgramStudiu = grupProgram.Key,
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

                // Studenții eligibili doar din programul respectiv și doar licență

                var studentiEligibili = fonduriPeProgram.Fonduri
                    .SelectMany(f => f.Studenti) 
                    .Where(s => s.Bursa == null || s.Bursa.Trim().ToLower() == "nicio bursă")
                    .OrderByDescending(s => s.Media)
                    .ToList();


                (decimal sumaNoua, var istoricBP2) = AssignOnlyBP2(
                    studentiEligibili,
                    sumaDisponibila,
                    fonduri,
                    sumaRamasaPeFond,
                    "3"
                );


                var groupedByMedia = studentiEligibili.GroupBy(s => s.Media);

                foreach (var group in groupedByMedia)
                {
                    // Colectăm toate valorile burselor distincte (inclusiv null)
                    var valoriDistincteBursa = group
                        .Select(s => string.IsNullOrWhiteSpace(s.Bursa) ? null : s.Bursa.Trim())
                        .Distinct()
                        .ToList();


                    // Dacă sunt 2 sau mai multe valori distincte, înseamnă inconsistență
                    if (valoriDistincteBursa.Count > 1)
                    {
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


            //etapa gresita se ofera burse pentru 1 singru grup cand trebuie pentru fiecare grup in parte gresit mai mult seamana cu etapa 4 mai mult.
            /*var grupuriBurse = await _grupuriHelper.GetGrupuriBurseAsync();

            string grupCastigatorNume = null;
            List<string> domeniiGrupCastigator = null;
            List<FondBurseMeritRepartizat> fonduriCastigatoare = null;
            decimal fractiuneMaxima = 0;
            decimal sumaDisponibilaAdd = 0;
            foreach (var grup in grupuriBurse)
            {
                string numeGrup = grup.Key;
                List<string> domeniiGrup = grup.Value;

                var fonduriInGrup = fonduriRepartizate
                    .Where(f =>
                        GetDomeniiDinGrupa(f.Grupa)
                            .Any(domeniu => domeniiGrup.Contains(domeniu)))
                    .ToList();

                if (!fonduriInGrup.Any()) continue;

                decimal sumaDisponibila = fonduriInGrup.Sum(f => f.SumaRamasa);
                decimal sumaInitiala = fonduriInGrup.Sum(f => f.bursaAlocatata);
                decimal fractiune = sumaDisponibila / sumaInitiala;

                if (sumaDisponibila <= 0) continue;
                sumaDisponibilaAdd += sumaDisponibila;

                if (fractiune > fractiuneMaxima)
                {
                    fractiuneMaxima = fractiune;
                    grupCastigatorNume = numeGrup;
                    domeniiGrupCastigator = domeniiGrup;
                    fonduriCastigatoare = fonduriInGrup;
                }
            }

            // ❗ Verificare dacă am un grup câștigător
            if (domeniiGrupCastigator == null || !fonduriCastigatoare.Any())
                return BadRequest("Nu există fonduri disponibile în niciun grup.");
            var fonduriInGrupFinal = fonduriRepartizate
                    .Where(f =>
                        GetDomeniiDinGrupa(f.Grupa)
                            .Any(domeniu => domeniiGrupCastigator.Contains(domeniu)))
                    .ToList();

            var sumaRamasaPeFondFinal = fonduriInGrupFinal.ToDictionary(f => f.ID, f => f.SumaRamasa);

            // 🎓 Toți studenții eligibili din domeniile grupului
            var studentiLicentaFinal = (await _fondBurseService
                .GetStudentiEligibiliPeDomeniiAsync(domeniiGrupCastigator))
                .Where(s => s.FondBurseMeritRepartizat?.programStudiu?.ToLower() == "licenta")
                .OrderByDescending(s => s.Media)
                .ToList();


            studentiLicentaFinal = studentiLicentaFinal
                   .OrderByDescending(s => s.Media)
                   .ToList();

            (decimal sumaNouaFinal, var istoricBP2Final) = AssignOnlyBP2(studentiLicentaFinal, sumaDisponibilaAdd, fonduri, sumaRamasaPeFondFinal, "3");
            sumaDisponibilaAdd = sumaNouaFinal;
            // 💾 Salvare
            await _fondBurseService.SaveNewStudentsAsync(studentiLicentaFinal);
            foreach (var fond in fonduriInGrupFinal)
            {
                fond.SumaRamasa = sumaRamasaPeFondFinal[fond.ID];
                await _fondBurseMeritRepartizatService.UpdateAsync(fond);
            }
            foreach (var entry in istoricBP2Final)
            {
                var existing = await _context.BursaIstoric.FirstOrDefaultAsync(x =>
                    x.StudentRecordId == entry.Istoric.StudentRecordId
                );

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
            //etapa gresita se ofera burse pentru 1 singru grup cand trebuie pentru fiecare grup in parte gresit mai mult seamana cu etapa 4 mai mult.
/*

/* foreach (var grup in grupuriBurse)
{
string numeGrup = grup.Key;
List<string> domeniiGrup = grup.Value;

// 🔎 Fondurile de licență pentru acest grup (toate domeniile)
var fonduriInGrup = fonduriRepartizate
.Where(f =>
f.programStudiu == "licenta" &&
GetDomeniiDinGrupa(f.Grupa)
.Any(domeniu => domeniiGrup.Contains(domeniu)))
.ToList();


if (!fonduriInGrup.Any()) continue;

// 🔢 Suma totală disponibilă în grup
decimal sumaDisponibila = fonduriInGrup.Sum(f => f.SumaRamasa);

if (sumaDisponibila<0)
continue;
// 🔁 Dicționar cu suma pe fiecare fond (pentru update ulterior)
var sumaRamasaPeFond = fonduriInGrup.ToDictionary(f => f.ID, f => f.SumaRamasa);

// 🎓 Toți studenții eligibili din domeniile grupului
var studentiLicenta = await _fondBurseService
.GetStudentiEligibiliPeDomeniiAsync(domeniiGrup);

if (!studentiLicenta.Any()) continue;

// 🔽 Sortează după medie descrescător
studentiLicenta = studentiLicenta
.OrderByDescending(s => s.Media)
.ToList();

// 🏆 Atribuire burse doar BP2
//AssignOnlyBP2(studentiLicenta, ref sumaDisponibila, fonduri, sumaRamasaPeFond);
(decimal sumaNoua, var istoricBP2) = AssignOnlyBP2(studentiLicenta, sumaDisponibila, fonduri, sumaRamasaPeFond, "3");
sumaDisponibila = sumaNoua;
// 💾 Salvare
await _fondBurseService.SaveNewStudentsAsync(studentiLicenta);

// 📦 Update pe toate fondurile
foreach (var fond in fonduriInGrup)
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
existing.Etapa = "3";
existing.Comentarii = entry.Istoric.Comentarii;
}
else
{
await _context.BursaIstoric.AddAsync(entry.Istoric);
}
}

await _context.SaveChangesAsync();
}*/




            // COUNT BURSE BP1 SI BPS2 SI INTRODUCEARE IN EXCEL
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

            //ExcelUpdater.UpdateScholarshipCounts("D:\\Licenta\\Burse_Studenți (4).xlsx", studentiClasificati3);



            // 🔁 PUNCTUL 4: Redistribuire fond rămas către grupul cu cea mai mare fracțiune financiară
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

                    //decimal fractiune = sumaInitiala > 0 ? sumaRamasa / sumaInitiala : 0;
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

            //preiau grupul cu sumaMaxima 
            var grupCuSumaMaxima = grupuriCuSumaRamasa.FirstOrDefault();

            if (grupCuSumaMaxima != null && grupCuSumaMaxima.SumaRamasa > 0)
            {
                Console.WriteLine($"\n📌 PUNCTUL 4 – Grup cu cea mai mare sumă rămasă: {grupCuSumaMaxima.NumeGrup} ({grupCuSumaMaxima.SumaRamasa} lei)");

                var studentiGrup = await _fondBurseService.GetStudentiEligibiliPeDomeniiAsync(grupCuSumaMaxima.Domenii);

                if (studentiGrup.Any())
                {
                    //preiau toti studentii din acel grup pentru a ii oferi burse
                    var fonduriGrup = fonduriRepartizate
                        .Where(f =>
                            GetDomeniiDinGrupa(f.Grupa)
                                .Any(d => grupCuSumaMaxima.Domenii.Contains(d)))
                        .ToList();

                    var sumaRamasaPeFond = fonduriGrup.ToDictionary(f => f.ID, f => f.SumaRamasa);
                    //decimal sumaDisponibila = grupCuSumaMaxima.SumaRamasa;
                    decimal sumaDisponibila = fonduriRepartizate.Sum(f => f.SumaRamasa);

                    // 🔽 Sortează după medie descrescător
                    studentiGrup = studentiGrup
                        .OrderByDescending(s => s.Media)
                        .ToList();

                    //AssignOnlyBP2(studentiGrup, ref sumaDisponibila, fonduri, sumaRamasaPeFond);
                    (decimal sumaNoua, var istoricBP2) = AssignOnlyBP2(studentiGrup, sumaDisponibila, fonduri, sumaRamasaPeFond, "4");
                    sumaDisponibila = sumaNoua;

                    var groupedByMedia = studentiGrup.GroupBy(s => s.Media);

                    foreach (var group in groupedByMedia)
                    {
                        // Colectăm toate valorile burselor distincte (inclusiv null)
                        var valoriDistincteBursa = group
                            .Select(s => string.IsNullOrWhiteSpace(s.Bursa) ? null : s.Bursa.Trim())
                            .Distinct()
                            .ToList();


                        // Dacă sunt 2 sau mai multe valori distincte, înseamnă inconsistență
                        if (valoriDistincteBursa.Count > 1)
                        {
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
            List<StudentScholarshipData> studentiClasificati4 = studentiCuBursa4
                .GroupBy(s => new { s.FondBurseMeritRepartizatId, s.FondBurseMeritRepartizat.domeniu })
                .Select(group => new StudentScholarshipData
                {
                    FondBurseId = group.Key.FondBurseMeritRepartizatId,
                    Domeniu = group.Key.domeniu,
                    BP1Count = group.Count(s => s.Bursa.ToLower().Contains("bp1")),
                    BP2Count = group.Count(s => s.Bursa.ToLower().Contains("bp2"))
                }).ToList();


            string etapa0Path = $"C:\\Licenta\\Etapa_0.xlsx";
            using (var fileStream = new FileStream(etapa0Path, FileMode.Create, FileAccess.Write))
           {
                using var initialStream = burseFile.OpenReadStream();
                await initialStream.CopyToAsync(fileStream);
            }

            // 🔁 Etape de procesare și salvare
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
                string etapaOutputPath = $"C:\\Licenta\\Etapa_{i + 1}.xlsx";

                using var input = new FileStream(etapaInputPath, FileMode.Open, FileAccess.Read);
                using var output = new FileStream(etapaOutputPath, FileMode.Create, FileAccess.Write);

                var updatedStream = ExcelUpdater.UpdateScholarshipCounts(input, toateEtapele[i]);
                updatedStream.Position = 0;
                await updatedStream.CopyToAsync(output);

                previousPath = etapaOutputPath; // pentru următoarea rundă
            }

            // 🟢 La final, returnăm ultimul fișier generat pentru download
            string finalFilePath = previousPath;
            var finalBytes = await System.IO.File.ReadAllBytesAsync(finalFilePath);
            return File(
                finalBytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                Path.GetFileName(finalFilePath)
            );

        }


        // GET DOAR DOMENIU IETTI/ RST SUNT DESPARTITE
        public static List<string> GetDomeniiDinGrupa(string grupa)
        {
            if (string.IsNullOrEmpty(grupa))
                return new List<string>();

            // Normalizezi ex: "IETTI/RST" => ["IETTI", "RST"]
            return grupa.Split('/')
                .Select(d => d.Trim().ToUpper()) // sau păstrezi lowercase, după cum ai în GrupuriBurse
                .ToList();
        }
        /// <summary>
        /// Elimină studenții neeligibili și îi sortează descrescător după medie.
        /// </summary>
        private List<StudentRecord> ProcessStudents(List<StudentRecord> students)
        {
            return EliminaStudentiNeeligibili(students).OrderByDescending(s => s.Media).ToList();
        }
        private int ExtractAnDinDomeniu(string domeniu)
        {
            Match match = Regex.Match(domeniu, @"\((\d+)\)");
            return match.Success ? int.Parse(match.Groups[1].Value) : 0;
        }

        /// <summary>
        /// Calculează valoarea anuală a burselor BP1 și BP2 în funcție de domeniu.
        /// </summary>
        private (decimal, decimal) CalculateScholarshipValues(string domeniu, List<FondBurse> fonduri, FondBurseMeritRepartizat fondRepartizat)
        {
            decimal valoareBP1, valoareBP2;

            if (domeniu.Contains("4") || (fondRepartizat.programStudiu == "master" && domeniu.Contains("2")))
            {
                valoareBP1 = fonduri[0].ValoreaLunara * 9.35M;
                valoareBP2 = fonduri[1].ValoreaLunara * 9.35M;
            }
            else
            {
                valoareBP1 = fonduri[0].ValoreaLunara * 12;
                valoareBP2 = fonduri[1].ValoreaLunara * 12;
            }

            return (valoareBP1, valoareBP2);
        }

        /// <summary>
        /// Alocă bursele studenților, respectând regulile de diferență între medii.
        /// </summary>
        private (decimal, List<(string Emplid, BursaIstoric Istoric)>) AssignScholarships(
    List<StudentRecord> students,
    decimal sumaDisponibila,
    decimal valoareAnualBP1,
    decimal valoareAnualBP2,
    decimal epsilon,
    FondBurseMeritRepartizat fondBurseMeritRepartizat,
    string etapa)
        {
            var istoricList = new List<(string Emplid, BursaIstoric Istoric)>();
            decimal? primaMedie = students.FirstOrDefault()?.Media;
            bool aFostAcordatBP2 = false;
            StudentRecord studentAnterior = null;

            foreach (var student in students)
            {
                //decimal diferenta = primaMedie.HasValue ? Math.Abs(primaMedie.Value - student.Media) : 0;
                decimal diferenta = studentAnterior != null ? Math.Abs(studentAnterior.Media - student.Media) : 0;

                string bursaAtribuita = null;
                decimal suma = 0;
                string motiv = "";
                string explicatie = "";
                string fallback = "";

                if (sumaDisponibila <= 0)
                {
                    student.Bursa = null;
                    continue;
                }
                bool eligibilPentruBP1 = ("licenta".Equals(fondBurseMeritRepartizat.programStudiu) && student.Media >= 9.00M)
                       || ("master".Equals(fondBurseMeritRepartizat.programStudiu) && student.Media >= 9.5M);

                if (eligibilPentruBP1)
                {
                    if (!aFostAcordatBP2 && diferenta <= epsilon)
                    {
                        if (sumaDisponibila >= valoareAnualBP1)
                        {
                            bursaAtribuita = "BP1";
                            suma = valoareAnualBP1;
                            motiv = "Media ≥ 9.00 și Δ ≤ ε – BP1 acordat";
                        }
                        else if (sumaDisponibila >= valoareAnualBP2)
                        {
                            bursaAtribuita = "BP2";
                            suma = valoareAnualBP2;
                            motiv = "Fond insuficient pentru BP1 – fallback la BP2";
                            fallback = $"(necesar BP1: {valoareAnualBP1:F2} lei, dar disponibil doar: {sumaDisponibila:F2} lei)";
                            aFostAcordatBP2 = true;
                        }
                    }
                    else if (sumaDisponibila >= valoareAnualBP2)
                    {
                        bursaAtribuita = "BP2";
                        suma = valoareAnualBP2;
                        motiv = "Δ > ε – fallback la BP2";
                        fallback = $"(Δ = {diferenta:F2} > ε = {epsilon:F2})";
                        aFostAcordatBP2 = true;
                    }
                }
                else if (sumaDisponibila >= valoareAnualBP2 && student.Media >= 8.00M)
                {
                    bursaAtribuita = "BP2";
                    suma = valoareAnualBP2;
                    motiv = "Media < 9.00 – BP2 acordat";
                    fallback = "(criteriu media)";
                    aFostAcordatBP2 = true;
                }

                if (primaMedie == null)
                {
                    explicatie = "Primul student – fără comparație anterioară";
                }
                else
                {
                    explicatie = $"Media primului student: {primaMedie:F2} → Δ = {diferenta:F2} {(diferenta <= epsilon ? "(Δ ≤ ε)" : "(Δ > ε)")}";
                }

                /*if (!string.IsNullOrEmpty(bursaAtribuita))
                {
                    student.Bursa = bursaAtribuita;
                    student.SumaBursa = suma;
                    sumaDisponibila -= suma;

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
                        DataModificare = DateTime.Now
                    }));
                }
                else
                {
                    student.Bursa = null;
                }*/
                if (!string.IsNullOrEmpty(bursaAtribuita))
                {
                    student.Bursa = bursaAtribuita;
                    student.SumaBursa = suma;
                    sumaDisponibila -= suma;

                    // Construim Comentariu AI
                    string anterior = studentAnterior != null
                        ? $"Studentul anterior: {studentAnterior.NumeStudent} (media {studentAnterior.Media:F2}, bursă {studentAnterior.Bursa})"
                        : "Acesta este primul student care primește bursă.";

                    var urmatorii = students
                        .Where(s => s.Bursa == null && s != student)
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
                        ComentariiAI = comentariuAI, // NOU
                        DataModificare = DateTime.Now
                    }));

                    studentAnterior = student; // actualizăm pentru următorul student
                }
                else
                {
                    student.Bursa = null;
                }

            }

            return (sumaDisponibila, istoricList);
        }
        private (decimal, List<(string Emplid, BursaIstoric Istoric)>) AssignScholarshipsOptimezedWithCriteriaMediilorAceleasi(
    List<StudentRecord> students,
    decimal sumaDisponibila,
    decimal valoareAnualBP1,
    decimal valoareAnualBP2,
    decimal epsilon,
    FondBurseMeritRepartizat fondBurseMeritRepartizat,
    string etapa)
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
                decimal diferenta = studentAnterior != null ? Math.Abs(studentAnterior.Media - student.Media) : 0;
                // The 'diferenta' calculation will now reflect the media of the previously processed student
                // in the *sorted* list. If two students had the same primary Media and were ordered by tie-breakers,
                // their 'diferenta' will be 0, correctly triggering the logic for close averages.

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

                if (eligibilPentruBP1)
                {
                    if (!aFostAcordatBP2 && diferenta <= epsilon)
                    {
                        if (sumaDisponibila >= valoareAnualBP1)
                        {
                            bursaAtribuita = "BP1";
                            suma = valoareAnualBP1;
                            motiv = "Media ≥ 9.00 și Δ ≤ ε – BP1 acordat";
                        }
                        else if (sumaDisponibila >= valoareAnualBP2)
                        {
                            bursaAtribuita = "BP2";
                            suma = valoareAnualBP2;
                            motiv = "Fond insuficient pentru BP1 – fallback la BP2";
                            fallback = $"(necesar BP1: {valoareAnualBP1:F2} lei, dar disponibil doar: {sumaDisponibila:F2} lei)";
                            aFostAcordatBP2 = true;
                        }
                    }
                    else if (sumaDisponibila >= valoareAnualBP2)
                    {
                        bursaAtribuita = "BP2";
                        suma = valoareAnualBP2;
                        motiv = "Δ > ε – fallback la BP2";
                        fallback = $"(Δ = {diferenta:F2} > ε = {epsilon:F2})";
                        aFostAcordatBP2 = true;
                    }
                }
                else if (sumaDisponibila >= valoareAnualBP2 && student.Media >= 8.00M)
                {
                    bursaAtribuita = "BP2";
                    suma = valoareAnualBP2;
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
                        .Where(s => s.Bursa == null && s != student) // Ensure 's != student' is used for the current iteration
                        .Take(5)
                        .Select(s => $"(Emplid: {s.Emplid}, Media: {s.Media:F2}, An: {s.An}, Bursa: {s.Bursa ?? "—"})")
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
                        DataModificare = DateTime.Now
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


        private (decimal, List<(string Emplid, BursaIstoric Istoric)>) AssignScholarshipsOptimized(
    List<StudentRecord> students,
    decimal sumaDisponibila,
    decimal valoareAnualBP1,
    decimal valoareAnualBP2,
    decimal epsilon,
    FondBurseMeritRepartizat fondBurseMeritRepartizat,
    string etapa)
        {
            var istoricList = new List<(string Emplid, BursaIstoric Istoric)>();
            decimal sumaDisponibilaInitiala = sumaDisponibila; // Păstrăm suma inițială pentru referință

            // 1. Pre-procesarea și categorizarea studenților
            // Sortăm studenții o singură dată la început pentru a asigura ordinea mediilor
            var sortedStudents = students.OrderByDescending(s => s.Media).ToList();

            var eligibleBP1Strict = new List<StudentRecord>(); // Studenți eligibili pentru BP1 cu diferență mică (sau primul)
            var eligibleBP1FallbackToBP2 = new List<StudentRecord>(); // Studenți eligibili pentru BP1, dar care ar primi BP2 dacă diferența e mare
            var eligibleOnlyBP2 = new List<StudentRecord>(); // Studenți eligibili doar pentru BP2 (media < 9.00/9.50, dar >= 8.00)

            StudentRecord previousStudentForEpsilonCheck = null; // Folosit pentru calculul diferenței epsilon

            foreach (var student in sortedStudents)
            {
                // Regula de eligibilitate BP1
                bool isEligibleForBP1Criterion = ("licenta".Equals(fondBurseMeritRepartizat.programStudiu) && student.Media >= 9.00M) ||
                                                 ("master".Equals(fondBurseMeritRepartizat.programStudiu) && student.Media >= 9.5M);

                // Regula de eligibilitate BP2 (dacă nu e eligibil pentru BP1)
                bool isEligibleForBP2Criterion = student.Media >= 8.00M;

                if (isEligibleForBP1Criterion)
                {
                    decimal diferenta = previousStudentForEpsilonCheck != null ? Math.Abs(previousStudentForEpsilonCheck.Media - student.Media) : 0;

                    if (previousStudentForEpsilonCheck == null || diferenta <= epsilon)
                    {
                        eligibleBP1Strict.Add(student);
                    }
                    else
                    {
                        // Este eligibil pentru BP1 după medie, dar diferența e prea mare -> ar cădea pe BP2 în alocarea efectivă
                        eligibleBP1FallbackToBP2.Add(student);
                    }
                    previousStudentForEpsilonCheck = student; // Actualizăm pentru următorul student
                }
                else if (isEligibleForBP2Criterion)
                {
                    eligibleOnlyBP2.Add(student);
                }
                // Studenții sub 8.00 nu sunt incluși în nicio listă de eligibilitate merit
            }

            // Combinăm toți studenții care pot primi cel puțin BP2 (inclusiv cei care ar fi putut primi BP1, dar cu diferență mare)
            // Această listă este esențială pentru a calcula numărul total de BP2 care pot fi acordate
            var potentialBP2Recipients = eligibleBP1FallbackToBP2
                                            .Concat(eligibleOnlyBP2)
                                            .OrderByDescending(s => s.Media)
                                            .ToList();

            int maxBP1Possible = eligibleBP1Strict.Count;
            int maxBP2Possible = potentialBP2Recipients.Count; // BP2 este mai general, poate fi acordat și celor ce nu au primit BP1

            int bestNumBP1 = 0;
            int bestNumBP2 = 0;
            int maxTotalScholarships = -1; // Folosim -1 pentru a ne asigura că orice combinație validă va fi mai bună

            // 2. Simularea tuturor combinațiilor posibile
            // Iterăm de la numărul maxim posibil de BP1 în jos, pentru a prioritiza BP1 dacă fondurile permit
            for (int numBP1 = maxBP1Possible; numBP1 >= 0; numBP1--)
            {
                // Asigurăm că avem suficienți studenți eligibili pentru numBP1
                if (numBP1 > eligibleBP1Strict.Count) continue;

                decimal costBP1 = numBP1 * valoareAnualBP1;
                decimal remainingFundsAfterBP1 = sumaDisponibila - costBP1;

                if (remainingFundsAfterBP1 < 0) continue; // Nu avem fonduri suficiente pentru acest număr de BP1

                // Câte BP2 putem acorda cu fondurile rămase?
                // Prioritizăm studenții din `potentialBP2Recipients` care nu au primit deja BP1
                int numBP2 = 0;
                int currentBP2Count = 0;

                foreach (var student in potentialBP2Recipients)
                {
                    // Asigurăm că studentul nu a fost deja selectat pentru un BP1 în această simulare
                    // (e.g., studentul din eligibleBP1Strict dacă am fi avut o listă combinată inițial)
                    // Pentru simplitate aici, ne bazăm pe faptul că eligibleBP1Strict și potentialBP2Recipients sunt disjuncte
                    // pentru studenții eligibili la BP1 cu epsilon ok și cei care cad pe BP2.
                    // Dacă un student e în eligibleBP1Strict, el nu e în potentialBP2Recipients.
                    if (currentBP2Count < maxBP2Possible && remainingFundsAfterBP1 >= valoareAnualBP2)
                    {
                        numBP2++;
                        remainingFundsAfterBP1 -= valoareAnualBP2;
                        currentBP2Count++;
                    }
                    else
                    {
                        break; // Nu mai avem fonduri sau studenți pentru BP2
                    }
                }


                int currentTotalScholarships = numBP1 + numBP2;

                // 3. Alegerea celei mai bune combinații
                if (currentTotalScholarships > maxTotalScholarships)
                {
                    maxTotalScholarships = currentTotalScholarships;
                    bestNumBP1 = numBP1;
                    bestNumBP2 = numBP2;
                }
                else if (currentTotalScholarships == maxTotalScholarships)
                {
                    // Criteriu de departajare: preferăm mai multe BP1 dacă numărul total e același
                    if (numBP1 > bestNumBP1)
                    {
                        bestNumBP1 = numBP1;
                        bestNumBP2 = numBP2;
                    }
                }
            }

            // 4. Alocarea efectivă a burselor pe baza celei mai bune combinații
            // Resetăm suma disponibilă pentru alocarea reală
            sumaDisponibila = sumaDisponibilaInitiala;

            // Alocăm BP1
            int bp1AllocatedCount = 0;
            foreach (var student in eligibleBP1Strict.OrderByDescending(s => s.Media)) // Re-sortăm pentru siguranță
            {
                if (bp1AllocatedCount < bestNumBP1 && sumaDisponibila >= valoareAnualBP1)
                {
                    student.Bursa = "BP1";
                    student.SumaBursa = valoareAnualBP1;
                    sumaDisponibila -= valoareAnualBP1;
                    bp1AllocatedCount++;

                    string motiv = "Media eligibilă și diferență de medie în limita epsilon.";
                    string explicatie = $"Media: {student.Media:F2}, Diferență față de anterior: {(student == sortedStudents.First() ? "N/A (primul)" : Math.Abs(previousStudentForEpsilonCheck.Media - student.Media).ToString("F2"))} (ε={epsilon:F2})";
                    string comentariu = $"Etapa: {etapa} | Media: {student.Media:F2} | {motiv} | {explicatie} | Suma acordată: {valoareAnualBP1:F2} lei | Rămas fond: {sumaDisponibila:F2} lei";
                    string comentariuAI = GenerateAIComment(student, null, sumaDisponibila, "BP1", motiv, "", etapa, sortedStudents); // Adapt AI comment

                    istoricList.Add((student.Emplid, new BursaIstoric
                    {
                        StudentRecordId = student.Id,
                        TipBursa = "BP1",
                        Actiune = "Acordare",
                        Suma = valoareAnualBP1,
                        Motiv = motiv,
                        Comentarii = comentariu,
                        ComentariiAI = comentariuAI,
                        DataModificare = DateTime.Now
                    }));
                }
                else
                {
                    // Dacă nu a primit BP1, asigură-te că nu are bursa setată dacă a fost setată anterior
                    student.Bursa = null;
                    student.SumaBursa = 0;
                }
            }

            // Alocăm BP2 pentru studenții rămași eligibili (cei din potentialBP2Recipients și cei din eligibleBP1FallbackToBP2 care nu au primit BP1)
            int bp2AllocatedCount = 0;
            foreach (var student in sortedStudents.Where(s => s.Bursa == null).OrderByDescending(s => s.Media))
            {
                // Verificăm dacă studentul este eligibil pentru BP2 (dacă media >= 8.00)
                bool isEligibleForBP2Criterion = student.Media >= 8.00M;

                if (isEligibleForBP2Criterion && bp2AllocatedCount < bestNumBP2 && sumaDisponibila >= valoareAnualBP2)
                {
                    student.Bursa = "BP2";
                    student.SumaBursa = valoareAnualBP2;
                    sumaDisponibila -= valoareAnualBP2;
                    bp2AllocatedCount++;

                    string motiv = "Media eligibilă pentru BP2.";
                    string fallback = ""; // Nu mai avem fallback aici, e alocare directă de BP2
                    string explicatie = $"Media: {student.Media:F2}";
                    string comentariu = $"Etapa: {etapa} | Media: {student.Media:F2} | {motiv} {fallback} | {explicatie} | Suma acordată: {valoareAnualBP2:F2} lei | Rămas fond: {sumaDisponibila:F2} lei";
                    string comentariuAI = GenerateAIComment(student, null, sumaDisponibila, "BP2", motiv, fallback, etapa, sortedStudents); // Adapt AI comment

                    istoricList.Add((student.Emplid, new BursaIstoric
                    {
                        StudentRecordId = student.Id,
                        TipBursa = "BP2",
                        Actiune = "Acordare",
                        Suma = valoareAnualBP2,
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

        // Această metodă este un placeholder. Va trebui să adaptezi logica reală de generare a comentariului AI
        // în funcție de contextul alocării specifice și de studenții anteriori/următori.
        private string GenerateAIComment(StudentRecord currentStudent, StudentRecord previousStudent, decimal remainingFunds, string scholarshipType, string reason, string fallback, string etapa, List<StudentRecord> allStudents)
        {
            string previousStudentInfo = previousStudent != null
                ? $"Studentul anterior: {previousStudent.NumeStudent} (media {previousStudent.Media:F2}, bursă {previousStudent.Bursa ?? "N/A"})"
                : "Acesta este primul student care primește bursă sau primul din categoria sa.";

            var nextEligibleStudents = allStudents
                .Where(s => s.Bursa == null && s.Media >= 8.00M && s != currentStudent) // Considerăm toți studenții eligibili pentru o bursă merit
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
    string etapa)
        {
            var istoricList = new List<(string Emplid, BursaIstoric Istoric)>();
            StudentRecord studentAnterior = null;

            foreach (var student in students)
            {
                string domeniu = student.FondBurseMeritRepartizat?.domeniu;
                if (string.IsNullOrEmpty(domeniu))
                {
                    student.Bursa = null;
                    continue;
                }

                Match match = Regex.Match(domeniu, @"\((\d+)\)");
                if (!match.Success)
                {
                    student.Bursa = null;
                    continue;
                }

                int an = int.Parse(match.Groups[1].Value);
                string program = student.FondBurseMeritRepartizat?.programStudiu;
                int? fondId = student.FondBurseMeritRepartizatId;

                // Calculează valoare BP2
                decimal valoareBP2 = (an == 4 || (program == "master" && an == 2))
                    ? fonduri[1].ValoreaLunara * 9.35M
                    : fonduri[1].ValoreaLunara * 12;


                if (sumaDisponibila >= valoareBP2 && student.Media >= 8.00M)
                {
                    student.Bursa = "BP2";
                    student.SumaBursa = valoareBP2;
                    sumaDisponibila -= valoareBP2;
                    if (fondId.HasValue)
                        sumaRamasaPeFond[fondId.Value] -= valoareBP2;

                    string infoDurata = (an == 4 || (program == "master" && an == 2)) ? "9.35 luni" : "12 luni";

                    string comentariu = $"Etapa: {etapa} | Media: {student.Media:F2} | " +
                        $"Acordare BP2 ({valoareBP2:F2} lei) – fond ID {fondId?.ToString() ?? "—"} | " +
                        $"Necesari: {valoareBP2:F2} lei | " +
                        $"Program: {program}, An: {an}, Durată: {infoDurata}";

                    var urmatorii = students
                        .Where(s => s.Bursa == null && s.Id != student.Id)
                        .Take(5)
                        .Select(s => $"{s.Emplid} (Media: {s.Media:F2}, An: {an}, Domeniu: {s.FondBurseMeritRepartizat?.domeniu ?? "—"})")
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
                        Suma = valoareBP2,
                        Motiv = "Acordare BP2 – fonduri suficiente",
                        Comentarii = comentariu,
                        ComentariiAI = comentariuAI, 
                        DataModificare = DateTime.Now
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





        public static List<StudentRecord> EliminaStudentiNeeligibili(List<StudentRecord> students)
        {
            return students.Where(s => s.RO == 0 && s.TR == 0).ToList();
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

                // Fallback la 0 pentru celule goale
                double bm1_excel1 = GetNumericValue(ws1.Cell(currentRow, colBM1));
                double bm1_excel2 = GetNumericValue(ws2.Cell(currentRow, colBM1));

                double bm2_excel1 = GetNumericValue(ws1.Cell(currentRow, colBM2));
                double bm2_excel2 = GetNumericValue(ws2.Cell(currentRow, colBM2));

                wsOut.Cell(outputRow, 1).Value = domeniu;

                // BM1 (D)
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

                // BM2 (F)
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
            Console.WriteLine("✅ Fișierul cu comparația a fost generat corect (cu valori lipsă tratate ca 0).");
        }

        // Funcție helper pentru extragere numerică cu fallback
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
                    .GroupBy(s => new { s.FondBurseMeritRepartizatId, s.FondBurseMeritRepartizat.domeniu })
                    .Select(group => new StudentScholarshipData
                    {
                        FondBurseId = group.Key.FondBurseMeritRepartizatId,
                        Domeniu = group.Key.domeniu,
                        BP1Count = group.Count(s => s.Bursa?.ToLower().Contains("bp1") ?? false),
                        BP2Count = group.Count(s => s.Bursa?.ToLower().Contains("bp2") ?? false)
                    }).ToList();

                using var input = new FileStream(etapa0Path, FileMode.Open, FileAccess.Read);
                using var outputStream = new MemoryStream();

                var updatedStream = ExcelUpdater.UpdateScholarshipCounts(input, studentiClasificati0);
                updatedStream.Position = 0;

                await updatedStream.CopyToAsync(outputStream);

                byte[] finalFileBytes = outputStream.ToArray();

                return File(finalFileBytes,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            "SituatieStudenti_modificati.xlsx");
            }
            catch (Exception ex)
            {
                return BadRequest($"❌ Eroare la generarea fișierului: {ex.Message}");
            }
        }



    }
}
