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
        private readonly GrupuriDomeniiHelper _grupuriHelper;
        private readonly IBurseIstoricService _burseIstoricService;

        public FondBurseController(BurseDBContext context, IFondBurseService fondBurseService, IFondBurseMeritRepartizatService fondBurseMeritRepartizatService, GrupuriDomeniiHelper grupuriHelper, IBurseIstoricService burseIstoricService    )
        {
            _context = context;
            _fondBurseService = fondBurseService;
            _fondBurseMeritRepartizatService = fondBurseMeritRepartizatService;
            _grupuriHelper = grupuriHelper;
            _burseIstoricService = burseIstoricService;
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
        public async Task<IActionResult> ProcessExcelFiles([FromForm] List<IFormFile> pathStudentiList,IFormFile burseFile)
        {
            //await _fondBurseService.ResetSumaRamasaAsync();
            //await _fondBurseService.ResetStudentiAsync();
            decimal epsilon = 0.05M;

            if (burseFile == null)
            {
                return BadRequest("Fișierul Burse_Studenti.xlsx nu a fost găsit.");
            }
            var streamBurseFile = burseFile.OpenReadStream();

            StudentExcelReader excelReader = new StudentExcelReader();
            List<FondBurse> fonduri = await _fondBurseService.GetDateFromBursePerformanteAsync();

            foreach (var pathStudenti in pathStudentiList)
            {
                using var stream = pathStudenti.OpenReadStream();

                Dictionary<string, List<StudentRecord>> studentRecords = excelReader.ReadStudentRecordsFromExcel(stream, pathStudenti.FileName);
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
                    (var sumaRamasa, var istoricePerDomeniu) = AssignScholarships(
                        students,
                        sumaDisponibila,
                        valoareAnualBP1,
                        valoareAnualBP2,
                        epsilon,
                        "0"
                    );
                    sumaDisponibila = sumaRamasa;

                    istoricList.AddRange(istoricePerDomeniu);


                    students.ForEach(s => s.FondBurseMeritRepartizatId = fondRepartizatByDomeniu.ID);
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
                    Console.WriteLine("Eroare la salvarea în BursaIstoric:");
                    Console.WriteLine(ex.Message);

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

            string grupCastigatorNume = null;
            List<string> domeniiGrupCastigator = null;
            List<FondBurseMeritRepartizat> fonduriCastigatoare = null;
            decimal sumaDisponibilaMax = 0;
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
                if (sumaDisponibila <= 0) continue;
                sumaDisponibilaAdd += sumaDisponibila;
                if (sumaDisponibila > sumaDisponibilaMax)
                {
                    sumaDisponibilaMax = sumaDisponibila;
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
                    var sumaRamasa = fonduriRepartizate
                        .Where(f =>
                            GetDomeniiDinGrupa(f.Grupa)
                                .Any(d => grup.Value.Contains(d)))
                        .Sum(f => f.SumaRamasa);

                    return new
                    {
                        NumeGrup = grup.Key,
                        Domenii = grup.Value,
                        SumaRamasa = sumaRamasa
                    };
                })
                .OrderByDescending(g => g.SumaRamasa)
                .ToList();

            var grupCuSumaMaxima = grupuriCuSumaRamasa.FirstOrDefault();

            if (grupCuSumaMaxima != null && grupCuSumaMaxima.SumaRamasa > 0)
            {
                Console.WriteLine($"\n📌 PUNCTUL 4 – Grup cu cea mai mare sumă rămasă: {grupCuSumaMaxima.NumeGrup} ({grupCuSumaMaxima.SumaRamasa} lei)");

                var studentiGrup = await _fondBurseService.GetStudentiEligibiliPeDomeniiAsync(grupCuSumaMaxima.Domenii);

                if (studentiGrup.Any())
                {
                    var fonduriGrup = fonduriRepartizate
                        .Where(f =>
                            f.programStudiu == "licenta" &&
                            GetDomeniiDinGrupa(f.Grupa)
                                .Any(d => grupCuSumaMaxima.Domenii.Contains(d)))
                        .ToList();

                    var sumaRamasaPeFond = fonduriGrup.ToDictionary(f => f.ID, f => f.SumaRamasa);
                    decimal sumaDisponibila = grupCuSumaMaxima.SumaRamasa;

                    // 🔽 Sortează după medie descrescător
                    studentiGrup = studentiGrup
                        .OrderByDescending(s => s.Media)
                        .ToList();

                    //AssignOnlyBP2(studentiGrup, ref sumaDisponibila, fonduri, sumaRamasaPeFond);
                    (decimal sumaNoua, var istoricBP2) = AssignOnlyBP2(studentiGrup, sumaDisponibila, fonduri, sumaRamasaPeFond, "4");
                    sumaDisponibila = sumaNoua;

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
    string etapa)
        {
            var istoricList = new List<(string Emplid, BursaIstoric Istoric)>();
            decimal ultimaMedie = -1;
            bool aFostAcordatBP2 = false;

            foreach (var student in students)
            {
                decimal diferenta = ultimaMedie < 0 ? 0 : Math.Abs(ultimaMedie - student.Media);
                string bursaAtribuita = null;
                decimal suma = 0;
                string motiv = "";
                string explicatie = "";
                string fallback = "";

                if (sumaDisponibila <= 0)
                {
                    student.Bursa = null;
                    ultimaMedie = student.Media;
                    continue;
                }

                if (student.Media >= 9.00M)
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
                else if (sumaDisponibila >= valoareAnualBP2)
                {
                    bursaAtribuita = "BP2";
                    suma = valoareAnualBP2;
                    motiv = "Media < 9.00 – BP2 acordat";
                    fallback = "(criteriu media)";
                    aFostAcordatBP2 = true;
                }

                // Explicație detaliată pentru istoric
                if (ultimaMedie < 0)
                {
                    explicatie = "Primul student – fără comparație anterioară";
                }
                else
                {
                    explicatie = $"Media precedentă: {ultimaMedie:F2} → Δ = {diferenta:F2} {(diferenta <= epsilon ? "(Δ ≤ ε)" : "(Δ > ε)")}";
                }

                if (!string.IsNullOrEmpty(bursaAtribuita))
                {
                    student.Bursa = bursaAtribuita;
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
                }

                ultimaMedie = student.Media;
            }

            return (sumaDisponibila, istoricList);
        }



        private (decimal, List<(string Emplid, BursaIstoric Istoric)>) AssignOnlyBP2(
    List<StudentRecord> students,
    decimal sumaDisponibila,
    List<FondBurse> fonduri,
    Dictionary<int, decimal> sumaRamasaPeFond,
    string etapa)
        {
            var istoricList = new List<(string Emplid, BursaIstoric Istoric)>();

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


                if (sumaDisponibila >= valoareBP2)
                {
                    student.Bursa = "BP2";
                    sumaDisponibila -= valoareBP2;
                    if (fondId.HasValue)
                        sumaRamasaPeFond[fondId.Value] -= valoareBP2;

                    string infoDurata = (an == 4 || (program == "master" && an == 2)) ? "9.35 luni" : "12 luni";

                    string comentariu = $"Etapa: {etapa} | Media: {student.Media:F2} | " +
                        $"Acordare BP2 ({valoareBP2:F2} lei) – fond ID {fondId?.ToString() ?? "—"} | " +
                        $"Necesari: {valoareBP2:F2} lei | " +
                        $"Program: {program}, An: {an}, Durată: {infoDurata}";

                    istoricList.Add((student.Emplid, new BursaIstoric
                    {
                        StudentRecordId = student.Id,
                        TipBursa = "BP2",
                        Actiune = "Acordare",
                        Suma = valoareBP2,
                        Motiv = "Acordare BP2 – fonduri suficiente",
                        Comentarii = comentariu,
                        DataModificare = DateTime.Now
                    }));
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

    }
}
