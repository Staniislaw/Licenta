using Burse.Data;
using Burse.Helpers;
using Burse.Models;
using Burse.Models.DTO;
using Burse.Services;
using Burse.Services.Abstractions;

using ClosedXML.Excel;

using Microsoft.AspNetCore.Mvc;

using System.Text.RegularExpressions;

namespace Burse.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class StudentsController : ControllerBase
    {
        private readonly IStudentService _studentService;
        private readonly GrupuriDomeniiHelper _helper;
        private readonly IFondBurseMeritRepartizatService _fondBurseMeritRepartizatService;
        private readonly AppLogger _logger;
        private readonly IFondBurseService _fondBurseService;
        public StudentsController(IStudentService studentService, GrupuriDomeniiHelper helper, IFondBurseMeritRepartizatService fondBurseMeritRepartizatService, AppLogger logger, IFondBurseService fondBurseService)
        {
            _studentService = studentService;
            _helper = helper;
            _fondBurseMeritRepartizatService = fondBurseMeritRepartizatService;
            _logger = logger;
            _fondBurseService = fondBurseService;
        }

        [HttpGet("getStudents")]
        public async Task<IActionResult> GetStudents()
        {
            var students = await _studentService.GetAllAsync();

            // Mapare manuală de la entitate la DTO
            var studentDtos = students.Select(student => new StudentDto
            {
                Id = student.Id,
                Emplid = student.Emplid,
                CNP = student.CNP,
                NumeStudent = student.NumeStudent,
                An = student.An + 1,
                Media = student.Media,
                PunctajAn = student.PunctajAn,
                CO = student.CO,
                RO = student.RO,
                TC = student.TC,
                TR = student.TR,
                Bursa = student.Bursa,
                Domeniu = Regex.Replace(student.FondBurseMeritRepartizat.domeniu, @"\s*\(\d+\)", ""),
                ProgramStudiu = student.FondBurseMeritRepartizat.programStudiu,
                Grupa = student.FondBurseMeritRepartizat.Grupa,

                IstoricBursa = student.IstoricBursa.Select(h => new BursaIstoricDto
                {
                    TipBursa = h.TipBursa,
                    Motiv = h.Motiv,
                    Actiune = h.Actiune,
                    Suma = h.Suma,
                    Comentarii = h.Comentarii,
                    DataModificare = h.DataModificare,
                    ComentariiAI = h.ComentariiAI
                }).ToList()
            }).ToList();


            return Ok(studentDtos);
        }
        /// <summary>
        /// Actualizează doar bursa unui student.
        /// </summary>
        /// 
        public class UpdateBursaDto
        {
            public string Bursa { get; set; }
        }
        [HttpGet("{id}/can-change-bursa")]
        public async Task<ActionResult<BursaChangeResponse>> CanChangeBursa(int id, [FromQuery] string bursaNoua)
        {
            List<FondBurse> fonduri = await _fondBurseService.GetDateFromBursePerformanteAsync();

            var student = await _studentService.GetByIdAsync(id);

            if (student == null)
            {
                _logger.LogStudentInfo($"Studenti cu ID {id} nu exista in baza de date");
                return NotFound();
            }

            (decimal valoareAnualBP1, decimal valoareAnualBP2) = CalculateScholarshipValues(student.FondBurseMeritRepartizat.domeniu, fonduri, student.FondBurseMeritRepartizat);

            string bursaVeche = student?.Bursa ?? string.Empty;

            if (bursaVeche == bursaNoua)
            {
                return new BursaChangeResponse
                {
                    CanChange = false,
                    Message = "Studentul deja are această bursă."
                };
            }

            var fonduriRepartizate = await _fondBurseMeritRepartizatService.GetAllAsync();
            var sumaTotalaRamasa = fonduriRepartizate.Sum(f => f.SumaRamasa);
            decimal diferenta = valoareAnualBP1 - valoareAnualBP2;

            if (bursaVeche == "BP2" && bursaNoua == "BP1")
            {
                if (sumaTotalaRamasa < diferenta)
                {
                    return new BursaChangeResponse
                    {
                        CanChange = false,
                        Message = $"Fondurile rămase ({sumaTotalaRamasa}) sunt insuficiente pentru modificarea burselor (diferența necesară: {diferenta})."
                    };
                }
         
            }
            else if (bursaVeche == "BP1" && bursaNoua == "BP2")
            {
              
            }
            else if (string.IsNullOrEmpty(bursaVeche) && bursaNoua == "BP1")
            {
                if (sumaTotalaRamasa < valoareAnualBP1)
                {
                    return new BursaChangeResponse
                    {
                        CanChange = false,
                        Message = $"Fondurile rămase ({sumaTotalaRamasa}) sunt insuficiente pentru modificarea burselor (diferența necesară: {valoareAnualBP1})."
                    };
                }
               
            }
            else if (string.IsNullOrEmpty(bursaVeche) && bursaNoua == "BP2")
            {
                if (sumaTotalaRamasa < valoareAnualBP2)
                {
                    return new BursaChangeResponse
                    {
                        CanChange = false,
                        Message = $"Fondurile rămase ({sumaTotalaRamasa}) sunt insuficiente pentru modificarea burselor (diferența necesară: {valoareAnualBP2})."
                    };
                }
               
            }
            else if (bursaVeche == "BP1" && string.IsNullOrEmpty(bursaNoua))
            {
               
            }
            else if (bursaVeche == "BP2" && string.IsNullOrEmpty(bursaNoua))
            {
             
            }

           
            return new BursaChangeResponse
            {
                CanChange = true,
                Message = "Modificarea este posibilă."
            };

        }

        [HttpPatch("{id}/bursa")]
        public async Task<IActionResult> UpdateBursa(int id, [FromBody] UpdateBursaDto dto)
        {
            
            List<FondBurse> fonduri = await _fondBurseService.GetDateFromBursePerformanteAsync();

            var student = await _studentService.GetByIdAsync(id);

            if (student == null)
            {
                _logger.LogStudentInfo($"Studenti cu ID {id} nu exista in baza de date");
                return NotFound();
            }

            (decimal valoareAnualBP1, decimal valoareAnualBP2) = CalculateScholarshipValues(student.FondBurseMeritRepartizat.domeniu, fonduri, student.FondBurseMeritRepartizat);

            string bursaVeche = student?.Bursa ?? string.Empty;

            var fonduriRepartizate = await _fondBurseMeritRepartizatService.GetAllAsync();
            var sumaTotalaRamasa = fonduriRepartizate.Sum(f => f.SumaRamasa);
            decimal diferenta = valoareAnualBP1 - valoareAnualBP2;

            decimal sumaDeModificat = 0;
            bool scadeFond = true;

            if (bursaVeche == "BP2" && dto.Bursa == "BP1")
            {
                
                sumaDeModificat = diferenta;
                scadeFond = true;
            }
            else if (bursaVeche == "BP1" && dto.Bursa == "BP2")
            {
                sumaDeModificat = diferenta;
                scadeFond = false;
            }
            else if (string.IsNullOrEmpty(bursaVeche) && dto.Bursa == "BP1")
            {
                
                sumaDeModificat = valoareAnualBP1;
                scadeFond = true;
            }
            else if (string.IsNullOrEmpty(bursaVeche) && dto.Bursa == "BP2")
            {
                sumaDeModificat = valoareAnualBP2;
                scadeFond = true;
            }
            else if (bursaVeche == "BP1" && string.IsNullOrEmpty(dto.Bursa))
            {
                sumaDeModificat = valoareAnualBP1;  // corectat din BP2 la BP1
                scadeFond = false;
            }
            else if (bursaVeche == "BP2" && string.IsNullOrEmpty(dto.Bursa))
            {
                sumaDeModificat = valoareAnualBP2;
                scadeFond = false;
            }


            await _fondBurseMeritRepartizatService.UpdateFondAsync(student.FondBurseMeritRepartizatId, sumaDeModificat, scadeFond);
            string bursaVecheText = string.IsNullOrEmpty(bursaVeche) ? "fără bursă" : bursaVeche;
            string bursaNouaText = string.IsNullOrEmpty(dto.Bursa) ? "fără bursă" : dto.Bursa;

            _logger.LogModificareBursa($"Studentul cu emplid {student.Emplid} si numele {student.NumeStudent} a fost modificat bursa din {bursaVecheText} în {bursaNouaText}");



            var updatedEntity = await _studentService.UpdateBursaAsync(id, dto.Bursa);
            if (updatedEntity == null)
                return NotFound();
            // Mapăm entitatea la DTO-ul de răspuns
            var updatedDto = new StudentDto
            {
                Id = updatedEntity.Id,
                Emplid = updatedEntity.Emplid,
                CNP = updatedEntity.CNP,
                NumeStudent = updatedEntity.NumeStudent,
                An = updatedEntity.An + 1,
                Media = updatedEntity.Media,
                PunctajAn = updatedEntity.PunctajAn,
                CO = updatedEntity.CO,
                RO = updatedEntity.RO,
                TC = updatedEntity.TC,
                TR = updatedEntity.TR,
                Bursa = updatedEntity.Bursa,
                Domeniu = Regex.Replace(updatedEntity.FondBurseMeritRepartizat.domeniu, @"\s*\(\d+\)", ""),
                ProgramStudiu = updatedEntity.FondBurseMeritRepartizat.programStudiu,
                Grupa = updatedEntity.FondBurseMeritRepartizat.Grupa,
                IstoricBursa = updatedEntity.IstoricBursa.Select(h => new BursaIstoricDto
                {
                    TipBursa = h.TipBursa,
                    Motiv = h.Motiv,
                    Actiune = h.Actiune,
                    Suma = h.Suma,
                    Comentarii = h.Comentarii,
                    DataModificare = h.DataModificare
                }).ToList()
            };

            return Ok(updatedDto);
        }

        [HttpGet("program-studiu-options")]
        public async Task<IActionResult> GetProgramStudiuOptions()
        {
            var grupuri = await _helper.GetGrupuriProgramStudiiAsync();

            var domenii = grupuri
                .SelectMany(g => g.Value)
                .Distinct()
                .OrderBy(d => d)
                .ToList();

            return Ok(domenii);
        }
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

    }

}
