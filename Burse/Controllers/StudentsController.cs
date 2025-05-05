using Burse.Data;
using Burse.Models.DTO;
using Burse.Services.Abstractions;

using Microsoft.AspNetCore.Mvc;

using System.Text.RegularExpressions;

namespace Burse.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class StudentsController : ControllerBase
    {
        private readonly IStudentService _studentService;

        public StudentsController(IStudentService studentService)
        {
            _studentService = studentService;
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
                    DataModificare = h.DataModificare
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
        [HttpPatch("{id}/bursa")]
        public async Task<IActionResult> UpdateBursa(int id, [FromBody] UpdateBursaDto dto)
        {
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

    }

}
