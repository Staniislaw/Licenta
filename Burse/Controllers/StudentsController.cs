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
                An = student.An+1,
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
            }).ToList();

            return Ok(studentDtos);
        }

    }

}
