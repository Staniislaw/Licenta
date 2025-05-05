using Burse.Data;
using Burse.Models;
using Burse.Services.Abstractions;

using Microsoft.EntityFrameworkCore;

namespace Burse.Services
{
    public class StudentService : IStudentService
    {
        private readonly BurseDBContext _context;

        public StudentService(BurseDBContext context)
        {
            _context = context;
        }

        public async Task<List<StudentRecord>> GetAllAsync()
        {
            return await _context.StudentRecord
                .Include(s => s.FondBurseMeritRepartizat)
                 .Include(s => s.IstoricBursa)
                .ToListAsync();
        }
        public async Task<StudentRecord> UpdateBursaAsync(int id, string bursa)
        {
            var student = await _context.StudentRecord
                .Include(s => s.FondBurseMeritRepartizat)
                .Include(s => s.IstoricBursa)
                .FirstOrDefaultAsync(s => s.Id == id);

            if (student == null)
                return null;

            var actiune = "Modificare bursa";
            if(string.IsNullOrEmpty(bursa))
            {
                bursa = null;
            }
            student.Bursa = bursa;

            var historyEntry = new BursaIstoric
            {
                StudentRecordId = student.Id,
                TipBursa = bursa,
                Motiv = $"Schimbare bursă prin interfață",
                Actiune = actiune,
                Etapa = null,               
                Suma = 0m,              
                Comentarii = "Modificat din interfață",
                DataModificare = DateTime.UtcNow
            };
            student.IstoricBursa.Add(historyEntry);

            await _context.SaveChangesAsync();

            return student;
        }



    }
}
