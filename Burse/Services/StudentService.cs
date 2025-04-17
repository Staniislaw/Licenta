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
                .Where(s => !string.IsNullOrWhiteSpace(s.Bursa))
                .ToListAsync();
        }
    }
}
