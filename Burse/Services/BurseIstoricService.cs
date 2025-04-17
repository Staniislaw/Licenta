using Burse.Data;
using Burse.Models;
using Burse.Services.Abstractions;

namespace Burse.Services
{
    public class BurseIstoricService : IBurseIstoricService
    {
        private readonly BurseDBContext _context;

        public BurseIstoricService(BurseDBContext context)
        {
            _context = context;
        }
        public async Task AddToIstoricAsync(StudentRecord student, string actiune, decimal suma, string motiv)
        {
            var istoric = new BursaIstoric
            {
                StudentRecordId = student.Id,
                TipBursa = student.Bursa,
                Motiv = motiv,
                Actiune = actiune,
                Suma = suma,
                Comentarii = $"Bursa asignată automat în funcție de medie: {student.Media}",
                DataModificare = DateTime.Now
            };

            _context.BursaIstoric.Add(istoric);
            await _context.SaveChangesAsync();
        }

    }
}
