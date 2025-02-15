using Burse.Data;
using Burse.Models;
using Burse.Services.Abstractions;

using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace Burse.Services
{
    public class FondBurseService : IFondBurseService
    {
        private readonly BurseDBContext _context;

        public FondBurseService(BurseDBContext context)
        {
            _context = context;
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
    }
}
