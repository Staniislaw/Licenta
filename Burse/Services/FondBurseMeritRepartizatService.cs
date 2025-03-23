using Burse.Data;
using Burse.Models;
using Burse.Services.Abstractions;

using Microsoft.EntityFrameworkCore;

namespace Burse.Services
{
    public class FondBurseMeritRepartizatService : IFondBurseMeritRepartizatService
    {
        private readonly BurseDBContext _context;

        public FondBurseMeritRepartizatService(BurseDBContext context)
        {
            _context = context;
        }

        // ✅ Get all records
        public async Task<List<FondBurseMeritRepartizat>> GetAllAsync()
        {
            return await _context.FondBurseMeritRepartizat.ToListAsync();
        }
       

        // ✅ Get a single record by domain (domeniu)
        public async Task<FondBurseMeritRepartizat> GetByDomeniuAsync(string domeniu)
        {
            return await _context.FondBurseMeritRepartizat
                                 .FirstOrDefaultAsync(f => f.domeniu == domeniu);
        }

        // ✅ UPDATE fonduri ramase
        public async Task UpdateAsync(FondBurseMeritRepartizat fond)
        {
            _context.FondBurseMeritRepartizat.Update(fond);
            await _context.SaveChangesAsync();
        }


        // ✅ Add a new record
        public async Task<bool> AddAsync(FondBurseMeritRepartizat newFond)
        {
            // Check if a record with the same 'domeniu' already exists
            var existingFond = await _context.FondBurseMeritRepartizat
                .FirstOrDefaultAsync(f => f.domeniu == newFond.domeniu);

            if (existingFond != null)
            {
                // ✅ Update existing record
                existingFond.bursaAlocatata = newFond.bursaAlocatata;
                existingFond.Grupa = newFond.Grupa;
                _context.FondBurseMeritRepartizat.Update(existingFond);
            }
            else
            {
                // ✅ Insert new record
                _context.FondBurseMeritRepartizat.Add(newFond);
            }

            await _context.SaveChangesAsync();
            return true;
        }


        // ✅ Update an existing record
        public async Task<bool> UpdateAsync(string domeniu, decimal newAmount)
        {
            var fond = await _context.FondBurseMeritRepartizat
                                     .FirstOrDefaultAsync(f => f.domeniu == domeniu);
            if (fond == null)
            {
                return false; // Not found
            }

            fond.bursaAlocatata = newAmount;
            await _context.SaveChangesAsync();
            return true;
        }

        // ✅ Delete a record
        public async Task<bool> DeleteAsync(string domeniu)
        {
            var fond = await _context.FondBurseMeritRepartizat
                                     .FirstOrDefaultAsync(f => f.domeniu == domeniu);
            if (fond == null)
            {
                return false; // Not found
            }

            _context.FondBurseMeritRepartizat.Remove(fond);
            await _context.SaveChangesAsync();
            return true;
        }
    }

}
