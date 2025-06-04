using Burse.Data;
using Burse.Models;

using Microsoft.EntityFrameworkCore;

using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

public class GrupuriService : IGrupuriService
{
    private readonly BurseDBContext _context;

    public GrupuriService(BurseDBContext context)
    {
        _context = context;
    }

    // Grupuri Bursa

    public async Task<bool> AddDomeniuToGrupBursaAsync(GrupBursaEntry payload)
    {
        var exists = await _context.GrupBursa.AnyAsync(e => e.GrupBursa == payload.GrupBursa && e.Domeniu == payload.Domeniu);
        if (!exists)
        {
            _context.GrupBursa.Add(payload);
            await _context.SaveChangesAsync();
            return true;
        }
        return false;
    }

    public async Task RemoveDomeniuFromGrupBursaAsync(string grup, string domeniu)
    {
        var entry = await _context.GrupBursa.FirstOrDefaultAsync(e => e.GrupBursa == grup && e.Domeniu == domeniu);
        if (entry != null)
        {
            _context.GrupBursa.Remove(entry);
            await _context.SaveChangesAsync();
        }
    }

    public async Task<Dictionary<string, List<string>>> GetGrupuriBurseAsync()
    {
        var entries = await _context.GrupBursa.ToListAsync();
        return entries
            .GroupBy(e => e.GrupBursa)
            .ToDictionary(g => g.Key, g => g.Select(x => x.Domeniu).ToList());
    }

    // Grupuri Domeniu

    public async Task<Dictionary<string, List<string>>> GetGrupuriAsync()
    {
        var entries = await _context.GrupDomeniu.ToListAsync();
        return entries
            .GroupBy(e => e.Grup)
            .ToDictionary(g => g.Key, g => g.Select(x => x.Domeniu).ToList());
    }

    public async Task<List<string>> GetGrupuriAsyncByGrup(string grup)
    {
        var grupLower = grup.ToLower();

        var entries = await _context.GrupDomeniu
            .Where(e => e.Grup.ToLower() == grupLower)
            .ToListAsync();

        return entries.Select(e => e.Domeniu).ToList();
    }



    public async Task<bool> AddDomeniuToGrupAsync(GrupDomeniuEntry payload)
    {
        var exists = await _context.GrupDomeniu.AnyAsync(e => e.Grup == payload.Grup && e.Domeniu == payload.Domeniu);
        if (!exists)
        {
            _context.GrupDomeniu.Add(payload);
            await _context.SaveChangesAsync();
            return true;
        }
        return false;
    }

    public async Task RemoveDomeniuFromGrupAsync(string grup, string domeniu)
    {
        var entry = await _context.GrupDomeniu.FirstOrDefaultAsync(e => e.Grup == grup && e.Domeniu == domeniu);
        if (entry != null)
        {
            _context.GrupDomeniu.Remove(entry);
            await _context.SaveChangesAsync();
        }
    }

    // Grupuri Program Studii

    public async Task<bool> AddDomeniuToGrupProgramStudiiAsync(GrupProgramStudiiEntry payload)
    {
        var exists = await _context.GrupProgramStudii.AnyAsync(e => e.Grup == payload.Grup && e.Domeniu == payload.Domeniu);
        if (!exists)
        {
            _context.GrupProgramStudii.Add(payload);
            await _context.SaveChangesAsync();
            return true;
        }
        return false;
    }

    public async Task RemoveDomeniuFromGrupProgramStudiiAsync(string grup, string domeniu)
    {
        var entry = await _context.GrupProgramStudii.FirstOrDefaultAsync(e => e.Grup == grup && e.Domeniu == domeniu);
        if (entry != null)
        {
            _context.GrupProgramStudii.Remove(entry);
            await _context.SaveChangesAsync();
        }
    }

    public async Task<Dictionary<string, List<string>>> GetGrupuriProgramStudiiAsync()
    {
        var entries = await _context.GrupProgramStudii.ToListAsync();
        return entries
            .GroupBy(e => e.Grup)
            .ToDictionary(g => g.Key, g => g.Select(x => x.Domeniu).ToList());
    }

    // Grupuri PDF

    public async Task<Dictionary<string, List<string>>> GetGrupuriPdfAsync()
    {
        var entries = await _context.GrupPDF.ToListAsync();
        return entries
            .GroupBy(e => e.Grup)
            .ToDictionary(g => g.Key, g => g.Select(x => x.Valoare).ToList());
    }

    public async Task<bool> AddValToPdfGroupAsync(GrupPdfEntry payload)
    {
        var exists = await _context.GrupPDF.AnyAsync(e => e.Grup == payload.Grup && e.Valoare == payload.Valoare);
        if (!exists)
        {
            _context.GrupPDF.Add(payload);
            await _context.SaveChangesAsync();
            return true;
        }
        return false;
    }

    public async Task RemoveValFromPdfGroupAsync(string grup, string valoare)
    {
        var entry = await _context.GrupPDF.FirstOrDefaultAsync(e => e.Grup == grup && e.Valoare == valoare);
        if (entry != null)
        {
            _context.GrupPDF.Remove(entry);
            await _context.SaveChangesAsync();
        }
    }
    // Grupuri Acronime
    // Grupuri Acronime

    public async Task<Dictionary<string, List<string>>> GetGrupuriAcronimeAsync()
    {
        var entries = await _context.GrupAcronim.ToListAsync();
        return entries
            .GroupBy(e => e.Grup)
            .ToDictionary(g => g.Key, g => g.Select(x => x.Valoare).ToList());
    }

    public async Task<bool> AddValToAcronimGroupAsync(GrupAcronimEntry payload)
    {
        var exists = await _context.GrupAcronim.AnyAsync(e => e.Grup == payload.Grup && e.Valoare == payload.Valoare);
        if (!exists)
        {
            _context.GrupAcronim.Add(payload);
            await _context.SaveChangesAsync();
            return true;
        }
        return false;
    }

    public async Task RemoveValFromAcronimGroupAsync(string grup, string valoare)
    {
        var entry = await _context.GrupAcronim.FirstOrDefaultAsync(e => e.Grup == grup && e.Valoare == valoare);
        if (entry != null)
        {
            _context.GrupAcronim.Remove(entry);
            await _context.SaveChangesAsync();
        }
    }
}
